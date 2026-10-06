<#
.SYNOPSIS
    One-word steps for an MP migration with MPMigration.ps1 -- no long command
    lines to type on a locked-down management server.

.DESCRIPTION
    Stage this in a working folder on the TARGET management server, next to
    MPMigration.ps1 and your MigrationManifest.csv. Expected layout:

        <working folder>\
            Invoke-ScomMigrationStep.ps1, MPMigration.ps1, MigrationManifest.csv
            Source\   <- copied from the source MS after Export-ScomEnvironment.ps1 -Role Source
                AllMPs\, SealedOriginals\, SealedDependencies\, SourceInventory.csv,
                GroupConversion.csv, StaticGroupMembers.csv, OverrideInstances.csv, ...
            Target\AllMPs\   <- from Export-ScomEnvironment.ps1 -Role Target

    ExportSource and ExportTarget fill Source\ and Target\ for you, so the
    whole migration can be run from the target management server. ExportSource
    runs Export-ScomEnvironment.ps1 ON the source management server through
    PowerShell remoting (with the source's own OperationsManager module) and
    copies the result back. An existing Source\ or Target\ folder is kept as
    <folder>.previous-<timestamp>.

    Folder names and version labels can be changed with parameters or with an
    optional Migration.settings.psd1 next to this script, e.g.
        @{ SourceFolder = 'Source2012R2'; TargetFolder = 'Target2022'; SourceVersion = '2012 R2'; TargetVersion = '2022' }

    Steps (typical order):
        .\Invoke-ScomMigrationStep.ps1 ExportSource    # export the SOURCE from here through PowerShell remoting (needs -SourceServer)
        .\Invoke-ScomMigrationStep.ps1 ExportTarget    # export this (TARGET) management group
        .\Invoke-ScomMigrationStep.ps1 Check           # pre-flight, changes nothing
        .\Invoke-ScomMigrationStep.ps1 FixReview       # REVIEW groups -> exact member names; watcher groups -> display name
        .\Invoke-ScomMigrationStep.ps1 EnableOverrides # optional: Migrate = Y on every OVERRIDES row (backup kept)
        .\Invoke-ScomMigrationStep.ps1 MapInstances    # per-object overrides: find each source object in the target (read-only)
        .\Invoke-ScomMigrationStep.ps1 Compile         # builds candidates + Import-Batch script, imports nothing
        .\Invoke-ScomMigrationStep.ps1 WhyBlocked      # one line per BLOCKED MP, grouped by reason
        .\Invoke-ScomMigrationStep.ps1 OverrideReport  # overrides kept / removed and why
        .\Invoke-ScomMigrationStep.ps1 DryRun          # what WOULD import, changes nothing
        .\Invoke-ScomMigrationStep.ps1 Import          # real import (asks you to type YES)
        .\Invoke-ScomMigrationStep.ps1 Reimport        # same, but also overwrites unsealed MPs already present at the same version
        .\Invoke-ScomMigrationStep.ps1 TestGroups      # converted groups in the target: working / waiting for agents / broken
        .\Invoke-ScomMigrationStep.ps1 Rollback        # REMOVES every MP the last import added (asks you to type YES)
        .\Invoke-ScomMigrationStep.ps1 Collect         # zips the result files for review

    See docs/step-reference.md for details.
#>
param(
    [Parameter(Mandatory, Position = 0)][ValidateSet('ExportSource', 'ExportTarget', 'Check', 'FixReview', 'EnableOverrides', 'MapInstances', 'OverrideReport', 'Compile', 'DryRun', 'Import', 'Reimport', 'TestGroups', 'WhyBlocked', 'Rollback', 'Collect')][string]$Step,
    [string]$ManagementServer = $env:COMPUTERNAME,
    [string]$SourceFolder = 'Source',
    [string]$TargetFolder = 'Target',
    [string]$SourceVersion = '2016',
    [string]$TargetVersion = '2025',
    # ExportSource: the source (old) management server to export through PowerShell remoting.
    [string]$SourceServer,
    # ExportSource: folders or shares with original .mp/.mpb files, searched from THIS server.
    [string[]]$SealedSearchPath,
    # ExportSource: credential for the remoting session (default: your own account).
    [System.Management.Automation.PSCredential]$SourceCredential
)

$ErrorActionPreference = 'Stop'
$root    = Split-Path -Parent $MyInvocation.MyCommand.Path
# Optional per-site settings; explicit parameters win.
$settingsFile = Join-Path $root 'Migration.settings.psd1'
if (Test-Path -LiteralPath $settingsFile) {
    $cfg = Import-PowerShellDataFile -LiteralPath $settingsFile
    foreach ($k in 'ManagementServer', 'SourceFolder', 'TargetFolder', 'SourceVersion', 'TargetVersion', 'SourceServer', 'SealedSearchPath') {
        if ($cfg.ContainsKey($k) -and $cfg[$k] -and -not $PSBoundParameters.ContainsKey($k)) { Set-Variable -Name $k -Value $cfg[$k] }
    }
}
$src     = if ([System.IO.Path]::IsPathRooted($SourceFolder)) { $SourceFolder } else { Join-Path $root $SourceFolder }
$tgtBase = if ([System.IO.Path]::IsPathRooted($TargetFolder)) { $TargetFolder } else { Join-Path $root $TargetFolder }
$tgt     = Join-Path $tgtBase 'AllMPs'
$srcName = Split-Path -Leaf $src
$tgtName = Split-Path -Leaf $tgtBase
$importScriptName = "Import-Batch.$TargetVersion.ps1"
$compile = Join-Path $root 'Compile'

function Import-GroupConversion([string]$path) {
    # Accepts files from older exports that used the column names ExtraIn2016 / MatchesIn2016.
    $rows = @(Import-Csv -LiteralPath $path)
    foreach ($r in $rows) {
        foreach ($pair in @(@('ExtraIn2016', 'ExtraInSource'), @('MatchesIn2016', 'MatchesInSource'))) {
            $old = $r.PSObject.Properties[$pair[0]]
            if ($old -and -not $r.PSObject.Properties[$pair[1]]) {
                $r | Add-Member -NotePropertyName $pair[1] -NotePropertyValue $old.Value
                $r.PSObject.Properties.Remove($pair[0])
            }
        }
    }
    $rows
}

function Need([string]$p, [string]$what) {
    if (-not (Test-Path -LiteralPath $p)) { throw "Missing $what : $p" }
}

# Moves an existing export folder aside before a fresh export; returns the new
# path, or $null if there was nothing to keep.
function Backup-Folder([string]$p) {
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    if (@(Get-ChildItem -LiteralPath $p -Force).Count -eq 0) { Remove-Item -LiteralPath $p -Force; return $null }
    $dest = "$p.previous-$(Get-Date -Format 'yyyyMMdd-HHmmss')"
    Rename-Item -LiteralPath $p -NewName (Split-Path -Leaf $dest)
    Write-Host "Previous export kept as $dest" -ForegroundColor Yellow
    return $dest
}

switch ($Step) {

    'ExportSource' {
        Need (Join-Path $root 'Export-ScomEnvironment.ps1') 'Export-ScomEnvironment.ps1'
        Need (Join-Path $root 'MigrationManifest.csv') 'MigrationManifest.csv'
        if (-not $SourceServer) { throw "Name the source management server: -SourceServer <name>, or SourceServer = '<name>' in Migration.settings.psd1." }
        $search = @($SealedSearchPath | Where-Object { $_ })
        $prev = Backup-Folder $src
        # Originals dropped into the previous Source\SealedOriginals are picked up again.
        if ($prev -and (Test-Path -LiteralPath (Join-Path $prev 'SealedOriginals'))) { $search += (Join-Path $prev 'SealedOriginals') }
        $a = @{ Role = 'Source'; SourceServer = $SourceServer; Manifest = (Join-Path $root 'MigrationManifest.csv'); OutputFolder = $src }
        if ($search.Count -gt 0) { $a.SealedSearchPath = $search }
        if ($SourceCredential) { $a.Credential = $SourceCredential }
        & (Join-Path $root 'Export-ScomEnvironment.ps1') @a
        if ($prev) {
            Write-Host ""
            Write-Host "Re-exported. Edits you made to the previous $srcName\GroupConversion.csv are in $prev -- run FixReview again or copy your edits across, and re-run MapInstances." -ForegroundColor Yellow
        }
    }

    'ExportTarget' {
        Need (Join-Path $root 'Export-ScomEnvironment.ps1') 'Export-ScomEnvironment.ps1'
        $null = Backup-Folder $tgtBase
        & (Join-Path $root 'Export-ScomEnvironment.ps1') -Role Target -ManagementServer $ManagementServer -OutputFolder $tgtBase
    }

    'Check' {
        Need (Join-Path $root 'MPMigration.ps1') 'MPMigration.ps1'
        Need (Join-Path $root 'MigrationManifest.csv') 'MigrationManifest.csv'
        Need (Join-Path $src 'AllMPs') "source export ($srcName\AllMPs)"
        Need $tgt "target export ($tgtName\AllMPs)"

        # Full detail goes to a file you read on the server; the screen gets a
        # one-screenshot summary.
        $report = Join-Path $root 'Check-Report.txt'
        $lines = New-Object System.Collections.Generic.List[string]
        function Add-Line([string]$t) { $lines.Add($t) }

        $n16 = @(Get-ChildItem (Join-Path $src 'AllMPs') -Filter *.xml).Count
        $n25 = @(Get-ChildItem $tgt -Filter *.xml).Count
        $cnt = { param($d) $p = Join-Path $src $d; if (Test-Path $p) { @(Get-ChildItem $p -Include *.mp, *.mpb -Recurse).Count } else { 0 } }
        $nOrig = & $cnt 'SealedOriginals'; $nDep = & $cnt 'SealedDependencies'

        $nm = @(); $miss = @(); $rows = @(); $rev = @()
        $mr = Join-Path $src 'ManifestResolution.csv'
        if (Test-Path $mr) { $nm = @(Import-Csv $mr | Where-Object Status -eq 'NO MATCH') }
        $sm = Join-Path $src 'SealedOriginalsMissing.csv'
        if (Test-Path $sm) { $miss = @(Import-Csv $sm) }
        $gc = Join-Path $src 'GroupConversion.csv'
        if (Test-Path $gc) { $rows = @(Import-GroupConversion $gc); $rev = @($rows | Where-Object Recommendation -like 'REVIEW*') }

        Add-Line "CHECK REPORT  $(Get-Date)"
        Add-Line ""
        Add-Line "Source MPs exported: $n16   Target MPs exported: $n25   Sealed originals: $nOrig   Sealed dependencies: $nDep"
        Add-Line ""
        Add-Line "=== Manifest rows marked Y with NO MATCH in source ($($nm.Count)) -- fix MatchPattern in MigrationManifest.csv or ignore"
        $nm | ForEach-Object { Add-Line "   $($_.ManagementPack)" }
        Add-Line ""
        Add-Line "=== In-scope SEALED MPs with no original .mp/.mpb ($($miss.Count)) -- will be BLOCKED; drop the files in $srcName\SealedOriginals"
        $miss | ForEach-Object { Add-Line "   $($_.MPID)  v$($_.InstalledVersion)" }
        Add-Line ""
        Add-Line "=== Static group rules flagged REVIEW ($($rev.Count)) -- pattern also catches servers NOT in the group today"
        Add-Line "    Edit $srcName\GroupConversion.csv: tighten Pattern, or set Convert = N"
        foreach ($r in $rev) { Add-Line "   $($r.GroupClass)"; Add-Line "        pattern: $($r.Pattern)"; Add-Line "        members: $($r.MemberNames)"; Add-Line "        EXTRA  : $($r.ExtraInSource)" }
        Add-Line ""
        Add-Line "=== All static group rules ($($rows.Count))"
        foreach ($r in $rows) { Add-Line ("   [{0}] {1,-50} {2}  ({3})" -f $r.Convert, $r.GroupClass, $r.Pattern, $r.Recommendation) }
        $lines | Set-Content -LiteralPath $report -Encoding UTF8

        $groupsWithExtras = @($rev | Where-Object { $_.GroupClass -match 'Disable|NonProd|Non.Prod|Test|Maint|Suppress|Exclude' })
        Write-Host ""
        Write-Host "=========== CHECK SUMMARY ===========" -ForegroundColor Cyan
        Write-Host ("Source MPs: {0}   Target MPs: {1}   sealed originals: {2}   sealed deps: {3}" -f $n16, $n25, $nOrig, $nDep)
        Write-Host ("Manifest rows with NO MATCH        : {0}" -f $nm.Count) -ForegroundColor $(if ($nm.Count) { 'Yellow' } else { 'Green' })
        Write-Host ("Sealed MPs missing original file   : {0}" -f $miss.Count) -ForegroundColor $(if ($miss.Count) { 'Yellow' } else { 'Green' })
        Write-Host ("Static group rules                 : {0} total, {1} convert, {2} REVIEW" -f $rows.Count, @($rows | Where-Object Convert -eq 'Y').Count, $rev.Count) -ForegroundColor $(if ($rev.Count) { 'Yellow' } else { 'Green' })
        Write-Host ("  REVIEW groups with risky names   : {0}  (Disable/NonProd/Test/Maint...)" -f $groupsWithExtras.Count) -ForegroundColor $(if ($groupsWithExtras.Count) { 'Red' } else { 'Green' })
        foreach ($g in ($groupsWithExtras | Select-Object -First 5)) { Write-Host ("     {0}" -f $g.GroupClass) -ForegroundColor Red }
        Write-Host "Full detail: $report" -ForegroundColor Cyan
        Write-Host "====================================" -ForegroundColor Cyan
    }

    'FixReview' {
        # 1) Picks up a FRESH GroupConversion.csv from a new source export (earlier
        #    versions kept re-reading the first backup forever).
        # 2) Health Service Watcher members: matched on display name (the agent
        #    FQDN), because a watcher is not hosted on the server it watches.
        # 3) REVIEW groups -> exact current member names.
        $gc  = Join-Path $src 'GroupConversion.csv'
        Need $gc "$srcName\GroupConversion.csv"
        $bak = Join-Path $src 'GroupConversion.original.csv'
        $stamp = Join-Path $src 'GroupConversion.fixreview.stamp'
        $curHash = (Get-FileHash -LiteralPath $gc -Algorithm SHA256).Hash
        $lastWritten = if (Test-Path $stamp) { (Get-Content -LiteralPath $stamp -Raw).Trim() } else { '' }
        if (-not (Test-Path $bak) -or $curHash -ne $lastWritten) {
            Copy-Item -LiteralPath $gc -Destination $bak -Force
            Write-Host "Using the current GroupConversion.csv ($((Get-Item $gc).LastWriteTime)) as the source." -ForegroundColor Cyan
        }
        $rows = @(Import-GroupConversion $bak)

        # Member display names/classes from the source export.
        $members = @{}
        $sgm = Join-Path $src 'StaticGroupMembers.csv'
        if (Test-Path $sgm) {
            foreach ($m in @(Import-Csv -LiteralPath $sgm | Where-Object { $_.List -eq 'Include' -and $_.Resolved -eq 'True' })) {
                $k = "$($m.MP)|$($m.GroupClass)|$($m.RuleIndex)"
                if (-not $members.ContainsKey($k)) { $members[$k] = New-Object System.Collections.Generic.List[object] }
                $members[$k].Add($m)
            }
        }
        else { Write-Host "StaticGroupMembers.csv not found in $srcName -- watcher groups cannot be repaired." -ForegroundColor Yellow }

        $watcherFixed = 0; $fixed = 0
        foreach ($r in $rows) {
            $k = "$($r.MP)|$($r.GroupClass)|$($r.RuleIndex)"
            if ($members.ContainsKey($k)) {
                $ms = $members[$k].ToArray()
                $watchers = @($ms | Where-Object { $_.Class -like '*Watcher*' })
                if ($watchers.Count -gt 0 -and $watchers.Count -eq $ms.Count) {
                    $names = @($watchers | ForEach-Object { (([string]$_.DisplayName -split '\.')[0]).Trim().ToUpperInvariant() } |
                               Where-Object { $_ -match '^[A-Z0-9][A-Z0-9_-]{0,14}$' } | Sort-Object -Unique)
                    if ($names.Count -gt 0) {
                        $r.MemberKind = 'NameOnly'
                        $r.MemberNames = ($names -join '; ')
                        $r.Operator = 'MatchesRegularExpression'
                        $r.Pattern = '^(' + ((@($names | ForEach-Object { [regex]::Escape($_) })) -join '|') + ')$'
                        $r.Recommendation = 'CONVERT (Health Service Watchers, matched by display name)'
                        $r.ExtraInSource = ''
                        $r.Convert = 'Y'
                        $watcherFixed++
                        continue
                    }
                }
            }
            if ([string]$r.Recommendation -notlike 'REVIEW*' -or -not $r.MemberNames) { continue }
            $names = @(([string]$r.MemberNames) -split ';\s*' | Where-Object { $_ } | Sort-Object -Unique)
            $r.Operator = 'MatchesRegularExpression'
            $r.Pattern = '^(' + ((@($names | ForEach-Object { [regex]::Escape($_) })) -join '|') + ')$'
            $r.Recommendation = "CONVERT (exact names; the name pattern would also have caught: $($r.ExtraInSource))"
            $r.ExtraInSource = ''
            $r.Convert = 'Y'
            $fixed++
        }
        $rows | Export-Csv -LiteralPath $gc -NoTypeInformation -Encoding UTF8
        (Get-FileHash -LiteralPath $gc -Algorithm SHA256).Hash | Set-Content -LiteralPath $stamp
        $bad = @($rows | Where-Object { $_.Convert -eq 'Y' -and $_.Pattern -match '\^\((MICROSOFT|WINDOWS)[|)]' })
        Write-Host "$fixed REVIEW rule(s) -> exact names; $watcherFixed Health Service Watcher rule(s) -> display-name match." -ForegroundColor Green
        Write-Host "Rules still showing a bad name (MICROSOFT/WINDOWS): $($bad.Count)" -ForegroundColor $(if ($bad.Count) { 'Red' } else { 'Green' })
        foreach ($x in ($bad | Select-Object -First 8)) { Write-Host "   $($x.GroupClass)  [$($x.RuleClass)]" -ForegroundColor Red }
    }

    'EnableOverrides' {
        $man = Join-Path $root 'MigrationManifest.csv'
        Need $man 'MigrationManifest.csv'
        $bak = Join-Path $root 'MigrationManifest.before-overrides.csv'
        if (-not (Test-Path $bak)) { Copy-Item -LiteralPath $man -Destination $bak }
        $rows = @(Import-Csv -LiteralPath $man)
        $n = 0
        foreach ($r in $rows) { if ($r.Action -eq 'OVERRIDES' -and $r.Migrate -ne 'Y') { $r.Migrate = 'Y'; $n++ } }
        $rows | Export-Csv -LiteralPath $man -NoTypeInformation -Encoding UTF8
        Write-Host "$n override pack(s) switched to Migrate = Y ($(@($rows | Where-Object Migrate -eq 'Y').Count) packs now in scope). Backup: $bak" -ForegroundColor Green
        Write-Host "Packs for retired products (SQL/IIS/Windows 2003-2014 era) stay out -- their targets are unlikely to exist in the target."
    }

    'MapInstances' {
        # For each source object an override targets by ID, find the same object
        # in the target (same class + key => same FullName). Read-only against SCOM.
        $oi = Join-Path $src 'OverrideInstances.csv'
        Need $oi "$srcName\OverrideInstances.csv (from the source export)"
        Import-Module OperationsManager
        New-SCOMManagementGroupConnection -ComputerName $ManagementServer | Out-Null
        $rows = @(Import-Csv -LiteralPath $oi)
        $out = New-Object System.Collections.Generic.List[object]
        $byClass = $rows | Where-Object { $_.Resolved -eq 'True' -and $_.Class } | Group-Object Class
        $i = 0
        foreach ($grp in $byClass) {
            $i++
            Write-Progress -Activity "Finding source objects in SCOM $TargetVersion" -Status $grp.Name -PercentComplete ([int](100 * $i / @($byClass).Count))
            $full = @{}; $disp = @{}
            $cls = $null
            try { $cls = @(Get-SCOMClass -Name $grp.Name -ErrorAction Stop) | Select-Object -First 1 } catch { }
            if ($cls) {
                foreach ($x in @(Get-SCOMClassInstance -Class $cls -ErrorAction SilentlyContinue)) {
                    $full[([string]$x.FullName).ToLowerInvariant()] = [string]$x.Id
                    $dk = ([string]$x.DisplayName).ToLowerInvariant()
                    if ($disp.ContainsKey($dk)) { $disp[$dk] = '' } else { $disp[$dk] = [string]$x.Id }   # '' = ambiguous
                }
            }
            foreach ($r in $grp.Group) {
                $new = ''; $how = ''
                if (-not $cls) { $how = "Class not in target" }
                elseif ($full.ContainsKey(([string]$r.FullName).ToLowerInvariant())) { $new = $full[([string]$r.FullName).ToLowerInvariant()]; $how = 'FullName' }
                elseif ($disp[([string]$r.DisplayName).ToLowerInvariant()]) { $new = $disp[([string]$r.DisplayName).ToLowerInvariant()]; $how = 'DisplayName' }
                else { $how = "Not found in target" }
                $out.Add([PSCustomObject]@{ OldGuid = $r.Guid; NewGuid = $new; MatchedBy = $how; Class = $r.Class; FullName = $r.FullName; DisplayName = $r.DisplayName; UsedByMPs = $r.UsedByMPs })
            }
        }
        foreach ($r in @($rows | Where-Object { $_.Resolved -ne 'True' -or -not $_.Class })) {
            $out.Add([PSCustomObject]@{ OldGuid = $r.Guid; NewGuid = ''; MatchedBy = "Gone from source already"; Class = ''; FullName = ''; DisplayName = ''; UsedByMPs = $r.UsedByMPs })
        }
        Write-Progress -Activity "Finding source objects in SCOM $TargetVersion" -Completed
        $map = Join-Path $root 'InstanceMap.csv'
        $out | Export-Csv -LiteralPath $map -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "=========== INSTANCE MAP ($($out.Count) objects) ===========" -ForegroundColor Cyan
        $out | Group-Object MatchedBy | Sort-Object Count -Descending | ForEach-Object {
            $c = if ($_.Name -in 'FullName', 'DisplayName') { 'Green' } else { 'Yellow' }
            Write-Host ("{0,6}  {1}" -f $_.Count, $_.Name) -ForegroundColor $c
        }
        Write-Host "Written: $map  (Compile uses it automatically)" -ForegroundColor Cyan
    }

    'OverrideReport' {
        $man = Join-Path $compile 'BatchImportManifest.csv'
        Need $man 'Compile\BatchImportManifest.csv (run Compile first)'
        $rows = @(Import-Csv $man | Where-Object { $_.ManifestAction -eq 'OVERRIDES' })
        $strip = @(); $sp = Join-Path $compile 'StrippedElements.csv'
        if (Test-Path $sp) { $strip = @(Import-Csv $sp | Where-Object { $_.Kind -like '*Override*' -and ($rows.MPID -contains $_.MP) }) }
        $sum = { param($col) ($rows | Measure-Object -Property $col -Sum).Sum }
        $reason = {
            param($t)
            if ($t -like '*specific source object*') { "Per-server override, server/object not in target" }
            elseif ($t -like '*no longer exists in*') { "Monitor/rule removed from the target version of its MP" }
            elseif ($t -like '*unavailable*') { "Its MP is not in target (retired product or MS pack not installed)" }
            elseif ($t -like '*was removed*' -or $t -like '*stripped*') { 'Points at a dropped group' }
            else { 'Other' } }
        Write-Host ""
        Write-Host "=========== OVERRIDE PACKS ($($rows.Count)) ===========" -ForegroundColor Cyan
        Write-Host ("Packs READY / BLOCKED        : {0} / {1}" -f @($rows | Where-Object Verdict -eq 'READY').Count, @($rows | Where-Object Verdict -ne 'READY').Count)
        Write-Host ("Overrides kept               : {0}" -f (& $sum 'OverridesKept')) -ForegroundColor Green
        Write-Host ("  of which re-pointed per-server : {0}" -f (& $sum 'InstanceRemapped')) -ForegroundColor Green
        Write-Host ("Overrides removed            : {0}" -f $strip.Count) -ForegroundColor Yellow
        $strip | ForEach-Object { & $reason $_.Reason } | Group-Object | Sort-Object Count -Descending | ForEach-Object { Write-Host ("  {0,6}  {1}" -f $_.Count, $_.Name) -ForegroundColor Yellow }
        $un = & $sum 'InstanceUnmapped'
        if ($un -gt 0) { Write-Host ("Per-server overrides NOT mapped (would do nothing): {0} -- run MapInstances, then Compile" -f $un) -ForegroundColor Red }
        $top = $strip | Where-Object { (& $reason $_.Reason) -like 'Its MP is not in target*' } | ForEach-Object { if ($_.Reason -match "MP '([^']+)'") { $Matches[1] } } | Group-Object | Sort-Object Count -Descending | Select-Object -First 8
        if ($top) {
            Write-Host "Most-needed MPs that are not in the target (install the current version if still used, then re-export Target + Compile):" -ForegroundColor Cyan
            $top | ForEach-Object { Write-Host ("  {0,6}  {1}" -f $_.Count, $_.Name) }
        }
    }

    'Compile' {
        Need (Join-Path $src 'AllMPs') "source export ($srcName\AllMPs)"
        Need $tgt "target export ($tgtName\AllMPs)"
        $args2 = @{
            NonInteractive   = $true
            InputPath        = @((Join-Path $src 'AllMPs')) + @(if (Test-Path (Join-Path $src 'SealedOriginals')) { Join-Path $src 'SealedOriginals' })
            RepositoryFolder = $tgt
            Manifest         = (Join-Path $root 'MigrationManifest.csv')
            OutputFolder     = $compile
            SourceVersion    = $SourceVersion
            TargetVersion    = $TargetVersion
        }
        if (Test-Path (Join-Path $src 'SourceInventory.csv')) { $args2.SourceInventory = Join-Path $src 'SourceInventory.csv' }
        if (Test-Path (Join-Path $src 'GroupConversion.csv')) { $args2.GroupConversionFile = Join-Path $src 'GroupConversion.csv' }
        if (Test-Path (Join-Path $root 'InstanceMap.csv')) { $args2.InstanceMapFile = Join-Path $root 'InstanceMap.csv' }
        $dep = Join-Path $src 'SealedDependencies'
        if ((Test-Path $dep) -and @(Get-ChildItem $dep -Include *.mp, *.mpb -Recurse).Count -gt 0) { $args2.SourceRepositoryFolder = $dep }
        & (Join-Path $root 'MPMigration.ps1') @args2
    }

    'DryRun' {
        Need (Join-Path $compile $importScriptName) 'compiled output (run Compile first)'
        & (Join-Path $compile $importScriptName) -ManagementServer $ManagementServer -WhatIfOnly
    }

    'Import' {
        Need (Join-Path $compile $importScriptName) 'compiled output (run Compile first)'
        & (Join-Path $compile $importScriptName) -ManagementServer $ManagementServer
    }

    'Reimport' {
        Need (Join-Path $compile $importScriptName) 'compiled output (run Compile first)'
        & (Join-Path $compile $importScriptName) -ManagementServer $ManagementServer -Reimport
    }

    'TestGroups' {
        # For every group rule the compile converted, compare: members the group
        # has in the target now, target Windows computers the pattern matches,
        # and the source member names. Read-only.
        $gcr = Join-Path $compile 'GroupConversionResults.csv'
        Need $gcr 'Compile\GroupConversionResults.csv (run Compile first)'
        Import-Module OperationsManager
        New-SCOMManagementGroupConnection -ComputerName $ManagementServer | Out-Null

        $netbiosOf = { param($o)
            $fn = [string]$o.FullName
            $h = if ($fn.Contains(':')) { $fn.Substring($fn.IndexOf(':') + 1) } else { [string]$o.DisplayName }
            ((($h -split ';')[0]) -split '\.')[0].ToUpperInvariant() }

        $wc = Get-SCOMClass -Name 'Microsoft.Windows.Computer'
        $computers25 = @(Get-SCOMClassInstance -Class $wc | ForEach-Object { & $netbiosOf $_ } | Sort-Object -Unique)
        Write-Host "Windows computers reporting to this (target) management group: $($computers25.Count)" -ForegroundColor Cyan

        $conv = @{}
        $gcSrc = Join-Path $src 'GroupConversion.csv'
        if (Test-Path $gcSrc) { foreach ($r in @(Import-GroupConversion $gcSrc)) { $conv["$($r.MP)|$($r.DiscoveryID)|$($r.RuleIndex)"] = $r } }

        $out = New-Object System.Collections.Generic.List[object]
        $memberCache = @{}
        foreach ($r in @(Import-Csv $gcr | Where-Object Action -eq 'ConvertedToDynamic')) {
            if (-not $memberCache.ContainsKey($r.GroupClass)) {
                $n = $null
                try {
                    $gcls = Get-SCOMClass -Name $r.GroupClass -ErrorAction Stop
                    $gi = @(Get-SCOMClassInstance -Class $gcls -ErrorAction Stop) | Select-Object -First 1
                    if ($gi) { $n = @($gi.GetRelatedMonitoringObjects()).Count }
                } catch { }
                $memberCache[$r.GroupClass] = $n
            }
            $members = $memberCache[$r.GroupClass]
            $pats = @(if ($r.Operator -eq 'MatchesWildcard') { $r.Pattern -split ';' } else { $r.Pattern })
            $match25 = @($computers25 | Where-Object { $nm = $_; @($pats | Where-Object { if ($r.Operator -eq 'MatchesWildcard') { $nm -like $_ } else { [regex]::IsMatch($nm, $_) } }).Count -gt 0 })
            $src16 = $conv["$($r.MP)|$($r.DiscoveryID)|$($r.RuleIndex)"]
            $names16 = if ($src16) { [string]$src16.MemberNames } else { '' }
            $kind = if ($src16) { [string]$src16.MemberKind } else { '' }
            $status = if ($null -eq $members) { 'GROUP NOT FOUND (MP not imported?)' }
                      elseif ($r.Pattern -match '^\^\((MICROSOFT|WINDOWS)[\)|]' -or ($names16 -and $names16 -notmatch '[0-9A-Z]{2,}')) { 'BROKEN - bad server name (re-export needed)' }
                      elseif ($members -gt 0) { 'WORKING' }
                      elseif ($match25.Count -eq 0) { "WAITING - none of its servers report to the target yet" }
                      else { "NOT CALCULATED YET - servers are in the target, group still empty" }
            $out.Add([PSCustomObject]@{
                Status = $status; GroupClass = $r.GroupClass; MP = $r.MP; MemberKind = $kind; MembersInTarget = $members
                MatchingAgentsInTarget = $match25.Count; Pattern = $r.Pattern; MemberNamesSource = $names16
                MatchingNamesTarget = ($match25 -join '; ') })
        }
        $file = Join-Path $root "GroupTest-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
        $out | Export-Csv -LiteralPath $file -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "=========== GROUP TEST ($($out.Count) converted rules) ===========" -ForegroundColor Cyan
        $out | Group-Object Status | Sort-Object Name | ForEach-Object {
            $c = switch -Wildcard ($_.Name) { 'WORKING' { 'Green' } 'WAITING*' { 'Yellow' } 'NOT CALC*' { 'Yellow' } default { 'Red' } }
            Write-Host ("{0,5}  {1}" -f $_.Count, $_.Name) -ForegroundColor $c
        }
        Write-Host "Detail (open in Excel, filter Status): $file" -ForegroundColor Cyan
    }

    'WhyBlocked' {
        $man = Join-Path $compile 'BatchImportManifest.csv'
        Need $man 'Compile\BatchImportManifest.csv (run Compile first)'
        $rows = @(Import-Csv $man)
        $blocked = @($rows | Where-Object Verdict -ne 'READY')
        $report = Join-Path $root "WhyBlocked-$(Get-Date -Format 'yyyyMMdd-HHmm').txt"
        $lines = New-Object System.Collections.Generic.List[string]
        # Short category per MP so the screen fits on one screenshot.
        $cat = {
            param($r)
            $t = [string]$r.BlockingReasons
            if ($t -like 'SEALED in the source*') { 'Sealed - no original .mp' }
            elseif ($t -match 'Needs (\S+) v[\d.]+ \[(BatchMemberNeedsSealedOriginal|BatchVersionLower)') { "Needs blocked pack $($Matches[1])" }
            elseif ($t -match 'Needs (\S+) v[\d.]+ \[(NotInTargetRepo|NotInstalledInTarget)') { "Needs pack not in target: $($Matches[1])" }
            elseif ($t -match "Element '([^!]+)!") { "Uses element missing from target version of $($Matches[1])" }
            else { ($t -split ' \| ')[0] }
        }
        foreach ($b in $blocked) {
            $lines.Add("$($b.MPID)")
            foreach ($x in (([string]$b.BlockingReasons) -split ' \| ')) { if ($x) { $lines.Add("    - $x") } }
        }
        $strip = Join-Path $compile 'StrippedElements.csv'
        $dropped = @()
        if (Test-Path $strip) { $dropped = @(Import-Csv $strip | Where-Object Kind -eq 'Group (DROPPED)') }
        $lines.Add(""); $lines.Add("=== Groups dropped so their MP could import ($($dropped.Count))")
        foreach ($d in $dropped) { $lines.Add("    $($d.MP): $($d.ElementID)") }
        $lines | Set-Content -LiteralPath $report -Encoding UTF8

        Write-Host ""
        Write-Host "=========== WHY BLOCKED ($($blocked.Count) of $($rows.Count) MPs) ===========" -ForegroundColor Cyan
        $blocked | ForEach-Object { & $cat $_ } | Group-Object | Sort-Object Count -Descending | ForEach-Object {
            Write-Host ("{0,4}  {1}" -f $_.Count, $_.Name) -ForegroundColor Yellow
        }
        Write-Host ("{0,4}  groups dropped (whole MP imports without them)" -f $dropped.Count) -ForegroundColor Green
        Write-Host "Full list with every reason: $report" -ForegroundColor Cyan
    }

    'Rollback' {
        # Removes every MP the most recent Import/Reimport run reported as
        # Imported, newest first (reverse dependency order).
        $res = Get-ChildItem $compile -Filter 'ImportResults_*.csv' | Sort-Object LastWriteTime -Descending
        $rows = @(); foreach ($f in $res) { $rows += @(Import-Csv $f.FullName | Where-Object Result -eq 'Imported') }
        $ids = @($rows | Sort-Object { [int]$_.Order } -Descending | ForEach-Object { $_.MPID } | Select-Object -Unique)
        if ($ids.Count -eq 0) { Write-Host "Nothing recorded as Imported -- nothing to roll back."; return }
        Import-Module OperationsManager
        New-SCOMManagementGroupConnection -ComputerName $ManagementServer | Out-Null
        $mg = @(Get-SCOMManagementGroupConnection | Where-Object IsActive) | Select-Object -First 1
        Write-Host "This will REMOVE $($ids.Count) MP(s) from management group '$($mg.ManagementGroupName)'." -ForegroundColor Red
        if ((Read-Host "Type YES to remove them") -ne 'YES') { Write-Host "Cancelled. Nothing removed."; return }
        $log = Join-Path $root "Rollback-$(Get-Date -Format 'yyyyMMdd-HHmm').csv"
        $done = New-Object System.Collections.Generic.List[object]
        for ($pass = 1; $pass -le 3 -and $ids.Count -gt 0; $pass++) {
            $left = @()
            foreach ($id in $ids) {
                $mp = Get-SCOMManagementPack -Name $id -ErrorAction SilentlyContinue
                if (-not $mp) { $done.Add([PSCustomObject]@{ MPID = $id; Result = 'NotPresent' }); continue }
                try { Remove-SCOMManagementPack -ManagementPack $mp -ErrorAction Stop; $done.Add([PSCustomObject]@{ MPID = $id; Result = 'Removed' }); Write-Host "Removed $id" -ForegroundColor Green }
                catch { $left += $id; if ($pass -eq 3) { $done.Add([PSCustomObject]@{ MPID = $id; Result = "Failed: $($_.Exception.Message)" }); Write-Host "FAILED $id : $($_.Exception.Message)" -ForegroundColor Red } }
            }
            $ids = $left
        }
        $done | Export-Csv -LiteralPath $log -NoTypeInformation -Encoding UTF8
        Write-Host "Rollback log: $log" -ForegroundColor Cyan
    }

    'Collect' {
        Need $compile 'Compile folder'
        $zip = Join-Path $root "MigrationResults-$(Get-Date -Format 'yyyyMMdd-HHmm').zip"
        $files = @()
        $files += Get-ChildItem $compile -File | Where-Object { $_.Name -match '^(ImportResults_|ImportErrors_|BatchImportManifest|MissingDependencies|GroupConversionResults|StrippedElements|ManifestMatch|StaticGroupMembership|MPCompiler\.)' }
        $files += Get-ChildItem $src -File -Filter *.csv -ErrorAction SilentlyContinue | Where-Object { $_.Name -ne 'SourceInventory.csv' }
        Compress-Archive -LiteralPath ($files.FullName) -DestinationPath $zip -Force
        Write-Host "Send this back: $zip ($($files.Count) files)" -ForegroundColor Cyan
    }
}
