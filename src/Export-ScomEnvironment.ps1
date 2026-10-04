<#
.SYNOPSIS
    Exports a SCOM management group's MPs plus an inventory, as input for
    MPMigration.ps1. Run it ON (or against) each environment's management
    server with that environment's own OperationsManager module.

.DESCRIPTION
    -Role Source (the OLD management group, e.g. SCOM 2016):
      * SourceInventory.csv   every installed MP with Version, Sealed, KeyToken,
                              DisplayName, TimeCreated, LastModified
      * AllMPs\               Export-SCOMManagementPack of every MP (sealed MPs
                              come out as .xml -- analysis only, never import these)
      * With -Manifest: ManifestResolution.csv mapping each workbook row marked
        Migrate = Y to the real MP ID(s), and whether it is sealed
      * SealedOriginals\      original .mp/.mpb files found for sealed in-scope
                              MPs (searched in -SealedSearchPath and the usual
                              install/extract folders); SealedOriginalsMissing.csv
                              lists the ones still needed

      * GroupConversion.csv   one row per static group membership rule: the
                              member servers (GUIDs resolved to names here,
                              while the source still knows them), the dynamic
                              NetBIOS-name pattern that replaces the static
                              list, and how that pattern behaves against EVERY
                              Windows computer in the source (extra matches, misses).
                              Edit Pattern / Convert before compiling.
      * StaticGroupMembers.csv one row per static member GUID
      * OverrideInstances.csv  every object that an override targets by ID
                              ("disable on server X"), with its FullName, so
                              the same object can be found in the target

    -Role Target (the NEW management group, e.g. SCOM 2025):
      * TargetInventory.csv
      * AllMPs\               complete export = -RepositoryFolder for MPMigration.ps1

    Read-only against SCOM in both roles: only Get-SCOMManagementPack and
    Export-SCOMManagementPack are called.

.EXAMPLE
    # On the source (old) management server
    .\Export-ScomEnvironment.ps1 -Role Source -Manifest .\MigrationManifest.csv -OutputFolder D:\SCOMMigration\Source

.EXAMPLE
    # On the target (new) management server
    .\Export-ScomEnvironment.ps1 -Role Target -OutputFolder D:\SCOMMigration\Target
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('Source', 'Target')][string]$Role,
    [string]$ManagementServer = 'localhost',
    [System.Management.Automation.PSCredential]$Credential,
    [string]$OutputFolder,
    [string]$Manifest,
    [string[]]$SealedSearchPath,
    # Static-group conversion (Source role). Tight = NAME-STEM followed by
    # digits only (APPSQL01, APPSQL02 -> ^(APPSQL[0-9]+)$). Wildcard =
    # NAME-STEM* (also catches APPSQLREPORT01). Exact = the listed names only.
    [ValidateSet('Tight', 'Wildcard', 'Exact')][string]$GroupPatternMode = 'Tight',
    [switch]$SkipGroupConversion
)

$ErrorActionPreference = 'Stop'
Set-StrictMode -Version Latest

if (-not $OutputFolder) { $OutputFolder = Join-Path 'C:\SCOMMigration' "$Role-$(Get-Date -Format 'yyyyMMdd-HHmmss')" }
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$OutputFolder = (Resolve-Path -LiteralPath $OutputFolder).ProviderPath
$log = Join-Path $OutputFolder "Export-$Role.log"
function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c; Add-Content -LiteralPath $log -Value "[$(Get-Date -Format s)] $m" -Encoding UTF8 }

Import-Module OperationsManager
if ($Credential) { New-SCOMManagementGroupConnection -ComputerName $ManagementServer -Credential $Credential | Out-Null }
else { New-SCOMManagementGroupConnection -ComputerName $ManagementServer | Out-Null }
$conn = @(Get-SCOMManagementGroupConnection | Where-Object { $_.IsActive }) | Select-Object -First 1
if (-not $conn) { throw "No active SCOM connection after connecting to '$ManagementServer'." }
Say "Connected to management group '$($conn.ManagementGroupName)' via '$($conn.ManagementServerName)' as $Role" 'Cyan'

$mps = @(Get-SCOMManagementPack)
Say "Installed MPs: $($mps.Count)"

$invName = if ($Role -eq 'Source') { 'SourceInventory.csv' } else { 'TargetInventory.csv' }
$inv = $mps | ForEach-Object {
    [PSCustomObject]@{
        Name         = $_.Name
        DisplayName  = $_.DisplayName
        Version      = [string]$_.Version
        Sealed       = [bool]$_.Sealed
        KeyToken     = [string]$_.KeyToken
        TimeCreated  = $_.TimeCreated
        LastModified = $_.LastModified
    }
}
$inv | Sort-Object Name | Export-Csv -LiteralPath (Join-Path $OutputFolder $invName) -NoTypeInformation -Encoding UTF8
Say "Inventory written: $invName ($(@($inv | Where-Object Sealed).Count) sealed, $(@($inv | Where-Object { -not $_.Sealed }).Count) unsealed)" 'Green'

$allFolder = Join-Path $OutputFolder 'AllMPs'
New-Item -ItemType Directory -Path $allFolder -Force | Out-Null
$fail = 0
foreach ($mp in $mps) {
    try { $mp | Export-SCOMManagementPack -Path $allFolder -ErrorAction Stop }
    catch { $fail++; Say "  export failed: $($mp.Name) v$($mp.Version): $($_.Exception.Message)" 'Yellow' }
}
Say "Exported $($mps.Count - $fail) of $($mps.Count) MPs to $allFolder" $(if ($fail) { 'Yellow' } else { 'Green' })

if ($Role -eq 'Target') {
    Say ""
    Say "Use this as -RepositoryFolder for MPMigration.ps1:  $allFolder" 'Cyan'
    return
}

# ---------------------------------------------------------------- Source only
New-Item -ItemType Directory -Path (Join-Path $OutputFolder 'SealedOriginals') -Force | Out-Null

# ======================= STATIC GROUP -> DYNAMIC GROUP ======================
# Static members are object GUIDs that only exist in 2016. Resolve them to
# server names HERE, derive a NetBIOS-name pattern per membership rule, and
# test each pattern against every Windows computer 2016 knows about, so the
# conversion is reviewed against real data before anything is imported.

function ConvertTo-RegexLiteral([string]$t) { return ([regex]::Escape($t) -replace '\ ', ' ') }

function Get-NamePattern {
    param([string[]]$Names, [string]$Mode)
    $names = @($Names | Where-Object { $_ } | ForEach-Object { $_.ToUpperInvariant() } | Sort-Object -Unique)
    if ($names.Count -eq 0) { return $null }
    $parts = New-Object System.Collections.Generic.List[string]
    foreach ($n in $names) {
        $m = [regex]::Match($n, '^(.*[^0-9])([0-9]+)$')
        if ($Mode -eq 'Exact' -or -not $m.Success) { $p = if ($Mode -eq 'Wildcard') { $n } else { ConvertTo-RegexLiteral $n } }
        elseif ($Mode -eq 'Wildcard') { $p = $m.Groups[1].Value + '*' }
        else { $p = (ConvertTo-RegexLiteral $m.Groups[1].Value) + '[0-9]+' }
        if (-not $parts.Contains($p)) { $parts.Add($p) }
    }
    if ($Mode -eq 'Wildcard') {
        return [PSCustomObject]@{ Operator = 'MatchesWildcard'; Pattern = ($parts -join ';') }   # ';' = OR of wildcards
    }
    return [PSCustomObject]@{ Operator = 'MatchesRegularExpression'; Pattern = '^(' + ($parts -join '|') + ')$' }
}

function Test-NamePattern([string]$Name, [string]$Operator, [string]$Pattern) {
    if ($Operator -eq 'MatchesWildcard') {
        foreach ($w in ($Pattern -split ';')) { if ($w -and $Name -like $w) { return $true } }
        return $false
    }
    return [regex]::IsMatch($Name, $Pattern)
}

# Host computer NetBIOS name of an object. FullName is "<Class>:<host FQDN>[;...]"
# for every hosted object and for computers themselves, so it is tried first.
# DisplayName is only trusted for computer objects -- for anything else it is
# a description ("Microsoft Windows Server 2016 Standard"), not a server name.
function Get-HostNetbios($inst, [bool]$IsComputer = $false) {
    $cands = New-Object System.Collections.Generic.List[string]
    try { $fn = [string]$inst.FullName; if ($fn -and $fn.Contains(':')) { $cands.Add($fn.Substring($fn.IndexOf(':') + 1)) } } catch { }
    try { if ($inst.Path) { $cands.Add([string]$inst.Path) } } catch { }
    if ($IsComputer) { try { $cands.Add([string]$inst.DisplayName) } catch { } }
    foreach ($c in $cands) {
        $h = ($c -split ';')[0].Trim()
        $n = (($h -split '\.')[0]).ToUpperInvariant()
        if ($n -match '^[A-Z0-9][A-Z0-9_-]{0,14}$') { return $n }
    }
    return ''
}

if (-not $SkipGroupConversion) {
    Say ""
    Say "Static group scan (pattern mode: $GroupPatternMode)..." 'Cyan'
    $unsealedFiles = @($mps | Where-Object { -not $_.Sealed } | ForEach-Object { Join-Path $allFolder "$($_.Name).xml" } | Where-Object { Test-Path -LiteralPath $_ })
    $ruleRows = New-Object System.Collections.Generic.List[object]
    foreach ($f in $unsealedFiles) {
        $doc = New-Object System.Xml.XmlDocument
        try { $doc.Load($f) } catch { continue }
        $mpId = [string]$doc.SelectSingleNode('/ManagementPack/Manifest/Identity/ID').InnerText
        foreach ($disc in @($doc.SelectNodes('/ManagementPack/Monitoring/Discoveries/Discovery'))) {
            $ds = $disc.SelectSingleNode('DataSource')
            if (-not $ds -or [string]$ds.GetAttribute('TypeID') -notlike '*GroupPopulator*') { continue }
            $gi = $ds.SelectSingleNode('GroupInstanceId')
            $groupClass = if ($gi) { ([string]$gi.InnerText) -replace '^\$MPElement\[Name="?', '' -replace '"?\]\$$', '' } else { [string]$disc.GetAttribute('Target') }
            $ri = 0
            foreach ($mr in @($ds.SelectNodes('MembershipRules/MembershipRule'))) {
                $ri++
                $inc = @($mr.SelectNodes('IncludeList/MonitoringObjectId') | ForEach-Object { ([string]$_.InnerText).Trim() })
                $exc = @($mr.SelectNodes('ExcludeList/MonitoringObjectId') | ForEach-Object { ([string]$_.InnerText).Trim() })
                if ($inc.Count -eq 0 -and $exc.Count -eq 0) { continue }
                $mcN = $mr.SelectSingleNode('MonitoringClass')
                $ruleRows.Add([PSCustomObject]@{
                    MP = $mpId; DiscoveryID = [string]$disc.GetAttribute('ID'); GroupClass = $groupClass; RuleIndex = $ri
                    RuleClass = if ($mcN) { [string]$mcN.InnerText } else { '' }
                    HasExpression = [bool]$mr.SelectSingleNode('Expression'); Include = $inc; Exclude = $exc
                })
            }
        }
    }
    Say "Membership rules with static include/exclude lists: $($ruleRows.Count) in $(@($ruleRows | Select-Object -ExpandProperty MP -Unique).Count) MP(s)"

    if ($ruleRows.Count -gt 0) {
        # Resolve every GUID once.
        $allGuids = @($ruleRows | ForEach-Object { $_.Include; $_.Exclude } | Where-Object { $_ } | Sort-Object -Unique)
        $resolved = @{}
        $chunk = 200
        for ($i = 0; $i -lt $allGuids.Count; $i += $chunk) {
            $slice = @($allGuids[$i..([Math]::Min($i + $chunk, $allGuids.Count) - 1)] | ForEach-Object { [guid]$_ })
            try {
                foreach ($inst in @(Get-SCOMClassInstance -Id $slice -ErrorAction SilentlyContinue)) {
                    $cls = ''; $single = $false
                    try { $c0 = $inst.GetLeastDerivedNonAbstractClass(); $cls = [string]$c0.Name; $single = [bool]$c0.Singleton } catch { }
                    $resolved[([string]$inst.Id).ToLowerInvariant()] = [PSCustomObject]@{
                        DisplayName = [string]$inst.DisplayName; Path = [string]$inst.Path; Class = $cls; Singleton = $single
                        Netbios = if ($single) { '' } elseif ($cls -like '*Watcher*') { Get-HostNetbios ([pscustomobject]@{ FullName = ''; Path = ''; DisplayName = [string]$inst.DisplayName }) $true } else { Get-HostNetbios $inst ($cls -like '*Computer*') }
                    }
                }
            }
            catch { Say "  GUID lookup failed for a batch: $($_.Exception.Message)" 'Yellow' }
        }
        Say "Resolved $($resolved.Count) of $($allGuids.Count) member GUID(s) (the rest are objects already gone from the source)."

        # Every Windows computer in 2016, for testing the patterns.
        $allComputers = @()
        try {
            $wc = Get-SCOMClass -Name 'Microsoft.Windows.Computer'
            $allComputers = @(Get-SCOMClassInstance -Class $wc | ForEach-Object { Get-HostNetbios $_ $true } | Where-Object { $_ } | Sort-Object -Unique)
        }
        catch { Say "  Could not list Windows computers: $($_.Exception.Message)" 'Yellow' }
        Say "Windows computers in the source for pattern testing: $($allComputers.Count)"

        $memberRows = New-Object System.Collections.Generic.List[object]
        $convRows = New-Object System.Collections.Generic.List[object]
        foreach ($r in $ruleRows) {
            $inc = @($r.Include | ForEach-Object { $x = $resolved[$_.ToLowerInvariant()]; foreach ($g in @($_)) { $memberRows.Add([PSCustomObject]@{ MP = $r.MP; GroupClass = $r.GroupClass; RuleIndex = $r.RuleIndex; List = 'Include'; Guid = $g; Resolved = [bool]$x; Netbios = if ($x) { $x.Netbios } else { '' }; DisplayName = if ($x) { $x.DisplayName } else { '' }; Class = if ($x) { $x.Class } else { '' } }) }; $x } | Where-Object { $_ })
            $exc = @($r.Exclude | ForEach-Object { $x = $resolved[$_.ToLowerInvariant()]; $memberRows.Add([PSCustomObject]@{ MP = $r.MP; GroupClass = $r.GroupClass; RuleIndex = $r.RuleIndex; List = 'Exclude'; Guid = $_; Resolved = [bool]$x; Netbios = if ($x) { $x.Netbios } else { '' }; DisplayName = if ($x) { $x.DisplayName } else { '' }; Class = if ($x) { $x.Class } else { '' } }); $x } | Where-Object { $_ })

            $incNames = @($inc | Where-Object { $_.Netbios } | ForEach-Object { $_.Netbios } | Sort-Object -Unique)
            $excNames = @($exc | Where-Object { $_.Netbios } | ForEach-Object { $_.Netbios } | Sort-Object -Unique)
            $isComputerRule = ($r.RuleClass -match 'Microsoft\.Windows\.(Server\.|Client\.)?Computer"?\]\$$')
            $isGroupMember = (@($inc | Where-Object { $_.Singleton }).Count -gt 0 -and $incNames.Count -eq 0)
            $memberKind = if ($isComputerRule) { 'Computer' } elseif ($isGroupMember) { 'Group' } else { 'Hosted' }

            $pat = Get-NamePattern -Names $incNames -Mode $GroupPatternMode
            $matchIn2016 = @(); $extra = @(); $missed = @()
            if ($pat -and $allComputers.Count -gt 0) {
                $matchIn2016 = @($allComputers | Where-Object { Test-NamePattern $_ $pat.Operator $pat.Pattern })
                $extra  = @($matchIn2016 | Where-Object { $incNames -notcontains $_ -and $excNames -notcontains $_ })
                $missed = @($incNames | Where-Object { $matchIn2016 -notcontains $_ })
            }
            $rec = if ($memberKind -eq 'Group') { 'NESTED GROUP (converted automatically, no pattern needed)' }
                   elseif (-not $pat) { 'NO RESOLVABLE MEMBERS - leave static' }
                   elseif ($extra.Count -gt 0) { "REVIEW - pattern also matches $($extra.Count) server(s) not in the group today" }
                   else { 'CONVERT' }
            $convRows.Add([PSCustomObject]@{
                Convert        = if ($pat) { 'Y' } else { 'N' }
                MP             = $r.MP
                GroupClass     = $r.GroupClass
                DiscoveryID    = $r.DiscoveryID
                RuleIndex      = $r.RuleIndex
                RuleClass      = $r.RuleClass
                MemberKind     = $memberKind
                Recommendation = $rec
                Operator       = if ($pat) { $pat.Operator } else { '' }
                Pattern        = if ($pat) { $pat.Pattern } else { '' }
                ExcludePattern = if ($excNames.Count -gt 0) { '^(' + ((@($excNames | ForEach-Object { ConvertTo-RegexLiteral $_ })) -join '|') + ')$' } else { '' }
                StaticMembers  = $r.Include.Count
                ResolvedMembers= $incNames.Count
                StaleMembers   = $r.Include.Count - @($inc).Count
                MatchesInSource  = $matchIn2016.Count
                ExtraInSource    = ($extra -join '; ')
                MissedMembers  = ($missed -join '; ')
                MemberNames    = ($incNames -join '; ')
                HadExpression  = $r.HasExpression
            })
        }
        $memberRows | Export-Csv -LiteralPath (Join-Path $OutputFolder 'StaticGroupMembers.csv') -NoTypeInformation -Encoding UTF8
        $convRows | Export-Csv -LiteralPath (Join-Path $OutputFolder 'GroupConversion.csv') -NoTypeInformation -Encoding UTF8
        $rev = @($convRows | Where-Object { $_.Recommendation -like 'REVIEW*' })
        $nested = @($convRows | Where-Object { $_.MemberKind -eq 'Group' }).Count
        Say "GroupConversion.csv: $(@($convRows | Where-Object Convert -eq 'Y').Count) rule(s) convert to dynamic NetBIOS patterns ($($rev.Count) of them also match extra servers -- review), $nested nested subgroup rule(s) convert automatically, $(@($convRows | Where-Object { $_.Convert -eq 'N' -and $_.MemberKind -ne 'Group' }).Count) stay static (no resolvable members)." 'Green'
        foreach ($rv in ($rev | Select-Object -First 15)) { Say "  REVIEW $($rv.GroupClass): $($rv.Pattern) also matches $($rv.ExtraInSource)" 'Yellow' }
    }
}
# =================== PER-OBJECT OVERRIDES (ContextInstance) ==================
# "Disable this monitor on server X" overrides point at a 2016 object ID. The
# 2025 object has a different ID, so record what each ID IS (FullName = class
# plus key, e.g. "Microsoft.Windows.Computer:server.domain") while 2016 still
# knows. Invoke-ScomMigrationStep.ps1 MapInstances finds the same object in the target.
if (-not $SkipGroupConversion) {
    Say ""
    Say "Per-object override scan..." 'Cyan'
    $ovFiles = @($mps | Where-Object { -not $_.Sealed } | ForEach-Object { Join-Path $allFolder "$($_.Name).xml" } | Where-Object { Test-Path -LiteralPath $_ })
    $ovGuids = @{}
    foreach ($f in $ovFiles) {
        $doc = New-Object System.Xml.XmlDocument
        try { $doc.Load($f) } catch { continue }
        $mpId = [string]$doc.SelectSingleNode('/ManagementPack/Manifest/Identity/ID').InnerText
        foreach ($o in @($doc.SelectNodes('/ManagementPack/Monitoring/Overrides/*[@ContextInstance]'))) {
            $g = ([string]$o.GetAttribute('ContextInstance')).Trim().ToLowerInvariant()
            if (-not $g) { continue }
            if (-not $ovGuids.ContainsKey($g)) { $ovGuids[$g] = New-Object System.Collections.Generic.List[string] }
            if (-not $ovGuids[$g].Contains($mpId)) { $ovGuids[$g].Add($mpId) }
        }
    }
    Say "Overrides aimed at a specific object: $($ovGuids.Count) distinct object(s)"
    if ($ovGuids.Count -gt 0) {
        $oiRows = New-Object System.Collections.Generic.List[object]
        $allG = @($ovGuids.Keys)
        $found = @{}
        for ($i = 0; $i -lt $allG.Count; $i += 200) {
            $slice = @($allG[$i..([Math]::Min($i + 200, $allG.Count) - 1)] | ForEach-Object { [guid]$_ })
            try {
                foreach ($inst in @(Get-SCOMClassInstance -Id $slice -ErrorAction SilentlyContinue)) {
                    $cls = ''; try { $cls = [string]$inst.GetLeastDerivedNonAbstractClass().Name } catch { }
                    $found[([string]$inst.Id).ToLowerInvariant()] = [PSCustomObject]@{ FullName = [string]$inst.FullName; DisplayName = [string]$inst.DisplayName; Path = [string]$inst.Path; Class = $cls }
                }
            }
            catch { Say "  lookup failed for a batch: $($_.Exception.Message)" 'Yellow' }
        }
        foreach ($g in $allG) {
            $x = $found[$g]
            $oiRows.Add([PSCustomObject]@{
                Guid = $g; Resolved = [bool]$x
                FullName = if ($x) { $x.FullName } else { '' }; DisplayName = if ($x) { $x.DisplayName } else { '' }
                Path = if ($x) { $x.Path } else { '' }; Class = if ($x) { $x.Class } else { '' }
                UsedByMPs = ($ovGuids[$g] -join '; ')
            })
        }
        $oiRows | Export-Csv -LiteralPath (Join-Path $OutputFolder 'OverrideInstances.csv') -NoTypeInformation -Encoding UTF8
        Say "OverrideInstances.csv: $($found.Count) of $($allG.Count) object(s) resolved (the rest no longer exist in the source -- those overrides were already doing nothing)." 'Green'
    }
}

if ($Manifest) {
    $allRows = @(Import-Csv -LiteralPath $Manifest)
    if ($allRows.Count -eq 0 -or -not $allRows[0].PSObject.Properties['Migrate'] -or -not $allRows[0].PSObject.Properties['ManagementPack']) { throw "Manifest needs 'Migrate' and 'ManagementPack' columns." }
    $rows = @($allRows | Where-Object { [string]$_.Migrate -match '^(y|yes|true|1)$' })
    Say "Manifest rows marked Migrate = Y: $($rows.Count)"
    $res = foreach ($r in $rows) {
        $pat = if ($r.PSObject.Properties['MatchPattern'] -and $r.MatchPattern) { $r.MatchPattern } else { $r.ManagementPack }
        try { [void]('x' -like $pat) } catch { $pat = [System.Management.Automation.WildcardPattern]::Escape($pat) }
        $act = if ($r.PSObject.Properties['Action']) { $r.Action } else { '' }
        $hits = @($mps | Where-Object { $_.Name -like $pat })
        $by = 'MPID'
        if ($hits.Count -eq 0) { $hits = @($mps | Where-Object { $_.DisplayName -like $pat }); $by = 'DisplayName' }
        if ($hits.Count -eq 0) {
            [PSCustomObject]@{ ManagementPack = $r.ManagementPack; Action = $act; Status = 'NO MATCH'; MatchedBy = ''; MPID = ''; DisplayName = ''; Version = ''; Sealed = '' }
        }
        foreach ($h in $hits) {
            [PSCustomObject]@{ ManagementPack = $r.ManagementPack; Action = $act; Status = 'Matched'; MatchedBy = $by; MPID = $h.Name; DisplayName = $h.DisplayName; Version = [string]$h.Version; Sealed = [bool]$h.Sealed }
        }
    }
    $res | Export-Csv -LiteralPath (Join-Path $OutputFolder 'ManifestResolution.csv') -NoTypeInformation -Encoding UTF8
    $noMatch = @($res | Where-Object { $_.Status -eq 'NO MATCH' })
    $sealedInScope = @($res | Where-Object { $_.Status -eq 'Matched' -and $_.Sealed -eq $true })
    Say "Manifest resolution: $(@($res | Where-Object Status -eq 'Matched').Count) matched, $($noMatch.Count) with no match, $($sealedInScope.Count) in-scope MPs are SEALED" 'Green'
    foreach ($n in $noMatch) { Say "  NO MATCH: $($n.ManagementPack)" 'Yellow' }

    # Look for the original sealed files of in-scope sealed MPs.
    if ($sealedInScope.Count -gt 0) {
        $roots = @()
        if ($SealedSearchPath) { $roots += $SealedSearchPath }
        $roots += @(
            "${env:ProgramFiles(x86)}\System Center Management Packs",
            "$env:ProgramFiles\System Center Management Packs",
            "$env:ProgramFiles\Microsoft System Center 2016\Operations Manager",
            "$env:ProgramFiles\Microsoft System Center\Operations Manager"
        )
        $roots = @($roots | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -Unique)
        Say "Searching for sealed originals under: $($roots -join '; ')"

        # Read ID/version from each candidate file with the SDK (the file name
        # is often not the MP ID).
        $sdkOk = [bool]('Microsoft.EnterpriseManagement.Configuration.ManagementPack' -as [type])
        if (-not $sdkOk) { try {
            $dll = Get-ChildItem -Path @("$env:ProgramFiles\Microsoft System Center*\Operations Manager", "$env:ProgramFiles\Microsoft System Center\Operations Manager") -Filter Microsoft.EnterpriseManagement.OperationsManager.dll -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
            if ($dll) {
                Add-Type -Path (Join-Path $dll.DirectoryName 'Microsoft.EnterpriseManagement.Core.dll') -ErrorAction SilentlyContinue
                Add-Type -Path $dll.FullName
                $pk = Join-Path $dll.DirectoryName 'Microsoft.EnterpriseManagement.Packaging.dll'
                if (Test-Path -LiteralPath $pk) { Add-Type -Path $pk -ErrorAction SilentlyContinue }
                $sdkOk = $true
            }
        }
        catch { Say "SDK not loadable ($($_.Exception.Message)); matching sealed files by file name only." 'Yellow' } }

        $found = @{}   # MPID -> list of file info
        $files = @(foreach ($r0 in $roots) { Get-ChildItem -LiteralPath $r0 -Recurse -Include *.mp, *.mpb -ErrorAction SilentlyContinue })
        Say "Candidate sealed files found: $($files.Count)"
        foreach ($f in $files) {
            $ids = @()
            if ($sdkOk -and $f.Extension -ieq '.mp') {
                try { $m = New-Object Microsoft.EnterpriseManagement.Configuration.ManagementPack($f.FullName); $ids += [PSCustomObject]@{ ID = $m.Name; Version = [string]$m.Version } } catch { }
            }
            elseif ($sdkOk -and $f.Extension -ieq '.mpb') {
                try {
                    $reader = [Microsoft.EnterpriseManagement.Packaging.ManagementPackBundleFactory]::CreateBundleReader()
                    $store = New-Object Microsoft.EnterpriseManagement.Configuration.IO.ManagementPackFileStore
                    $store.AddDirectory($f.DirectoryName)
                    $b = $reader.Read($f.FullName, $store)
                    foreach ($bm in $b.ManagementPacks) { $ids += [PSCustomObject]@{ ID = $bm.Name; Version = [string]$bm.Version } }
                }
                catch { }
            }
            if ($ids.Count -eq 0) { $ids += [PSCustomObject]@{ ID = $f.BaseName; Version = '' } }
            foreach ($i in $ids) {
                if (-not $found.ContainsKey($i.ID)) { $found[$i.ID] = @() }
                $found[$i.ID] += [PSCustomObject]@{ File = $f; Version = $i.Version }
            }
        }

        # Every sealed MP installed in 2016 is a candidate dependency too, so
        # copy originals for ALL sealed in-scope MPs AND anything else found.
        $origFolder = Join-Path $OutputFolder 'SealedOriginals'
        New-Item -ItemType Directory -Path $origFolder -Force | Out-Null
        $missing = New-Object System.Collections.Generic.List[object]
        # Best file for an MP: the exact installed version, else the highest
        # version found. Returns $null if the best is LOWER than installed.
        function Select-Original([string]$Id, [string]$InstalledVer) {
            if (-not $found.ContainsKey($Id)) { return $null }
            $exact = @($found[$Id] | Where-Object { $_.Version -eq $InstalledVer }) | Select-Object -First 1
            if ($exact) { return $exact }
            $best = @($found[$Id] | Sort-Object { try { [version]$_.Version } catch { [version]'0.0.0.0' } } -Descending) | Select-Object -First 1
            if ($best.Version) {
                try { if ([version]$best.Version -lt [version]$InstalledVer) { return $null } } catch { }
            }
            return $best
        }

        foreach ($s in $sealedInScope) {
            $installedVer = [string]$s.Version
            $pick = Select-Original $s.MPID $installedVer
            if ($pick) {
                Copy-Item -LiteralPath $pick.File.FullName -Destination $origFolder -Force
                Say "  original found: $($s.MPID) v$($pick.Version) <- $($pick.File.FullName)" 'Green'
                if ($pick.Version -and $pick.Version -ne $installedVer) { Say "    NOTE: installed in the source is v$installedVer, file is v$($pick.Version)" 'Yellow' }
            }
            else {
                $missing.Add([PSCustomObject]@{ MPID = $s.MPID; DisplayName = $s.DisplayName; InstalledVersion = $installedVer; WorkbookName = $s.ManagementPack; Action = $s.Action })
            }
        }

        # Sealed originals for everything ELSE installed in 2016 that we found a
        # file for: potential dependencies of the in-scope MPs. Vendor libraries
        # go to SealedDependencies (pass as -SourceRepositoryFolder). Microsoft
        # packs go to a separate folder and are NOT used by default -- the
        # target should run current Microsoft packs, not 2016-era ones.
        $depFolder = Join-Path $OutputFolder 'SealedDependencies'
        $msFolder  = Join-Path $OutputFolder 'SealedDependencies_Microsoft_optional'
        New-Item -ItemType Directory -Path $depFolder, $msFolder -Force | Out-Null
        $inScopeIds = @($sealedInScope | ForEach-Object { $_.MPID })
        $depCount = 0
        foreach ($m in @($mps | Where-Object { $_.Sealed -and $inScopeIds -notcontains $_.Name })) {
            $pick = Select-Original $m.Name ([string]$m.Version)
            if (-not $pick) { continue }
            $dest = if ($m.Name -like 'Microsoft.*' -or $m.Name -like 'System.*') { $msFolder } else { $depFolder }
            Copy-Item -LiteralPath $pick.File.FullName -Destination $dest -Force
            $depCount++
        }
        Say "Copied $depCount other sealed original(s): vendor ones to SealedDependencies\, Microsoft ones to SealedDependencies_Microsoft_optional\"
        $missing | Export-Csv -LiteralPath (Join-Path $OutputFolder 'SealedOriginalsMissing.csv') -NoTypeInformation -Encoding UTF8
        if ($missing.Count -gt 0) {
            Say "$($missing.Count) in-scope SEALED MP(s) have no original .mp/.mpb on this server -- see SealedOriginalsMissing.csv. Get them from the vendor / old install media, drop them in SealedOriginals\, and they will be picked up." 'Yellow'
        }
    }
}

Say ""
Say "Copy this whole folder to the target management server. For MPMigration.ps1 use:" 'Cyan'
Say "  -InputPath '$allFolder','$(Join-Path $OutputFolder 'SealedOriginals')'  -SourceInventory '$(Join-Path $OutputFolder $invName)'" 'Cyan'
if (Test-Path -LiteralPath (Join-Path $OutputFolder 'SealedDependencies')) {
    Say "  -SourceRepositoryFolder '$(Join-Path $OutputFolder 'SealedDependencies')'" 'Cyan'
}
