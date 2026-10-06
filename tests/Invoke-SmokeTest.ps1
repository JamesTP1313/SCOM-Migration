<#
.SYNOPSIS
    End-to-end smoke test of the toolkit against synthetic MPs and stand-in
    OperationsManager modules. No SCOM needed; runs on Windows PowerShell 5.1
    or PowerShell 7 (Windows, Linux, macOS).

.DESCRIPTION
    1. Source export   : Export-ScomEnvironment.ps1 -Role Source against tests/fixtures/source
    2. Target export   : tests/fixtures/target copied in as Target\AllMPs
    3. Steps           : Check, FixReview, EnableOverrides, MapInstances, Compile, WhyBlocked, OverrideReport, DryRun, TestGroups
    4. Assertions      : READY/BLOCKED verdicts, stripping, group conversion, per-object override re-pointing
    5. ExportSource    : the same source export run from the "target" through a
                         PowerShell-remoting stand-in (tests/mocks/remoting),
                         including the -SealedSearchPath pass, re-export and an
                         unreachable server

    Exit code 0 = all assertions passed.

.EXAMPLE
    pwsh ./tests/Invoke-SmokeTest.ps1
    pwsh ./tests/Invoke-SmokeTest.ps1 -KeepWorkFolder
#>
param([switch]$KeepWorkFolder)

$ErrorActionPreference = 'Stop'
$repo  = Split-Path -Parent $PSScriptRoot
$fx    = Join-Path $PSScriptRoot 'fixtures'
$mocks = Join-Path $PSScriptRoot 'mocks'
$exe   = (Get-Process -Id $PID).Path
$work  = Join-Path ([IO.Path]::GetTempPath()) ("scommig-smoke-" + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $work | Out-Null
Copy-Item (Join-Path $repo 'src\*.ps1') $work
Copy-Item (Join-Path $fx 'MigrationManifest.csv') $work

$sep = [IO.Path]::PathSeparator
$basePath = $env:PSModulePath
$env:SCOMMIG_TEST_SOURCE = Join-Path $fx 'source'
$env:SCOMMIG_TEST_STATE  = Join-Path $work 'mock-installed.json'

$script:failures = 0
function Assert([bool]$cond, [string]$what) {
    if ($cond) { Write-Host "  PASS  $what" -ForegroundColor Green }
    else       { Write-Host "  FAIL  $what" -ForegroundColor Red; $script:failures++ }
}
function Invoke-Isolated([string]$mockSet, [string]$script, [string[]]$arguments) {
    $env:PSModulePath = (Join-Path $mocks $mockSet) + $sep + $basePath
    Push-Location $work
    try {
        $out = & $exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $work $script) @arguments 2>&1 | Out-String -Width 250
        $code = $LASTEXITCODE
    }
    finally { Pop-Location; $env:PSModulePath = $basePath }
    $log = Join-Path $work ("smoke-" + ($script -replace '\.ps1$', '') + "-" + ($arguments -join '_' -replace '[^\w-]', '') + ".txt")
    $out | Set-Content -LiteralPath $log
    return [pscustomobject]@{ Code = $code; Output = $out }
}

try {
    Write-Host "Work folder: $work" -ForegroundColor Cyan

    Write-Host "`n[1] Source export" -ForegroundColor Cyan
    $r = Invoke-Isolated 'source' 'Export-ScomEnvironment.ps1' @('-Role', 'Source', '-Manifest', (Join-Path $work 'MigrationManifest.csv'), '-OutputFolder', (Join-Path $work 'Source'))
    Assert ($r.Code -eq 0) "Export -Role Source exits cleanly"
    foreach ($f in 'SourceInventory.csv', 'GroupConversion.csv', 'StaticGroupMembers.csv', 'OverrideInstances.csv', 'ManifestResolution.csv') {
        Assert (Test-Path (Join-Path $work "Source\$f")) "Source export wrote $f"
    }

    Write-Host "`n[2] Target export (fixture copy)" -ForegroundColor Cyan
    $tAll = Join-Path $work 'Target\AllMPs'
    New-Item -ItemType Directory -Path $tAll -Force | Out-Null
    Copy-Item (Join-Path $fx 'target\*.xml') $tAll

    Write-Host "`n[3] Steps" -ForegroundColor Cyan
    foreach ($step in 'Check', 'FixReview', 'EnableOverrides', 'MapInstances', 'Compile', 'WhyBlocked', 'OverrideReport', 'DryRun', 'TestGroups') {
        $r = Invoke-Isolated 'target' 'Invoke-ScomMigrationStep.ps1' @($step, '-ManagementServer', 'SCOMMS01')
        Assert ($r.Code -eq 0 -and $r.Output -notmatch 'Exception') "Step $step runs without error"
    }

    Write-Host "`n[4] Results" -ForegroundColor Cyan
    $compile = Join-Path $work 'Compile'
    $bim = @(Import-Csv (Join-Path $compile 'BatchImportManifest.csv'))
    $expected = [ordered]@{
        'Contoso.Company.Knowledge'        = 'READY'    # dead category stripped
        'Contoso.Custom.Overrides'         = 'READY'    # dead overrides stripped, per-object override re-pointed
        'Contoso.SQL.Server.Group.All'     = 'READY'    # static groups converted to dynamic
        'Contoso.Windows.Server.Overrides' = 'READY'
        'Contoso.EHR.Application'          = 'BLOCKED'  # unit monitor targets a retired SQL 2012 class
        'Contoso.Custom.Library'           = 'BLOCKED'  # sealed in source, only an .xml export supplied
        'Contoso.App2'                     = 'BLOCKED'  # depends on the blocked sealed library
        'Contoso.Rule.Removed'             = 'BLOCKED'  # rule targets a class removed from the target MP version
    }
    foreach ($k in $expected.Keys) {
        $row = $bim | Where-Object MPID -eq $k | Select-Object -First 1
        Assert ($row -and $row.Verdict -eq $expected[$k]) ("{0,-34} {1}" -f $k, $expected[$k])
    }

    $strip = @(Import-Csv (Join-Path $compile 'StrippedElements.csv'))
    foreach ($id in 'OV.DeadElement', 'OV.DeadMP', 'OV.Inst.Gone', 'Cat.Dead') {
        Assert (@($strip | Where-Object ElementID -eq $id).Count -ge 1) "Stripped: $id"
    }
    Assert (@($strip | Where-Object ElementID -eq 'OV.Valid').Count -eq 0) "Kept: OV.Valid"

    $ov = Get-Content (Join-Path $compile 'CandidateMPs\Contoso.Custom.Overrides.xml') -Raw
    Assert ($ov -match 'ID="OV.Inst.Mapped"' -and $ov -notmatch 'eeeeeeee-2222') "Per-object override re-pointed to the target object"

    $grp = Get-Content (Join-Path $compile 'CandidateMPs\Contoso.SQL.Server.Group.All.xml') -Raw
    Assert ($grp -notmatch '<IncludeList>')            "Group candidate has no static IncludeList left"
    Assert ($grp -match 'NetbiosComputerName')          "Group candidate matches on NetBIOS name"
    Assert ($grp -match '<HostProperty>')               "Hosted-class rule uses HostProperty"
    Assert ($grp -match 'APPSQL01\|APPSQL02')           "REVIEW rule narrowed to exact member names"

    $gcr = @(Import-Csv (Join-Path $compile 'GroupConversionResults.csv'))
    Assert (@($gcr | Where-Object Action -eq 'ConvertedToDynamic').Count -ge 3) "At least 3 membership rules converted to dynamic"

    $map = @(Import-Csv (Join-Path $work 'InstanceMap.csv'))
    Assert (@($map | Where-Object { $_.NewGuid }).Count -eq 1) "InstanceMap: 1 object found in target, 1 not"

    Write-Host "`n[5] ExportSource through PowerShell remoting (stand-in)" -ForegroundColor Cyan
    $rw = Join-Path $work 'remote'
    New-Item -ItemType Directory -Path $rw, (Join-Path $rw 'remote-temp'), (Join-Path $rw 'shares') -Force | Out-Null
    Copy-Item (Join-Path $repo 'src\*.ps1') $rw
    Copy-Item (Join-Path $fx 'MigrationManifest.csv') $rw
    # An original for the in-scope sealed MP that the source itself doesn't have: on a "share" searched from the target.
    Set-Content -LiteralPath (Join-Path $rw 'shares\Contoso.Custom.Library.mp') -Value 'fake sealed original'
    $env:SCOMMIG_TEST_REMOTE_TEMP = Join-Path $rw 'remote-temp'
    $loop = Join-Path $mocks 'remoting\Invoke-WithLoopback.ps1'
    function Invoke-Remote([hashtable]$scriptArgs) {
        $b64 = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes(($scriptArgs | ConvertTo-Json -Compress)))
        $env:PSModulePath = (Join-Path $mocks 'source') + $sep + $basePath
        Push-Location $rw
        try {
            $out = & $exe -NoProfile -ExecutionPolicy Bypass -File $loop -Script (Join-Path $rw 'Invoke-ScomMigrationStep.ps1') -ArgsB64 $b64 2>&1 | Out-String -Width 250
            $code = $LASTEXITCODE
        }
        finally { Pop-Location; $env:PSModulePath = $basePath }
        Add-Content -LiteralPath (Join-Path $rw 'smoke-remote.txt') -Value $out
        return [pscustomobject]@{ Code = $code; Output = $out }
    }

    $r = Invoke-Remote @{ Step = 'ExportSource'; SourceServer = 'OLDSCOM01'; SealedSearchPath = @((Join-Path $rw 'shares')) }
    Assert ($r.Code -eq 0) "ExportSource exits cleanly"
    $rs = Join-Path $rw 'Source'
    foreach ($f in 'SourceInventory.csv', 'GroupConversion.csv', 'StaticGroupMembers.csv', 'OverrideInstances.csv', 'ManifestResolution.csv', 'Export-Source.log') {
        Assert (Test-Path (Join-Path $rs $f)) "Copied back: $f"
    }
    $nLocal  = @(Get-ChildItem (Join-Path $work 'Source\AllMPs') -Filter *.xml).Count
    $nRemote = @(Get-ChildItem (Join-Path $rs 'AllMPs') -Filter *.xml).Count
    Assert ($nRemote -gt 0 -and $nRemote -eq $nLocal) "Same MPs exported as a local export ($nRemote)"
    Assert ((Get-Content (Join-Path $rs 'GroupConversion.csv') -Raw) -eq (Get-Content (Join-Path $work 'Source\GroupConversion.original.csv') -Raw)) "GroupConversion.csv identical to a local export"
    Assert (Test-Path (Join-Path $rs 'SealedOriginals\Contoso.Custom.Library.mp')) "Sealed original found on a share searched from the target"
    Assert (@(Import-Csv (Join-Path $rs 'SealedOriginalsMissing.csv') | Where-Object MPID -eq 'Contoso.Custom.Library').Count -eq 0) "...and no longer listed as missing"
    Assert (@(Import-Csv (Join-Path $rs 'SealedOriginalsFound.csv') | Where-Object MPID -eq 'Contoso.Custom.Library').Count -eq 1) "...and listed in SealedOriginalsFound.csv"
    Assert (@(Get-ChildItem (Join-Path $rw 'remote-temp') -Force).Count -eq 0) "Temporary folder on the source removed"
    $calls = Get-Content (Join-Path $rw 'remoting-calls.txt')
    Assert ($calls -contains 'New-PSSession OLDSCOM01' -and $calls -contains 'Remove-PSSession') "Session opened to the source and closed"

    $r = Invoke-Remote @{ Step = 'ExportSource'; SourceServer = 'OLDSCOM01' }
    Assert ($r.Code -eq 0) "Re-export exits cleanly"
    Assert (@(Get-ChildItem $rw -Directory -Filter 'Source.previous-*').Count -eq 1) "Previous export kept as Source.previous-*"
    Assert (Test-Path (Join-Path $rs 'SealedOriginals\Contoso.Custom.Library.mp')) "Original from the previous SealedOriginals picked up again"

    $r = Invoke-Remote @{ Step = 'ExportSource'; SourceServer = 'UNREACHABLE' }
    Assert ($r.Code -ne 0 -and $r.Output -match 'Test-WSMan') "Unreachable source stops with a Test-WSMan hint"
    $r = Invoke-Remote @{ Step = 'ExportSource' }
    Assert ($r.Code -ne 0 -and $r.Output -match 'SourceServer') "ExportSource without -SourceServer says what to set"
}
finally {
    $env:SCOMMIG_TEST_SOURCE = $null; $env:SCOMMIG_TEST_STATE = $null; $env:SCOMMIG_TEST_REMOTE_TEMP = $null
    if ($KeepWorkFolder -or $script:failures) { Write-Host "`nWork folder kept: $work" -ForegroundColor Yellow }
    else { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
if ($script:failures) { Write-Host "$($script:failures) assertion(s) FAILED" -ForegroundColor Red; exit 1 }
Write-Host "All assertions passed." -ForegroundColor Green
exit 0
