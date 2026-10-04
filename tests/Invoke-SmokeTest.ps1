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
}
finally {
    $env:SCOMMIG_TEST_SOURCE = $null; $env:SCOMMIG_TEST_STATE = $null
    if ($KeepWorkFolder -or $script:failures) { Write-Host "`nWork folder kept: $work" -ForegroundColor Yellow }
    else { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
}

Write-Host ""
if ($script:failures) { Write-Host "$($script:failures) assertion(s) FAILED" -ForegroundColor Red; exit 1 }
Write-Host "All assertions passed." -ForegroundColor Green
exit 0
