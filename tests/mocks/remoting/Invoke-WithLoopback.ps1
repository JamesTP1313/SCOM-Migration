<#
    Runs a toolkit script with the remoting stand-in loaded (offline tests only).
    -ArgsB64 is a base64 UTF-8 JSON object of the script's named parameters.
    Writes the recorded remoting calls to $env:SCOMMIG_TEST_REMOTE_TEMP\..\remoting-calls.txt
#>
param([Parameter(Mandatory = $true)][string]$Script, [string]$ArgsB64)
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'RemotingLoopback.ps1')
$a = @{}
if ($ArgsB64) {
    $obj = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($ArgsB64)) | ConvertFrom-Json
    foreach ($pr in $obj.PSObject.Properties) {
        $v = $pr.Value
        if ($v -is [System.Array]) { $v = [string[]]$v }
        $a[$pr.Name] = $v
    }
}
try { & $Script @a; $code = $LASTEXITCODE }
finally {
    $callFile = Join-Path (Split-Path -Parent $env:SCOMMIG_TEST_REMOTE_TEMP) 'remoting-calls.txt'
    $global:RemotingCalls | Set-Content -LiteralPath $callFile
}
if ($code) { exit $code }
