<#
    Stand-in for PowerShell remoting, for the offline tests only.
    Dot-source it: New-PSSession / Invoke-Command / Remove-PSSession / Copy-Item
    become global functions that run the "remote" side in this same process.
    $env:SCOMMIG_TEST_REMOTE_TEMP is used as the remote side's TEMP folder.
    Calls are recorded in $global:RemotingCalls.
#>
$global:RemotingCalls = New-Object System.Collections.ArrayList

function global:New-PSSession {
    [CmdletBinding()]
    param([string[]]$ComputerName, [System.Management.Automation.PSCredential]$Credential)
    [void]$global:RemotingCalls.Add("New-PSSession $ComputerName")
    if ($ComputerName -contains 'UNREACHABLE') { throw "Connecting to remote server UNREACHABLE failed: WinRM cannot complete the operation." }
    [PSCustomObject]@{ ComputerName = ($ComputerName -join ','); Id = 1; Open = $true }
}

function global:Remove-PSSession {
    [CmdletBinding()]
    param($Session)
    [void]$global:RemotingCalls.Add('Remove-PSSession')
    if ($Session) { $Session.Open = $false }
}

function global:Invoke-Command {
    [CmdletBinding()]
    param($Session, [scriptblock]$ScriptBlock, [object[]]$ArgumentList)
    if (-not $Session -or -not $Session.Open) { throw 'loopback: Invoke-Command needs an open -Session' }
    [void]$global:RemotingCalls.Add('Invoke-Command')
    $savedTemp = $env:TEMP
    $env:TEMP = $env:SCOMMIG_TEST_REMOTE_TEMP
    try {
        if ($ArgumentList) { & $ScriptBlock @ArgumentList } else { & $ScriptBlock }
    }
    finally { $env:TEMP = $savedTemp }
}

function global:Copy-Item {
    [CmdletBinding()]
    param(
        [Parameter(Position = 0, ValueFromPipeline = $true)][string[]]$Path,
        [string[]]$LiteralPath,
        [Parameter(Position = 1)][string]$Destination,
        $FromSession, $ToSession,
        [switch]$Recurse, [switch]$Force
    )
    process {
        if ($FromSession) { [void]$global:RemotingCalls.Add("Copy-Item -FromSession $Path") }
        if ($ToSession)   { [void]$global:RemotingCalls.Add("Copy-Item -ToSession $Path") }
        $p = @{}
        if ($Path) { $p['Path'] = $Path }
        if ($LiteralPath) { $p['LiteralPath'] = $LiteralPath }
        if ($Destination) { $p['Destination'] = $Destination }
        if ($Recurse) { $p['Recurse'] = $true }
        if ($Force) { $p['Force'] = $true }
        Microsoft.PowerShell.Management\Copy-Item @p -ErrorAction $ErrorActionPreference
    }
}
