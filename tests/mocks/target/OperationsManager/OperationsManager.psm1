# Offline stand-in for the OperationsManager module on the TARGET management group (tests only).
# Import of Contoso.Company.Knowledge fails on purpose, to exercise error capture.
$script:conns = New-Object System.Collections.ArrayList
$script:store = @{}
$statePath = if ($env:SCOMMIG_TEST_STATE) { $env:SCOMMIG_TEST_STATE } else { Join-Path ([IO.Path]::GetTempPath()) 'scommig-mock-installed.json' }
function Load { if (Test-Path $statePath) { $o = Get-Content $statePath -Raw | ConvertFrom-Json; foreach ($p in $o.PSObject.Properties) { $script:store[$p.Name] = $p.Value } } }
function Save { $script:store | ConvertTo-Json -Depth 3 | Set-Content $statePath }
Load
function New-SCOMManagementGroupConnection { param([string]$ComputerName, $Credential)
  foreach ($c in $script:conns) { $c.IsActive = $false }
  [void]$script:conns.Add([pscustomobject]@{ ManagementServerName = "$ComputerName.contoso.local"; ManagementGroupName = "MG_$ComputerName"; IsActive = $true }) }
function Get-SCOMManagementGroupConnection { $script:conns }
function Set-SCOMManagementGroupConnection { param([Parameter(ValueFromPipeline)]$Connection) process { foreach ($c in $script:conns) { $c.IsActive = ($c -eq $Connection) } } }
function Get-SCOMManagementPack { param([string]$Name)
  $active = $script:conns | ? IsActive | select -First 1
  Write-Verbose "Get-SCOMManagementPack on $($active.ManagementGroupName)"
  $all = $script:store.Values | % { [pscustomobject]@{ Name=$_.Name; Version=[version]$_.Version; Sealed=[bool]$_.Sealed; KeyToken=$_.KeyToken } }
  if ($Name) { $all | ? Name -eq $Name } else { $all } }
function Import-SCOMManagementPack { param([string]$Fullname)
  [xml]$x = Get-Content $Fullname -Raw
  $id = $x.ManagementPack.Manifest.Identity.ID
  if ($id -eq 'Contoso.Company.Knowledge') {
    $inner = New-Object System.Exception 'Verification failed with 1 errors: Cannot find element Foo in MP Bar.'
    $mid = New-Object System.InvalidOperationException ('The management pack is not valid.', $inner)
    throw (New-Object System.Exception ('Import failed.', $mid)) }
  $script:store[$id] = @{ Name=$id; Version=$x.ManagementPack.Manifest.Identity.Version; Sealed=$false; KeyToken='' }; Save }
function Get-SCOMClass { param([string]$Name) [pscustomobject]@{Name=$Name} }
function Get-SCOMClassInstance { param($Class)
  if ($Class.Name -in 'Microsoft.Windows.Computer','Microsoft.Windows.Server.Computer') { 'APPSQL01','OTHER09' | % { [pscustomobject]@{Id=[guid]::NewGuid();DisplayName="$($_.ToLower()).contoso.com";FullName="Microsoft.Windows.Computer:$($_.ToLower()).contoso.com"} } ; return }
  if ($Class.Name -like 'Contoso.*Group') { $o=[pscustomobject]@{DisplayName=$Class.Name;FullName=$Class.Name}; $n = if ($Class.Name -eq 'Contoso.SQL.Group') {1} else {0}; $o | Add-Member ScriptMethod GetRelatedMonitoringObjects ([scriptblock]::Create($(if ($n) { "@(1)" } else { "@()" }))); return $o }
}
function Remove-SCOMManagementPack { param($ManagementPack) $script:store.Remove($ManagementPack.Name); Save }

Export-ModuleMember -Function *
