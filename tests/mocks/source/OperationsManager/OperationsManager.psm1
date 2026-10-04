# Offline stand-in for the OperationsManager module on the SOURCE management group (tests only).
# Set $env:SCOMMIG_TEST_SOURCE to tests/fixtures/source before importing.
$script:conns = New-Object System.Collections.ArrayList
function New-SCOMManagementGroupConnection { param([string]$ComputerName,$Credential) foreach ($c in $script:conns) { $c.IsActive=$false }; [void]$script:conns.Add([pscustomobject]@{ManagementServerName="$ComputerName.x";ManagementGroupName='SourceMG';IsActive=$true}) }
function Get-SCOMManagementGroupConnection { $script:conns }
function Get-SCOMManagementPack {
  Get-ChildItem (Join-Path $env:SCOMMIG_TEST_SOURCE '*.xml') | % { [xml]$x=Get-Content $_ -Raw; $id=$x.ManagementPack.Manifest.Identity.ID
    [pscustomobject]@{Name=$id;DisplayName="$id display";Version=[version]$x.ManagementPack.Manifest.Identity.Version;Sealed=($id -like 'Microsoft*' -or $id -eq 'Contoso.Custom.Library');KeyToken=$(if($id -like 'Microsoft*'){'31bf3856ad364e35'}else{$null});TimeCreated=(Get-Date);LastModified=(Get-Date);Path=$_.FullName} } }
function Export-SCOMManagementPack { param([Parameter(ValueFromPipeline)]$ManagementPack,[string]$Path) process { Copy-Item $ManagementPack.Path (Join-Path $Path "$($ManagementPack.Name).xml") } }
$script:inst = @{
 '11111111-2222-3333-4444-555555555555' = @{DisplayName='appsql01.contoso.com';Path='';Cls='Microsoft.Windows.Server.Computer';S=$false}
 '66666666-2222-3333-4444-555555555555' = @{DisplayName='appsql02.contoso.com';Path='';Cls='Microsoft.Windows.Server.Computer';S=$false}
 'aaaaaaaa-2222-3333-4444-555555555555' = @{DisplayName='Contoso Sub Group';Path='';Cls='Contoso.Sub.Group';S=$true}
 'bbbbbbbb-2222-3333-4444-555555555555' = @{DisplayName='Microsoft.Windows Server 2016';Path='';FullName='Microsoft.Windows.Server.2016.OperatingSystem:webapp07.contoso.com';Cls='Microsoft.Windows.Server.2016.OperatingSystem';S=$false}
 'dddddddd-2222-3333-4444-555555555555' = @{DisplayName='batch01.contoso.com';Path='';FullName='Microsoft.SystemCenter.HealthServiceWatcher:Microsoft.SystemCenter.AgentWatchersGroup;abc';Cls='Microsoft.SystemCenter.HealthServiceWatcher';S=$false}
 'eeeeeeee-2222-3333-4444-555555555555' = @{DisplayName='appsql01.contoso.com';Path='';FullName='Microsoft.Windows.Computer:appsql01.contoso.com';Cls='Microsoft.Windows.Server.Computer';S=$false}
 'ffffffff-2222-3333-4444-555555555555' = @{DisplayName='old-retired01.contoso.com';Path='';FullName='Microsoft.Windows.Computer:old-retired01.contoso.com';Cls='Microsoft.Windows.Server.Computer';S=$false}
 'cccccccc-2222-3333-4444-555555555555' = @{DisplayName='Microsoft.Windows Server 2016';Path='';FullName='Microsoft.Windows.Server.2016.OperatingSystem:webapp09.contoso.com';Cls='Microsoft.Windows.Server.2016.OperatingSystem';S=$false}
}
function New-Inst($id,$h){ $o=[pscustomobject]@{Id=[guid]$id;DisplayName=$h.DisplayName;Path=$h.Path;FullName=$(if($h.ContainsKey('FullName')){$h.FullName}else{"Microsoft.Windows.Computer:$($h.DisplayName)"})}; $c=[pscustomobject]@{Name=$h.Cls;Singleton=$h.S}; $o | Add-Member ScriptMethod GetLeastDerivedNonAbstractClass ([scriptblock]::Create("[pscustomobject]@{Name='$($h.Cls)';Singleton=`$$($h.S)}")); $o }
function Get-SCOMClass { param([string]$Name) [pscustomobject]@{Name=$Name} }
function Get-SCOMClassInstance { param([guid[]]$Id,$Class)
  if ($Id) { foreach ($g in $Id) { $k=([string]$g).ToLower(); if ($script:inst.ContainsKey($k)) { New-Inst $k $script:inst[$k] } } }
  else { 'APPSQL01','APPSQL02','APPSQL03','APPSQLSQL01','WEBAPP07','WEBAPP08','WEBAPP09','OTHER01' | % { [pscustomobject]@{Id=[guid]::NewGuid();DisplayName="$($_.ToLower()).contoso.com";Path='';FullName="Microsoft.Windows.Computer:$($_.ToLower()).contoso.com"} } } }

Export-ModuleMember -Function *
