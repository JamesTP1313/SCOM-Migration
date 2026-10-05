<#
.SYNOPSIS
    Offline test for src/Migrate-ScomNotifications.ps1. No SCOM needed.

.DESCRIPTION
    Runs every scenario in its own PowerShell process against a synthetic
    Notifications.Internal MP and stand-in SCOM cmdlets, then asserts on the
    files each step writes and on the stand-in management group's final state.

        pwsh ./tests/Invoke-NotificationsTest.ps1                # PowerShell 7, any OS
        powershell -File .\tests\Invoke-NotificationsTest.ps1    # Windows PowerShell 5.1

    Exit code 0 = all assertions passed.
#>
param(
    # Internal: run one scenario (the driver below calls itself with this).
    [string]$Scenario
)
$ErrorActionPreference = 'Stop'
$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$script  = Join-Path (Split-Path -Parent $here) 'src/Migrate-ScomNotifications.ps1'
$workAll = Join-Path ([System.IO.Path]::GetTempPath()) 'ScomNotificationsTest'

###########################################################################
# Driver: run each scenario in a fresh process, then assert
###########################################################################
if (-not $Scenario) {
    $exe = (Get-Process -Id $PID).Path
    $scenarios = 'happy', 'othermp', 'missed', 'blockedref', 'tamper', 'noconfirm', 'damaged'
    $failures = New-Object System.Collections.ArrayList
    $passes = 0
    if (Test-Path $workAll) { Remove-Item -Recurse -Force $workAll }
    New-Item -ItemType Directory -Path $workAll -Force | Out-Null

    function Assert([string]$Scen, [string]$What, [bool]$Ok) {
        if ($Ok) { $script:passes++; Write-Host "  PASS  $What" -ForegroundColor Green }
        else     { [void]$failures.Add("[$Scen] $What"); Write-Host "  FAIL  $What" -ForegroundColor Red }
    }
    function Get-Names($Value) { @($Value | Where-Object { $_ }) }
    function Get-Json($Path) { if (Test-Path -LiteralPath $Path) { Get-Content -LiteralPath $Path -Raw | ConvertFrom-Json } else { $null } }

    foreach ($s in $scenarios) {
        Write-Host "`n=== $s ===" -ForegroundColor Cyan
        $work = Join-Path $workAll $s
        & $exe -NoProfile -File $PSCommandPath -Scenario $s *> (Join-Path $workAll "$s.console.txt")
        $result  = Get-Json (Join-Path $work 'result.json')
        $prep    = Get-Json (Join-Path $work 'Prepare/PrepareSummary.json')
        $xmlPath = Join-Path $work 'Prepare/Microsoft.SystemCenter.Notifications.Internal.xml'
        Assert $s 'scenario ran to completion' ($null -ne $result)
        if ($null -eq $result) { continue }

        switch ($s) {
            'happy' {
                Assert $s 'first Prepare is READY' ($result.FirstPrepareStatus -eq 'READY')
                [xml]$x = Get-Content -LiteralPath $result.FirstPreparedFile -Raw
                $subRules = @($x.SelectNodes('//Rule') | Where-Object { $_.OuterXml -match 'SubscribedAlertProvider' })
                Assert $s 'all 3 subscription rules are Enabled=false' ($subRules.Count -eq 3 -and @($subRules | Where-Object { $_.GetAttribute('Enabled') -ne 'false' }).Count -eq 0)
                $helper = $x.SelectSingleNode("//Rule[@ID='Notifications.HelperRule']")
                Assert $s 'non-subscription rule left untouched' ($helper.GetAttribute('Enabled') -eq 'true')
                $raw = $x.OuterXml
                Assert $s 'per-server scope re-pointed by FullName' ($raw -match 'aaaaaaaa-aaaa' -and $raw -notmatch '33333333-3333')
                Assert $s 'MP version set above the target copy' ($x.ManagementPack.Manifest.Identity.Version -eq '10.25.10132.1')
                $refs = @($x.ManagementPack.Manifest.References.Reference)
                Assert $s 'unused missing reference removed' (@($refs | Where-Object { $_.ID -like '*InternetInformationServices.2008' }).Count -eq 0)
                Assert $s 'references set to target versions' (@($refs | Where-Object { $_.Version -ne '10.25.10132.0' }).Count -eq 0)
                Assert $s 'import happened' ($result.Imported)
                Assert $s 'nothing enabled right after import' ((Get-Names $result.EnabledAfterImport).Count -eq 0)
                $verify = @(Get-ChildItem (Join-Path $work 'Import') -Filter 'VerifyReport.*.csv' | Select-Object -First 1 | ForEach-Object { Import-Csv $_.FullName })
                Assert $s 'Verify reports nothing missing' ($verify.Count -gt 0 -and @($verify | Where-Object { $_.Status -in 'MISSING', 'DIFFERENT' }).Count -eq 0)
                Assert $s 'Enable -AllEnabledOnSource enabled only the clean one' (((Get-Names $result.EnabledAfterEnable) -join ',') -eq 'Ops - Critical')
                Assert $s 'Disable -All left nothing enabled' ((Get-Names $result.EnabledAfterDisable).Count -eq 0)
                Assert $s 'second Prepare blocks (target not empty)' ($result.SecondPrepareStatus -eq 'BLOCKED')
            }
            'othermp' {
                Assert $s 'subscription stored in another MP is only a warning' ($prep.Status -eq 'READY')
                $nim = @(Import-Csv (Join-Path $work 'Prepare/NotInThisMp.csv'))
                Assert $s 'it is listed in NotInThisMp.csv' ($nim.Count -eq 1 -and $nim[0].DisplayName -eq 'Stored Elsewhere')
                Assert $s 'import happened' ($result.Imported)
            }
            'missed' {
                Assert $s 'name in MP but not a recognised rule blocks' ($prep.Status -eq 'BLOCKED' -and (@($prep.Blockers) -join ' ') -match 'weren.t recognised')
                Assert $s 'nothing imported' (-not $result.Imported)
            }
            'blockedref' {
                Assert $s 'used reference missing on target blocks' ($prep.Status -eq 'BLOCKED' -and (@($prep.Blockers) -join ' ') -match 'Contoso\.Legacy\.Library')
                Assert $s 'nothing imported' (-not $result.Imported)
            }
            'tamper'    { Assert $s 'changed prepared file is refused' (-not $result.Imported) }
            'noconfirm' { Assert $s 'anything but YES imports nothing' (-not $result.Imported) }
            'damaged' {
                Assert $s 'damaged inventory stops Prepare' ($null -eq $prep)
                Assert $s 'nothing imported' (-not $result.Imported)
            }
        }
    }

    Write-Host ''
    if ($failures.Count -eq 0) { Write-Host "All $passes assertions passed." -ForegroundColor Green; exit 0 }
    Write-Host "$($failures.Count) assertion(s) failed:" -ForegroundColor Red
    $failures | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    Write-Host "Console output per scenario: $workAll" -ForegroundColor Yellow
    exit 1
}

###########################################################################
# Single scenario (child process)
###########################################################################
$work = Join-Path $workAll $Scenario
if (Test-Path $work) { Remove-Item -Recurse -Force $work }

$G_ClassWin  = '11111111-1111-1111-1111-111111111111'   # class, exists on both sides
$G_GroupGone = '22222222-2222-2222-2222-222222222222'   # group, source only
$G_InstSrv   = '33333333-3333-3333-3333-333333333333'   # server, re-pointed by FullName
$G_InstGone  = '44444444-4444-4444-4444-444444444444'   # server, gone on target
$G_Monitor   = '55555555-5555-5555-5555-555555555555'   # monitor, exists on both sides
$G_Internal  = '66666666-6666-6666-6666-666666666666'   # unresolvable internal id
$G_InstSrvT  = 'aaaaaaaa-aaaa-aaaa-aaaa-aaaaaaaaaaaa'   # the same server's id on target

$legacyRef = ''; $legacyUse = ''
if ($Scenario -eq 'blockedref') {
    $legacyRef = '<Reference Alias="Legacy"><ID>Contoso.Legacy.Library</ID><Version>1.0.0.0</Version><PublicKeyToken>0123456789abcdef</PublicKeyToken></Reference>'
    $legacyUse = '<Category>Legacy!Something</Category>'
}

$srcXml = @"
<?xml version="1.0" encoding="utf-8"?>
<ManagementPack ContentReadable="true" SchemaVersion="2.0" OriginalSchemaVersion="1.1">
  <Manifest>
    <Identity><ID>Microsoft.SystemCenter.Notifications.Internal</ID><Version>7.2.12345.27</Version></Identity>
    <Name>Notifications Internal Library</Name>
    <References>
      <Reference Alias="SystemCenter"><ID>Microsoft.SystemCenter.Library</ID><Version>7.0.8433.0</Version><PublicKeyToken>31bf3856ad364e35</PublicKeyToken></Reference>
      <Reference Alias="Notification"><ID>Microsoft.SystemCenter.Notifications.Library</ID><Version>7.2.11719.0</Version><PublicKeyToken>31bf3856ad364e35</PublicKeyToken></Reference>
      <Reference Alias="OldIIS"><ID>Microsoft.Windows.InternetInformationServices.2008</ID><Version>7.0.0.0</Version><PublicKeyToken>31bf3856ad364e35</PublicKeyToken></Reference>
      $legacyRef
    </References>
  </Manifest>
  <Monitoring>
    <Rules>
      <Rule ID="Subscription0a1b" Enabled="true" Target="SystemCenter!Microsoft.SystemCenter.AlertNotificationSubscriptionServer" ConfirmDelivery="false" Remotable="true" Priority="Normal" DiscardLevel="100">
        <Category>Notification</Category>$legacyUse
        <DataSources><DataSource ID="DS1" TypeID="SystemCenter!Microsoft.SystemCenter.SubscribedAlertProvider">
          <AlertChangedSubscriptionConfiguration><Criteria><Expression>Severity=2</Expression></Criteria>
          <MonitoringClassId>$G_ClassWin</MonitoringClassId><MonitoringObjectId>$G_InstSrv</MonitoringObjectId><RuleId>$G_Monitor</RuleId><X>$G_Internal</X>
          </AlertChangedSubscriptionConfiguration></DataSource></DataSources>
        <WriteActions><WriteAction ID="Smtp1" TypeID="Notification!Microsoft.SystemCenter.Notification.SmtpWriteAction"/></WriteActions>
      </Rule>
      <Rule ID="Subscription2c3d" Target="SystemCenter!Microsoft.SystemCenter.AlertNotificationSubscriptionServer" ConfirmDelivery="false" Remotable="true" Priority="Normal" DiscardLevel="100">
        <Category>Notification</Category>
        <DataSources><DataSource ID="DS1" TypeID="SystemCenter!Microsoft.SystemCenter.SubscribedAlertProvider">
          <AlertChangedSubscriptionConfiguration><MonitoringObjectGroupId>$G_GroupGone</MonitoringObjectGroupId><MonitoringObjectId>$G_InstGone</MonitoringObjectId></AlertChangedSubscriptionConfiguration></DataSource></DataSources>
      </Rule>
      <Rule ID="Subscription4e5f" Enabled="false" Target="SystemCenter!Microsoft.SystemCenter.AlertNotificationSubscriptionServer" ConfirmDelivery="false" Remotable="true" Priority="Normal" DiscardLevel="100">
        <Category>Notification</Category>
        <DataSources><DataSource ID="DS1" TypeID="SystemCenter!Microsoft.SystemCenter.SubscribedAlertProvider"><AlertChangedSubscriptionConfiguration/></DataSource></DataSources>
      </Rule>
      <Rule ID="Notifications.HelperRule" Enabled="true" Target="SystemCenter!Microsoft.SystemCenter.RootManagementServer" ConfirmDelivery="false" Remotable="true" Priority="Normal" DiscardLevel="100">
        <Category>Custom</Category><DataSources/>
      </Rule>
    </Rules>
  </Monitoring>
</ManagementPack>
"@

class SmtpNotificationAction { $Name; $DisplayName; $Endpoint; $From; $IsBodyHtml }
class CommandNotificationAction { $Name; $DisplayName; $ApplicationName; $CommandLine; $WorkingDirectory }
class FakeInstance { $Id; $FullName; $DisplayName; $ClassName; [object] GetLeastDerivedNonAbstractClass() { return [PSCustomObject]@{ Name = $this.ClassName } } }

function New-Sub($n, $dn, $en, $to) {
    [PSCustomObject]@{ Name = $n; DisplayName = $dn; Enabled = $en; ToRecipients = @([PSCustomObject]@{ Name = $to }); CcRecipients = @(); BccRecipients = @(); Actions = @([PSCustomObject]@{ DisplayName = 'Email Ops' }) }
}

$global:MockState = @{
    Source = @{
        Mg = 'CONTOSO-OLD'
        Mps = @{ 'Microsoft.SystemCenter.Notifications.Internal' = [PSCustomObject]@{ Name = 'Microsoft.SystemCenter.Notifications.Internal'; Version = [version]'7.2.12345.27'; KeyToken = $null } }
        Xml = $srcXml
        Subs = @((New-Sub 'Subscription0a1b' 'Ops - Critical' $true 'Ops Team'), (New-Sub 'Subscription2c3d' 'DBA - SQL Group' $true 'DBA Team'), (New-Sub 'Subscription4e5f' 'Old - Disabled' $false 'Ops Team'))
        Chan = @(
            [SmtpNotificationAction]@{ Name = 'Smtp.1'; DisplayName = 'Email Ops'; From = 'scom@contoso.com'; IsBodyHtml = $true; Endpoint = [PSCustomObject]@{ PrimaryServer = [PSCustomObject]@{ Address = 'smtp.contoso.com'; PortNumber = 25; AuthenticationType = 'WindowsIntegrated' } } },
            [CommandNotificationAction]@{ Name = 'Cmd.1'; DisplayName = 'Ticket Script'; ApplicationName = 'C:\Scripts\ticket.exe'; CommandLine = '-a'; WorkingDirectory = 'C:\Scripts' })
        Subr = @([PSCustomObject]@{ Name = 'Ops Team'; Devices = @([PSCustomObject]@{ Protocol = 'Smtp'; Address = 'ops@contoso.com' }); ScheduleEntries = @() },
                 [PSCustomObject]@{ Name = 'DBA Team'; Devices = @([PSCustomObject]@{ Protocol = 'Smtp'; Address = 'dba@contoso.com' }, [PSCustomObject]@{ Protocol = 'Sms'; Address = '5551234' }); ScheduleEntries = @('x') })
        Classes = @{ $G_ClassWin = 'Microsoft.Windows.Computer'; $G_GroupGone = 'Contoso.Group.SQL2008Servers' }
        Monitors = @{ $G_Monitor = 'Contoso.Some.Monitor' }
        Instances = @([FakeInstance]@{ Id = $G_InstSrv; FullName = 'Microsoft.Windows.Computer:srv01.contoso.com'; DisplayName = 'srv01'; ClassName = 'Microsoft.Windows.Computer' },
                      [FakeInstance]@{ Id = $G_InstGone; FullName = 'Microsoft.Windows.Computer:old99.contoso.com'; DisplayName = 'old99'; ClassName = 'Microsoft.Windows.Computer' })
    }
    Target = @{
        Mg = 'CONTOSO-NEW'
        Mps = @{
            'Microsoft.SystemCenter.Library' = [PSCustomObject]@{ Name = 'Microsoft.SystemCenter.Library'; Version = [version]'10.25.10132.0'; KeyToken = '31bf3856ad364e35' }
            'Microsoft.SystemCenter.Notifications.Library' = [PSCustomObject]@{ Name = 'Microsoft.SystemCenter.Notifications.Library'; Version = [version]'10.25.10132.0'; KeyToken = '31bf3856ad364e35' }
            'Microsoft.SystemCenter.Notifications.Internal' = [PSCustomObject]@{ Name = 'Microsoft.SystemCenter.Notifications.Internal'; Version = [version]'10.25.10132.0'; KeyToken = $null }
        }
        Xml = '<ManagementPack><Manifest><Identity><ID>Microsoft.SystemCenter.Notifications.Internal</ID><Version>10.25.10132.0</Version></Identity></Manifest></ManagementPack>'
        Subs = @(); Chan = @(); Subr = @()
        Classes = @{ $G_ClassWin = 'Microsoft.Windows.Computer' }
        Monitors = @{ $G_Monitor = 'Contoso.Some.Monitor' }
        Instances = @([FakeInstance]@{ Id = $G_InstSrvT; FullName = 'Microsoft.Windows.Computer:srv01.contoso.com'; DisplayName = 'srv01'; ClassName = 'Microsoft.Windows.Computer' })
    }
}
$global:Side = 'Source'; $global:Answer = 'YES'; $global:Imported = $false

# ---- stand-in OperationsManager cmdlets ----
function global:New-SCOMManagementGroupConnection { param($ComputerName) }
function global:Get-SCOMManagementGroupConnection { [PSCustomObject]@{ ManagementGroupName = $MockState[$Side].Mg; IsActive = $true } }
function global:Get-SCOMManagementPack { param($Name) if ($Name) { $MockState[$Side].Mps[$Name] } else { $MockState[$Side].Mps.Values } }
function global:Export-SCOMManagementPack { param($ManagementPack, $Path) Set-Content -LiteralPath (Join-Path $Path "$($ManagementPack.Name).xml") -Value $MockState[$Side].Xml -Encoding UTF8 }
function global:Import-SCOMManagementPack {
    param($Fullname)
    [xml]$d = Get-Content -LiteralPath $Fullname -Raw
    foreach ($r in $d.ManagementPack.Manifest.References.Reference) { if (-not $MockState[$Side].Mps.ContainsKey($r.ID)) { throw "Reference $($r.ID) not found" } }
    $global:Imported = $true
    $MockState[$Side].Mps['Microsoft.SystemCenter.Notifications.Internal'].Version = [version]$d.ManagementPack.Manifest.Identity.Version
    $MockState[$Side].Subs = @(foreach ($r in $d.SelectNodes('//Rule')) {
        if ($r.OuterXml -notmatch 'SubscribedAlertProvider') { continue }
        $src = $MockState.Source.Subs | Where-Object Name -eq $r.ID
        $en = if ($r.GetAttribute('Enabled')) { $r.GetAttribute('Enabled') -ne 'false' } else { $true }
        New-Sub $r.ID $src.DisplayName $en ($src.ToRecipients[0].Name)
    })
    $MockState[$Side].Chan = $MockState.Source.Chan; $MockState[$Side].Subr = $MockState.Source.Subr
}
function global:Get-SCOMNotificationSubscription { $MockState[$Side].Subs }
function global:Get-SCOMNotificationChannel { $MockState[$Side].Chan }
function global:Get-SCOMNotificationSubscriber { $MockState[$Side].Subr }
function global:Enable-SCOMNotificationSubscription { param($Subscription) $Subscription.Enabled = $true }
function global:Disable-SCOMNotificationSubscription { param($Subscription) $Subscription.Enabled = $false }
function global:Get-SCOMClass { param($Id, $Name)
    if ($Id) { $n = $MockState[$Side].Classes[$Id]; if ($n) { return [PSCustomObject]@{ Id = $Id; Name = $n; DisplayName = $n } } else { throw 'not found' } }
    if ($Name) { $k = $MockState[$Side].Classes.Keys | Where-Object { $MockState[$Side].Classes[$_] -eq $Name }; if ($k) { return [PSCustomObject]@{ Id = $k; Name = $Name } } } }
function global:Get-SCOMMonitor { param($Id) $n = $MockState[$Side].Monitors[$Id]; if ($n) { [PSCustomObject]@{ Name = $n; DisplayName = $n } } else { throw 'not found' } }
function global:Get-SCOMRule { param($Id) throw 'not found' }
function global:Get-SCOMClassInstance { param($Id, $Class)
    if ($Id) { $i = $MockState[$Side].Instances | Where-Object Id -eq $Id; if ($i) { return $i } else { throw 'not found' } }
    if ($Class) { $MockState[$Side].Instances | Where-Object ClassName -eq $Class.Name } }
function global:Read-Host { param($Prompt) $global:Answer }

function Run($Where, [hashtable]$P) { $global:Side = $Where; & $script -WorkFolder $work @P }
function Get-EnabledNames { @($MockState.Target.Subs | Where-Object { $_.Enabled } | ForEach-Object { $_.DisplayName }) }
function Get-PrepStatus { $f = Join-Path $work 'Prepare/PrepareSummary.json'; if (Test-Path $f) { (Get-Content $f -Raw | ConvertFrom-Json).Status } }

$result = [ordered]@{}

if ($Scenario -eq 'othermp') { $MockState.Source.Subs += (New-Sub 'SubscriptionZZ' 'Stored Elsewhere' $true 'Ops Team') }
if ($Scenario -eq 'missed')  { $MockState.Source.Subs += (New-Sub 'Smtp1' 'Referenced But Not A Rule' $true 'Ops Team') }

Run 'Source' @{ Step = 'Export' }

if ($Scenario -eq 'damaged') {
    # Simulate a hand-edit that emptied most of the inventory
    $sumFile = Join-Path $work 'Export/SourceSummary.json'
    $sum = Get-Content $sumFile -Raw | ConvertFrom-Json; $sum.Subscriptions = 135
    $sum | ConvertTo-Json | Set-Content $sumFile
    $csv = Join-Path $work 'Export/SourceSubscriptions.csv'
    (Get-Content $csv)[0..1] | Set-Content $csv
}

Run 'Target' @{ Step = 'Prepare' }
$result.FirstPrepareStatus = Get-PrepStatus
$result.FirstPreparedFile  = Join-Path $work 'Prepare/Microsoft.SystemCenter.Notifications.Internal.xml'

switch ($Scenario) {
    'tamper'    { Add-Content $result.FirstPreparedFile '<!-- edited -->'; Run 'Target' @{ Step = 'Import' } }
    'noconfirm' { $global:Answer = 'yes'; Run 'Target' @{ Step = 'Import' } }
    'happy' {
        # keep a copy of the first prepared file; the second Prepare overwrites it
        $copy = Join-Path $work 'FirstPrepared.xml'; Copy-Item $result.FirstPreparedFile $copy; $result.FirstPreparedFile = $copy
        Run 'Target' @{ Step = 'Import' }
        $result.EnabledAfterImport = Get-EnabledNames
        Run 'Target' @{ Step = 'Enable'; AllEnabledOnSource = $true }
        $result.EnabledAfterEnable = Get-EnabledNames
        Run 'Target' @{ Step = 'Disable'; All = $true }
        $result.EnabledAfterDisable = Get-EnabledNames
        Run 'Target' @{ Step = 'Prepare' }
        $result.SecondPrepareStatus = Get-PrepStatus
    }
    default { Run 'Target' @{ Step = 'Import' } }
}

$result.Imported = [bool]$global:Imported
$result | ConvertTo-Json -Depth 4 | Set-Content (Join-Path $work 'result.json')
