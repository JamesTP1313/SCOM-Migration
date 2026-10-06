<#
.SYNOPSIS
    Migrate SCOM notification channels, subscribers and subscriptions from one
    management group to another by moving the Notifications.Internal MP.

.DESCRIPTION
    All notification config in a management group lives in one unsealed MP:
    Microsoft.SystemCenter.Notifications.Internal. Moving that MP carries
    everything as-is: criteria, class/group scope, To/CC/BCC, schedules,
    HTML bodies, SMTP port/auth, command channels.

    The TARGET must have no notification config yet (a fresh management group).
    Each step runs on ONE management group only, with that environment's own
    OperationsManager module (a newer console can't talk to an older group).

      Step      Where    Changes SCOM?   What it does
      --------  -------  -------------   ---------------------------------------
      Export    source   No              Exports the MP + inventory CSVs + resolves
                                         every GUID the subscriptions use.
      Prepare   target   No              Builds the import file: every subscription
                                         DISABLED, version above the target's copy,
                                         references set to the target's versions,
                                         per-server scope re-pointed by FullName.
                                         Writes READY or BLOCKED with reasons.
      Import    target   YES             Imports the prepared MP (type YES).
                                         Re-checks that nothing is enabled.
      Verify    target   No              Compares the target with the source inventory.
      Enable    target   YES             Enables named subscriptions (type YES).
      Disable   target   YES             Emergency stop: disables subscriptions.

.PARAMETER Step
    Export | Prepare | Import | Verify | Enable | Disable

.PARAMETER WorkFolder
    Root folder (default: C:\SCOMMigration\Notifications). Export writes
    <WorkFolder>\Export; copy that folder to the same path on the target
    management server before Prepare.

.PARAMETER ManagementServer
    Management server to connect to (default: this server).

.PARAMETER Name
    Enable/Disable: subscription display name(s). Wildcards allowed.

.PARAMETER AllEnabledOnSource
    Enable: every subscription that was enabled on the source and has no scope issues.

.PARAMETER All
    Disable: every subscription.

.PARAMETER Force
    Enable: also enable subscriptions that Prepare flagged with scope issues.

.EXAMPLE
    # On the source management server
    .\Migrate-ScomNotifications.ps1 -Step Export

    # On the target management server (after copying the Export folder)
    .\Migrate-ScomNotifications.ps1 -Step Prepare
    .\Migrate-ScomNotifications.ps1 -Step Import
    .\Migrate-ScomNotifications.ps1 -Step Enable -Name 'Ops - Critical'
    .\Migrate-ScomNotifications.ps1 -Step Disable -All      # emergency stop

.EXAMPLE
    # Everything from the target management server: the export runs on the
    # source through PowerShell remoting and lands in <WorkFolder>\Export here.
    .\Migrate-ScomNotifications.ps1 -Step Export -SourceServer OLDSCOM01
    .\Migrate-ScomNotifications.ps1 -Step Prepare

.NOTES
    Version 2.1.0. Used on a production SCOM 2016 -> 2025 migration.
    2.1.0: -SourceServer runs the export through PowerShell remoting.
    See docs/notifications.md.
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [ValidateSet('Export', 'Prepare', 'Import', 'Verify', 'Enable', 'Disable')]
    [string]$Step,

    [string]$WorkFolder = 'C:\SCOMMigration\Notifications',

    [string]$ManagementServer = 'localhost',

    [string[]]$Name,

    [switch]$AllEnabledOnSource,

    [switch]$All,

    [switch]$Force,

    # Export only: run the export ON this source management server through
    # PowerShell remoting (with the source's own OperationsManager module) and
    # copy the result into <WorkFolder>\Export here. Lets you run every step
    # from the target management server.
    [string]$SourceServer,

    # Export with -SourceServer: credential for the remoting session.
    [System.Management.Automation.PSCredential]$Credential,

    # Internal: set on the copy that runs on the source during a -SourceServer export.
    [Parameter(DontShow)][switch]$RemoteChild
)

#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
$ThisScriptPath = $PSCommandPath

# In the copy running on the source during a -SourceServer export, 'exit' would
# end the remoting session silently, so stop with an error instead.
function Stop-Script([string]$Message) {
    if ($RemoteChild) { throw $Message }
    exit 1
}

# Use the same work folder path on both servers.
if (-not $WorkFolder) { $WorkFolder = 'C:\SCOMMigration\Notifications' }
$WorkFolder = $WorkFolder.Trim().Trim('"').Trim("'")
$WorkFolder = [System.IO.Path]::GetFullPath($WorkFolder)
$driveRoot  = [System.IO.Path]::GetPathRoot($WorkFolder)
if (-not $driveRoot -or -not (Test-Path -LiteralPath $driveRoot)) {
    Write-Host "Work folder '$WorkFolder' is on drive '$driveRoot', which isn't available on this server." -ForegroundColor Red
    Write-Host "Pick a folder on a drive that exists, for example:  -WorkFolder C:\SCOMMigration\Notifications" -ForegroundColor Yellow
    Stop-Script "Work folder drive '$driveRoot' not available."
}
try {
    if (-not (Test-Path -LiteralPath $WorkFolder)) { New-Item -ItemType Directory -Path $WorkFolder -Force -ErrorAction Stop | Out-Null }
    $probe = Join-Path $WorkFolder '.write-test'
    Set-Content -LiteralPath $probe -Value 'ok' -ErrorAction Stop
    Remove-Item -LiteralPath $probe -Force -ErrorAction SilentlyContinue
}
catch {
    Write-Host "Can't write to work folder '$WorkFolder': $($_.Exception.Message)" -ForegroundColor Red
    Write-Host 'Run PowerShell as Administrator, or pick another folder with -WorkFolder.' -ForegroundColor Yellow
    Stop-Script "Can't write to work folder '$WorkFolder'."
}

Write-Host "Work folder: $WorkFolder" -ForegroundColor DarkGray

$InternalMpName = 'Microsoft.SystemCenter.Notifications.Internal'
$GuidPattern    = '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}'

$ExportDir  = Join-Path $WorkFolder 'Export'
$PrepareDir = Join-Path $WorkFolder 'Prepare'
$ImportDir  = Join-Path $WorkFolder 'Import'
$LogDir     = Join-Path $WorkFolder 'Logs'

###########################################################################
# Helpers
###########################################################################

function Initialize-Folder([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { New-Item -ItemType Directory -Path $Path -Force | Out-Null }
}

Initialize-Folder $WorkFolder
Initialize-Folder $LogDir
$RunStamp       = Get-Date -Format 'yyyyMMdd-HHmmss'
$NotifLogFile = Join-Path $LogDir "Notifications.$Step.$RunStamp.log"

function Write-Log {
    param([string]$Message, [ValidateSet('INFO', 'WARN', 'ERROR', 'OK')][string]$Level = 'INFO')
    $line = '[{0}] [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    Add-Content -LiteralPath $NotifLogFile -Value $line -Encoding UTF8
    switch ($Level) {
        'WARN'  { Write-Host $Message -ForegroundColor Yellow }
        'ERROR' { Write-Host $Message -ForegroundColor Red }
        'OK'    { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message }
    }
}

function Write-Banner([string]$Text) {
    $line = '=' * 72
    Write-Host ''
    Write-Host $line -ForegroundColor DarkCyan
    Write-Host $Text -ForegroundColor Cyan
    Write-Host $line -ForegroundColor DarkCyan
    Add-Content -LiteralPath $NotifLogFile -Value "`r`n$line`r`n$Text`r`n$line" -Encoding UTF8
}

# Safe property read for SDK objects under StrictMode. Accepts a dotted path.
# Returns $null instead of throwing when any part is missing.
function Get-Prop {
    param($Object, [string]$Path)
    $cur = $Object
    foreach ($part in $Path.Split('.')) {
        if ($null -eq $cur) { return $null }
        $p = $cur.PSObject.Properties[$part]
        if (-not $p) { return $null }
        try { $cur = $p.Value } catch { return $null }
    }
    return $cur
}

function Join-Names($Items) {
    $names = foreach ($i in @($Items)) {
        if ($null -eq $i) { continue }
        $n = Get-Prop $i 'DisplayName'
        if (-not $n) { $n = Get-Prop $i 'Name' }
        if (-not $n) { $n = [string]$i }
        $n
    }
    return (@($names) -join '; ')
}

function Get-ExceptionChain($ErrorRecord) {
    $msgs = @()
    $ex = $ErrorRecord.Exception
    while ($ex) { $msgs += ('{0}: {1}' -f $ex.GetType().FullName, $ex.Message); $ex = $ex.InnerException }
    return ($msgs -join "`r`n  -> ")
}

function Confirm-Yes([string]$Question) {
    Write-Host ''
    Write-Host $Question -ForegroundColor Yellow
    $answer = Read-Host 'Type YES to continue'
    if ($answer -cne 'YES') {
        Write-Log 'Not confirmed. Nothing was changed.' 'WARN'
        return $false
    }
    Write-Log 'Confirmed by operator.'
    return $true
}

function Connect-Scom {
    if (-not (Get-Command Get-SCOMManagementPack -ErrorAction SilentlyContinue)) {
        Import-Module OperationsManager -ErrorAction Stop
    }
    New-SCOMManagementGroupConnection -ComputerName $ManagementServer -ErrorAction Stop | Out-Null
    $conn = @(Get-SCOMManagementGroupConnection | Where-Object { Get-Prop $_ 'IsActive' }) | Select-Object -First 1
    $mgName = Get-Prop $conn 'ManagementGroupName'
    Write-Log "Connected to management group '$mgName' via $ManagementServer." 'OK'
    return $mgName
}

function Get-InternalMp {
    $mp = @(Get-SCOMManagementPack -Name $InternalMpName -ErrorAction SilentlyContinue)
    if ($mp.Count -eq 0) { return $null }
    return $mp[0]
}

# Subscription rules in the MP XML. A rule counts as a subscription if it uses
# the subscription data source/config, OR its ID matches a subscription Name
# recorded by the cmdlets at export time. Prepare blocks if the two disagree.
function Get-SubscriptionRules([xml]$Doc, [string[]]$KnownNames) {
    $rules = @($Doc.SelectNodes('//Monitoring/Rules/Rule'))
    $known = @{}
    foreach ($n in @($KnownNames)) { if ($n) { $known[$n] = $true } }
    return @($rules | Where-Object {
        $_.OuterXml -match 'SubscribedAlertProvider|AlertChangedSubscriptionConfiguration' -or
        $known.ContainsKey($_.GetAttribute('ID'))
    })
}

function Get-RuleEnabled($Rule) {
    $v = $Rule.GetAttribute('Enabled')
    if (-not $v) { return $true }   # attribute absent = enabled
    return ($v -ne 'false')
}

function Import-CsvRequired([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { throw "Required file not found: $Path. Run the earlier step first." }
    return @(Import-Csv -LiteralPath $Path)
}

###########################################################################
# STEP: EXPORT (source)
###########################################################################

function Invoke-Export {
    Write-Banner 'EXPORT (SOURCE) - read only'
    $mgName = Connect-Scom
    Initialize-Folder $ExportDir

    $mp = Get-InternalMp
    if (-not $mp) { throw "$InternalMpName not found on this management group. Is this the source management server?" }

    Export-SCOMManagementPack -ManagementPack $mp -Path $ExportDir -ErrorAction Stop
    $xmlPath = Join-Path $ExportDir "$InternalMpName.xml"
    if (-not (Test-Path -LiteralPath $xmlPath)) { throw "Export did not produce $xmlPath." }
    Write-Log "Exported $InternalMpName version $($mp.Version) -> $xmlPath" 'OK'

    # --- Inventory through the cmdlets (human-readable, and used to cross-check the XML)
    $subs = @(Get-SCOMNotificationSubscription -ErrorAction Stop)
    $subRows = foreach ($s in $subs) {
        [PSCustomObject]@{
            Name        = Get-Prop $s 'Name'
            DisplayName = Get-Prop $s 'DisplayName'
            Enabled     = Get-Prop $s 'Enabled'
            To          = Join-Names (Get-Prop $s 'ToRecipients')
            Cc          = Join-Names (Get-Prop $s 'CcRecipients')
            Bcc         = Join-Names (Get-Prop $s 'BccRecipients')
            Channels    = Join-Names (Get-Prop $s 'Actions')
        }
    }
    @($subRows) | Export-Csv -LiteralPath (Join-Path $ExportDir 'SourceSubscriptions.csv') -NoTypeInformation -Encoding UTF8

    $channels = @()
    try { $channels = @(Get-SCOMNotificationChannel -ErrorAction Stop) } catch { Write-Log "Could not list channels: $($_.Exception.Message)" 'WARN' }
    $chanRows = foreach ($c in $channels) {
        [PSCustomObject]@{
            Name             = Get-Prop $c 'Name'
            DisplayName      = Get-Prop $c 'DisplayName'
            Type             = $c.GetType().Name
            Server           = Get-Prop $c 'Endpoint.PrimaryServer.Address'
            Port             = Get-Prop $c 'Endpoint.PrimaryServer.PortNumber'
            Authentication   = [string](Get-Prop $c 'Endpoint.PrimaryServer.AuthenticationType')
            From             = Get-Prop $c 'From'
            IsBodyHtml       = Get-Prop $c 'IsBodyHtml'
            ApplicationName  = Get-Prop $c 'ApplicationName'
            CommandLine      = Get-Prop $c 'CommandLine'
            WorkingDirectory = Get-Prop $c 'WorkingDirectory'
        }
    }
    @($chanRows) | Export-Csv -LiteralPath (Join-Path $ExportDir 'SourceChannels.csv') -NoTypeInformation -Encoding UTF8

    $subscribers = @()
    try { $subscribers = @(Get-SCOMNotificationSubscriber -ErrorAction Stop) } catch { Write-Log "Could not list subscribers: $($_.Exception.Message)" 'WARN' }
    $subrRows = foreach ($r in $subscribers) {
        $devices = @(Get-Prop $r 'Devices')
        [PSCustomObject]@{
            Name          = Get-Prop $r 'Name'
            DeviceCount   = @($devices | Where-Object { $_ }).Count
            Devices       = (@($devices | Where-Object { $_ } | ForEach-Object { '{0}:{1}' -f (Get-Prop $_ 'Protocol'), (Get-Prop $_ 'Address') }) -join '; ')
            ScheduleCount = @(Get-Prop $r 'ScheduleEntries' | Where-Object { $_ }).Count
        }
    }
    @($subrRows) | Export-Csv -LiteralPath (Join-Path $ExportDir 'SourceSubscribers.csv') -NoTypeInformation -Encoding UTF8

    # --- Resolve every GUID the subscription rules use, while the source can still tell us what it is
    $doc = New-Object System.Xml.XmlDocument
    $doc.Load($xmlPath)
    $rules = @(Get-SubscriptionRules $doc @($subRows | ForEach-Object { $_.Name }))
    if ($rules.Count -ne $subs.Count) {
        Write-Log "The MP has $($rules.Count) subscription rule(s) but the cmdlets list $($subs.Count) subscription(s). Prepare will block until this is understood." 'WARN'
    }

    $usage = @{}
    foreach ($rule in $rules) {
        foreach ($m in [regex]::Matches($rule.OuterXml, $GuidPattern)) {
            $g = $m.Value.ToLowerInvariant()
            if (-not $usage.ContainsKey($g)) { $usage[$g] = [System.Collections.Generic.List[string]]::new() }
            if (-not $usage[$g].Contains($rule.GetAttribute('ID'))) { $usage[$g].Add($rule.GetAttribute('ID')) }
        }
    }

    $guidRows = foreach ($g in $usage.Keys) {
        $kind = 'Unresolved'; $elName = ''; $display = ''; $className = ''
        $o = $null
        try { $o = Get-SCOMClass -Id $g -ErrorAction Stop } catch { $o = $null }
        if ($o) { $kind = 'Class'; $elName = Get-Prop $o 'Name'; $display = Get-Prop $o 'DisplayName' }
        if (-not $o) {
            try { $o = Get-SCOMMonitor -Id $g -ErrorAction Stop } catch { $o = $null }
            if ($o) { $kind = 'Monitor'; $elName = Get-Prop $o 'Name'; $display = Get-Prop $o 'DisplayName' }
        }
        if (-not $o) {
            try { $o = Get-SCOMRule -Id $g -ErrorAction Stop } catch { $o = $null }
            if ($o) { $kind = 'Rule'; $elName = Get-Prop $o 'Name'; $display = Get-Prop $o 'DisplayName' }
        }
        if (-not $o) {
            try { $o = Get-SCOMClassInstance -Id $g -ErrorAction Stop } catch { $o = $null }
            if ($o) {
                $kind = 'Instance'; $elName = Get-Prop $o 'FullName'; $display = Get-Prop $o 'DisplayName'
                try { $className = $o.GetLeastDerivedNonAbstractClass().Name } catch { $className = '' }
            }
        }
        [PSCustomObject]@{
            Guid          = $g
            Kind          = $kind
            Name          = $elName
            DisplayName   = $display
            ClassName     = $className
            Subscriptions = ($usage[$g] -join '; ')
        }
    }
    @($guidRows) | Sort-Object Kind, Name | Export-Csv -LiteralPath (Join-Path $ExportDir 'SourceGuidMap.csv') -NoTypeInformation -Encoding UTF8

    [PSCustomObject]@{
        ManagementGroup = $mgName
        ExportedAt      = (Get-Date).ToString('s')
        MpVersion       = [string]$mp.Version
        Subscriptions   = $subs.Count
        SubscriptionRulesInXml = $rules.Count
        Channels        = $channels.Count
        Subscribers     = $subscribers.Count
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $ExportDir 'SourceSummary.json') -Encoding UTF8

    $resolved = @($guidRows | Where-Object { $_.Kind -ne 'Unresolved' })
    Write-Log ''
    Write-Log "Subscriptions : $($subs.Count) ($(@($subRows | Where-Object { $_.Enabled -eq $true }).Count) enabled)"
    Write-Log "Channels      : $($channels.Count)"
    Write-Log "Subscribers   : $($subscribers.Count)"
    Write-Log "Scope GUIDs   : $($resolved.Count) resolved ($(@($resolved | Where-Object Kind -eq 'Instance').Count) per-object)"
    Write-Log ''
    if (-not $RemoteChild) { Write-Log "Next: copy $ExportDir to the same path on the target management server, then run -Step Prepare there." 'OK' }
}

###########################################################################
# STEP: EXPORT through PowerShell remoting (run on the target)
###########################################################################
# The target's OperationsManager module can't connect to an older management
# group, so the export runs ON the source management server, with its own
# module, inside a remoting session. This script's text is sent as a script
# block (no file copy; the source's execution policy doesn't apply), writes to
# a temporary folder there, and the Export folder is copied back here. The
# temporary folder is removed afterwards; nothing else on the source changes.

function Invoke-RemoteExport {
    Write-Banner "EXPORT (SOURCE) via PowerShell remoting from $SourceServer - read only"
    $sessArgs = @{ ComputerName = $SourceServer; ErrorAction = 'Stop' }
    if ($Credential) { $sessArgs['Credential'] = $Credential }
    try { $session = New-PSSession @sessArgs }
    catch { throw "Can't open a PowerShell remoting session to '$SourceServer': $($_.Exception.Message)  Check with: Test-WSMan $SourceServer  (your account must be an administrator on that server)." }

    $remoteRoot = $null
    try {
        $remoteRoot = Invoke-Command -Session $session -ErrorAction Stop -ScriptBlock {
            $p = Join-Path $env:TEMP ('ScomNotificationsExport-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
            New-Item -ItemType Directory -Path $p -Force | Out-Null
            $p
        }
        Write-Log "Working folder on $($SourceServer): $remoteRoot"

        $code = [System.IO.File]::ReadAllText($ThisScriptPath)
        Invoke-Command -Session $session -ErrorAction Stop -ScriptBlock {
            param($Code, $Root)
            $childArgs = @{ Step = 'Export'; WorkFolder = [System.IO.Path]::Combine($Root, 'Work'); RemoteChild = $true }
            $sb = [scriptblock]::Create($Code)
            & $sb @childArgs
        } -ArgumentList $code, $remoteRoot

        if (Test-Path -LiteralPath $ExportDir) {
            $bk = "$ExportDir.previous-$RunStamp"
            Rename-Item -LiteralPath $ExportDir -NewName (Split-Path -Leaf $bk)
            Write-Log "Previous Export folder kept as $bk" 'WARN'
        }
        Copy-Item -FromSession $session -Path ([System.IO.Path]::Combine($remoteRoot, 'Work', 'Export')) -Destination $WorkFolder -Recurse -Force -ErrorAction Stop
        # The source's own log goes to Logs\FromSource-<time>\ (same file name pattern as ours).
        try { Copy-Item -FromSession $session -Path ([System.IO.Path]::Combine($remoteRoot, 'Work', 'Logs')) -Destination (Join-Path $LogDir "FromSource-$RunStamp") -Recurse -Force -ErrorAction Stop } catch { }
    }
    finally {
        if ($remoteRoot) {
            try { Invoke-Command -Session $session -ScriptBlock { param($p) Remove-Item -LiteralPath $p -Recurse -Force -ErrorAction SilentlyContinue } -ArgumentList $remoteRoot } catch { }
        }
        Remove-PSSession -Session $session -ErrorAction SilentlyContinue
    }

    $summaryPath = Join-Path $ExportDir 'SourceSummary.json'
    if (-not (Test-Path -LiteralPath $summaryPath)) { throw "The export on $SourceServer didn't complete (no SourceSummary.json came back). See the messages above." }
    $sum = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    Write-Log ''
    Write-Log "Copied back from $SourceServer ($($sum.ManagementGroup)): $($sum.Subscriptions) subscriptions, $($sum.Channels) channels, $($sum.Subscribers) subscribers." 'OK'
    Write-Log "Export folder: $ExportDir"
    Write-Log 'Next: run -Step Prepare on this server.' 'OK'
}

###########################################################################
# STEP: PREPARE (target, read only)
###########################################################################

function Invoke-Prepare {
    Write-Banner 'PREPARE (TARGET) - read only, nothing is imported'
    $null = Connect-Scom
    Initialize-Folder $PrepareDir

    $srcXmlPath = Join-Path $ExportDir "$InternalMpName.xml"
    if (-not (Test-Path -LiteralPath $srcXmlPath)) { throw "Source MP not found: $srcXmlPath. Copy the Export folder from the source first." }
    $srcSubs  = @(Import-CsvRequired (Join-Path $ExportDir 'SourceSubscriptions.csv'))
    $guidMap  = @(Import-CsvRequired (Join-Path $ExportDir 'SourceGuidMap.csv'))

    # The subscription inventory must match what Export recorded (it may have
    # been hand-edited to drop a subscription; anything more is a damaged file).
    $srcSummaryPath = Join-Path $ExportDir 'SourceSummary.json'
    if (Test-Path -LiteralPath $srcSummaryPath) {
        $srcSummary = Get-Content -LiteralPath $srcSummaryPath -Raw | ConvertFrom-Json
        $exported = [int]$srcSummary.Subscriptions
        if ($srcSubs.Count -ne $exported) {
            Write-Log "SourceSubscriptions.csv has $($srcSubs.Count) row(s); Export recorded $exported." 'WARN'
            if ($srcSubs.Count -lt ($exported - 5)) {
                throw "SourceSubscriptions.csv has only $($srcSubs.Count) of $exported subscriptions, so it looks damaged. Restore it from the Export copy (or its .bak) and re-run Prepare."
            }
        }
    }

    $blockers = [System.Collections.Generic.List[string]]::new()
    $warnings = [System.Collections.Generic.List[string]]::new()

    # --- 1. Target must be empty of notification config
    $tgtSubs = @(Get-SCOMNotificationSubscription -ErrorAction Stop)
    $tgtChan = @(); try { $tgtChan = @(Get-SCOMNotificationChannel -ErrorAction Stop) } catch { }
    $tgtSubr = @(); try { $tgtSubr = @(Get-SCOMNotificationSubscriber -ErrorAction Stop) } catch { }
    if ($tgtSubs.Count -gt 0 -or $tgtChan.Count -gt 0 -or $tgtSubr.Count -gt 0) {
        $blockers.Add("The target already has notification config ($($tgtSubs.Count) subscriptions, $($tgtChan.Count) channels, $($tgtSubr.Count) subscribers). Importing would replace it.")
    }

    $tgtMp = Get-InternalMp
    $tgtVersion = $null
    if ($tgtMp) {
        $tgtVersion = [string]$tgtMp.Version
        $backupDir = Join-Path $PrepareDir "TargetBackup.$RunStamp"
        Initialize-Folder $backupDir
        Export-SCOMManagementPack -ManagementPack $tgtMp -Path $backupDir -ErrorAction Stop
        Write-Log "The target already has $InternalMpName $tgtVersion (backed up to $backupDir)."
    }
    else {
        Write-Log "The target has no $InternalMpName yet; the import will create it."
    }

    # --- 2. Re-point per-object scope by FullName (text-level, before parsing)
    $raw = [System.IO.File]::ReadAllText($srcXmlPath)
    $guidResults = [System.Collections.Generic.List[object]]::new()
    $classCache = @{}
    $scopeIssuesBySub = @{}

    foreach ($row in $guidMap) {
        if ($row.Kind -eq 'Unresolved') { continue }
        $found = $null; $status = ''; $newGuid = ''
        switch ($row.Kind) {
            'Class'   { try { $found = Get-SCOMClass   -Id $row.Guid -ErrorAction Stop } catch { } }
            'Monitor' { try { $found = Get-SCOMMonitor -Id $row.Guid -ErrorAction Stop } catch { } }
            'Rule'    { try { $found = Get-SCOMRule    -Id $row.Guid -ErrorAction Stop } catch { } }
            'Instance'{ try { $found = Get-SCOMClassInstance -Id $row.Guid -ErrorAction Stop } catch { } }
        }
        if ($found) {
            $status = 'OK'
        }
        elseif ($row.Kind -eq 'Instance' -and $row.ClassName -and $row.Name) {
            if (-not $classCache.ContainsKey($row.ClassName)) {
                $cls = $null
                try { $cls = @(Get-SCOMClass -Name $row.ClassName -ErrorAction Stop) | Select-Object -First 1 } catch { }
                $classCache[$row.ClassName] = if ($cls) { @(Get-SCOMClassInstance -Class $cls -ErrorAction SilentlyContinue) } else { $null }
            }
            $pool = $classCache[$row.ClassName]
            if ($null -eq $pool) {
                $status = 'MISSING (class not in the target)'
            }
            else {
                $cands = @($pool | Where-Object { (Get-Prop $_ 'FullName') -eq $row.Name })
                if ($cands.Count -eq 1) {
                    $newGuid = ([string](Get-Prop $cands[0] 'Id')).ToLowerInvariant()
                    $raw = [regex]::Replace($raw, [regex]::Escape($row.Guid), $newGuid, 'IgnoreCase')
                    $status = 'REMAPPED'
                }
                elseif ($cands.Count -gt 1) { $status = 'AMBIGUOUS (several objects with this FullName)' }
                else { $status = 'MISSING (object not in the target)' }
            }
        }
        else {
            $status = "MISSING ($($row.Kind) not in the target)"
        }

        if ($status -notin @('OK', 'REMAPPED')) {
            foreach ($s in ($row.Subscriptions -split ';\s*' | Where-Object { $_ })) {
                if (-not $scopeIssuesBySub.ContainsKey($s)) { $scopeIssuesBySub[$s] = [System.Collections.Generic.List[string]]::new() }
                $scopeIssuesBySub[$s].Add("$($row.Kind) '$($row.Name)': $status")
            }
        }
        $guidResults.Add([PSCustomObject]@{
            Kind = $row.Kind; Name = $row.Name; ClassName = $row.ClassName
            SourceGuid = $row.Guid; TargetGuid = $newGuid; Status = $status; Subscriptions = $row.Subscriptions
        })
    }
    @($guidResults) | Export-Csv -LiteralPath (Join-Path $PrepareDir 'ScopeCheck.csv') -NoTypeInformation -Encoding UTF8

    $doc = New-Object System.Xml.XmlDocument
    $doc.PreserveWhitespace = $true
    $doc.LoadXml($raw)

    $idNode = $doc.SelectSingleNode('/ManagementPack/Manifest/Identity/ID')
    if (-not $idNode -or $idNode.InnerText -ne $InternalMpName) { throw "The exported file is not $InternalMpName." }

    # --- 3. Version above whatever the target has
    $verNode = $doc.SelectSingleNode('/ManagementPack/Manifest/Identity/Version')
    $srcVersion = [version]$verNode.InnerText
    $newVersion = $srcVersion
    if ($tgtVersion) {
        $t = [version]$tgtVersion
        if ($t -ge $srcVersion) {
            $build = [Math]::Max($t.Build, 0); $rev = [Math]::Max($t.Revision, 0)
            $newVersion = New-Object System.Version $t.Major, $t.Minor, $build, ($rev + 1)
        }
    }
    $verNode.InnerText = $newVersion.ToString()
    Write-Log "MP version: source $srcVersion -> import as $newVersion"

    # --- 4. References -> versions installed on the target
    $tgtMps = @{}
    foreach ($m in @(Get-SCOMManagementPack -ErrorAction Stop)) { $tgtMps[[string]$m.Name] = $m }

    $refRows = [System.Collections.Generic.List[object]]::new()
    $refsNode = $doc.SelectSingleNode('/ManagementPack/Manifest/References')
    foreach ($ref in @($doc.SelectNodes('/ManagementPack/Manifest/References/Reference'))) {
        $alias = $ref.GetAttribute('Alias')
        $refId = $ref.SelectSingleNode('ID').InnerText
        $refVerNode = $ref.SelectSingleNode('Version')
        $keyNode = $ref.SelectSingleNode('PublicKeyToken')
        $oldVer = $refVerNode.InnerText
        $used = [regex]::IsMatch($doc.OuterXml, '[">\s]' + [regex]::Escape($alias) + '!')

        if ($tgtMps.ContainsKey($refId)) {
            $tm = $tgtMps[$refId]
            $refVerNode.InnerText = [string]$tm.Version
            $note = ''
            $tKey = [string](Get-Prop $tm 'KeyToken')
            if ($keyNode -and $tKey -and $keyNode.InnerText -ne $tKey) {
                $note = "KeyToken changed $($keyNode.InnerText) -> $tKey"
                $keyNode.InnerText = $tKey
                $warnings.Add("Reference $refId has a different signing key on the target.")
            }
            $refRows.Add([PSCustomObject]@{ Alias = $alias; MP = $refId; SourceVersion = $oldVer; TargetVersion = [string]$tm.Version; Used = $used; Status = 'OK'; Note = $note })
        }
        elseif (-not $used) {
            [void]$refsNode.RemoveChild($ref)
            $refRows.Add([PSCustomObject]@{ Alias = $alias; MP = $refId; SourceVersion = $oldVer; TargetVersion = ''; Used = $false; Status = 'REMOVED (unused, not in the target)'; Note = '' })
        }
        else {
            $blockers.Add("Reference $refId (alias $alias) is used but not installed on the target.")
            $refRows.Add([PSCustomObject]@{ Alias = $alias; MP = $refId; SourceVersion = $oldVer; TargetVersion = ''; Used = $true; Status = 'MISSING - BLOCKS IMPORT'; Note = '' })
        }
    }
    @($refRows) | Export-Csv -LiteralPath (Join-Path $PrepareDir 'ReferenceCheck.csv') -NoTypeInformation -Encoding UTF8

    # --- 5. Disable every subscription rule
    $rules = @(Get-SubscriptionRules $doc @($srcSubs | ForEach-Object { $_.Name }))
    $ruleIds = @{}
    foreach ($r in $rules) { $ruleIds[$r.GetAttribute('ID')] = $r }

    # Subscriptions the source cmdlets list but that have no rule here. Two cases:
    #  - Their internal Name appears nowhere in this MP: they are stored in a
    #    different MP. This import doesn't carry them, so they can't arrive
    #    enabled. Warn and list them; they need migrating separately.
    #  - Their Name DOES appear in this MP but wasn't recognised as a rule:
    #    detection missed something, so we can't guarantee it's disabled. Block.
    $notInXml   = @($srcSubs | Where-Object { -not $ruleIds.ContainsKey($_.Name) })
    $docText    = $doc.OuterXml
    $otherMp    = @($notInXml | Where-Object { $_.Name -and $docText.IndexOf($_.Name, [System.StringComparison]::OrdinalIgnoreCase) -lt 0 })
    $unexplained = @($notInXml | Where-Object { $otherMp -notcontains $_ })

    if ($otherMp.Count -gt 0) {
        @($otherMp | Select-Object Name, DisplayName, Enabled, To, Cc, Bcc, Channels) |
            Export-Csv -LiteralPath (Join-Path $PrepareDir 'NotInThisMp.csv') -NoTypeInformation -Encoding UTF8
        $warnings.Add("$($otherMp.Count) subscription(s) are stored in a different MP and are NOT part of this import (see NotInThisMp.csv): $(($otherMp | ForEach-Object { $_.DisplayName }) -join ', ')")
    }
    if ($unexplained.Count -gt 0) {
        $blockers.Add("$($unexplained.Count) subscription(s) are referenced in this MP but weren't recognised as subscription rules, so they can't be guaranteed disabled: $(($unexplained | ForEach-Object { $_.DisplayName }) -join ', ')")
    }
    $expected = $srcSubs.Count - $otherMp.Count
    if ($rules.Count -ne $expected) {
        $blockers.Add("The MP has $($rules.Count) subscription rule(s) but $expected were expected (the source listed $($srcSubs.Count), $($otherMp.Count) stored in other MPs).")
    }

    $srcByName = @{}
    foreach ($s in $srcSubs) { $srcByName[$s.Name] = $s }

    $subRows = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $rules) {
        $id = $r.GetAttribute('ID')
        $wasEnabled = Get-RuleEnabled $r
        $r.SetAttribute('Enabled', 'false')
        $src = $srcByName[$id]
        $issues = if ($scopeIssuesBySub.ContainsKey($id)) { ($scopeIssuesBySub[$id] -join ' | ') } else { '' }
        $subRows.Add([PSCustomObject]@{
            Name          = $id
            DisplayName   = if ($src) { $src.DisplayName } else { $id }
            EnabledOnSource = $wasEnabled
            To            = if ($src) { $src.To } else { '' }
            Cc            = if ($src) { $src.Cc } else { '' }
            Bcc           = if ($src) { $src.Bcc } else { '' }
            Channels      = if ($src) { $src.Channels } else { '' }
            ScopeIssues   = $issues
            EnableAllowed = (-not $issues)
        })
    }

    # Belt and braces: nothing in the output may be enabled
    $stillEnabled = @(Get-SubscriptionRules $doc @($ruleIds.Keys) | Where-Object { Get-RuleEnabled $_ })
    if ($stillEnabled.Count -gt 0) { $blockers.Add("$($stillEnabled.Count) subscription rule(s) are still enabled after preparation.") }

    @($subRows) | Sort-Object DisplayName | Export-Csv -LiteralPath (Join-Path $PrepareDir 'PreparedSubscriptions.csv') -NoTypeInformation -Encoding UTF8

    # --- 6. Write, re-read, hash
    $outPath = Join-Path $PrepareDir "$InternalMpName.xml"
    $doc.Save($outPath)
    $check = New-Object System.Xml.XmlDocument
    $check.Load($outPath)   # throws if the output isn't well-formed
    $hash = (Get-FileHash -LiteralPath $outPath -Algorithm SHA256).Hash

    $status = if ($blockers.Count -eq 0) { 'READY' } else { 'BLOCKED' }
    [PSCustomObject]@{
        Status          = $status
        PreparedAt      = (Get-Date).ToString('s')
        PreparedFile    = $outPath
        Sha256          = $hash
        SourceVersion   = $srcVersion.ToString()
        TargetVersion   = $tgtVersion
        ImportVersion   = $newVersion.ToString()
        Subscriptions   = $rules.Count
        Blockers        = @($blockers)
        Warnings        = @($warnings)
    } | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $PrepareDir 'PrepareSummary.json') -Encoding UTF8

    $flagged = @($subRows | Where-Object { -not $_.EnableAllowed })
    Write-Banner "BOTTOM LINE: $status"
    Write-Log "Subscriptions        : $($rules.Count) (all will import DISABLED; $(@($subRows | Where-Object { $_.EnabledOnSource }).Count) were enabled on the source)"
    Write-Log "Scope re-pointed     : $(@($guidResults | Where-Object Status -eq 'REMAPPED').Count) per-object GUID(s)"
    Write-Log "Scope issues         : $($flagged.Count) subscription(s) -> PreparedSubscriptions.csv / ScopeCheck.csv" $(if ($flagged.Count) { 'WARN' } else { 'INFO' })
    Write-Log "References           : $(@($refRows | Where-Object Status -eq 'OK').Count) OK, $(@($refRows | Where-Object Status -like 'REMOVED*').Count) removed, $(@($refRows | Where-Object Status -like 'MISSING*').Count) missing"
    foreach ($w in $warnings) { Write-Log "WARNING: $w" 'WARN' }
    foreach ($b in $blockers) { Write-Log "BLOCKED: $b" 'ERROR' }
    if ($status -eq 'READY') { Write-Log 'Next: review PreparedSubscriptions.csv, then run -Step Import.' 'OK' }
}

###########################################################################
# STEP: IMPORT (target)
###########################################################################

function Invoke-Import {
    Write-Banner 'IMPORT (TARGET) - changes the management group'
    $null = Connect-Scom
    Initialize-Folder $ImportDir

    $summaryPath = Join-Path $PrepareDir 'PrepareSummary.json'
    if (-not (Test-Path -LiteralPath $summaryPath)) { throw 'Run -Step Prepare first.' }
    $summary = Get-Content -LiteralPath $summaryPath -Raw | ConvertFrom-Json
    if ($summary.Status -ne 'READY') { throw "Prepare status is $($summary.Status). Fix the blockers in PrepareSummary.json and re-run Prepare." }
    if (-not (Test-Path -LiteralPath $summary.PreparedFile)) { throw "Prepared file missing: $($summary.PreparedFile)" }
    $hash = (Get-FileHash -LiteralPath $summary.PreparedFile -Algorithm SHA256).Hash
    if ($hash -ne $summary.Sha256) { throw 'The prepared MP has changed since Prepare ran. Re-run -Step Prepare.' }

    # State may have changed since Prepare
    $existing = @(Get-SCOMNotificationSubscription -ErrorAction Stop)
    if ($existing.Count -gt 0) { throw "The target now has $($existing.Count) subscription(s). Re-run Prepare." }
    $tgtMp = Get-InternalMp
    if ($tgtMp -and [string]$tgtMp.Version -ne [string]$summary.TargetVersion) { throw "The target's $InternalMpName is now $($tgtMp.Version), not $($summary.TargetVersion). Re-run Prepare." }

    Write-Log "File          : $($summary.PreparedFile)"
    Write-Log "Version       : $($summary.ImportVersion)"
    Write-Log "Subscriptions : $($summary.Subscriptions) (all DISABLED)"
    if (-not (Confirm-Yes "Import $InternalMpName into this management group?")) { return }

    try {
        Import-SCOMManagementPack -Fullname $summary.PreparedFile -ErrorAction Stop
        Write-Log 'Import completed.' 'OK'
    }
    catch {
        $chain = Get-ExceptionChain $_
        $errFile = Join-Path $ImportDir "ImportErrors.$RunStamp.txt"
        Set-Content -LiteralPath $errFile -Value $chain -Encoding UTF8
        Write-Log "IMPORT FAILED. Nothing changed. Full error: $errFile" 'ERROR'
        Write-Log $chain 'ERROR'
        return
    }

    # Safety net: nothing should be enabled. If anything is, turn it off now.
    $enabled = @(Get-SCOMNotificationSubscription -ErrorAction Stop | Where-Object { Get-Prop $_ 'Enabled' })
    if ($enabled.Count -gt 0) {
        Write-Log "$($enabled.Count) subscription(s) came in ENABLED. Disabling them now." 'ERROR'
        foreach ($s in $enabled) { Disable-SCOMNotificationSubscription -Subscription $s -ErrorAction Stop }
        Write-Log 'Disabled.' 'OK'
    }

    Invoke-Verify -SkipConnect
}

###########################################################################
# STEP: VERIFY (target)
###########################################################################

function Invoke-Verify {
    param([switch]$SkipConnect)
    if (-not $SkipConnect) { Write-Banner 'VERIFY (TARGET) - read only'; $null = Connect-Scom } else { Write-Banner 'VERIFY' }
    Initialize-Folder $ImportDir

    $prep    = @(Import-CsvRequired (Join-Path $PrepareDir 'PreparedSubscriptions.csv'))
    $srcChan = @(Import-CsvRequired (Join-Path $ExportDir 'SourceChannels.csv'))
    $srcSubr = @(Import-CsvRequired (Join-Path $ExportDir 'SourceSubscribers.csv'))

    $rows = [System.Collections.Generic.List[object]]::new()
    function Add-Row($Type, $Name, $Status, $Detail) { $rows.Add([PSCustomObject]@{ Type = $Type; Name = $Name; Status = $Status; Detail = $Detail }) }

    $tSubs = @(Get-SCOMNotificationSubscription -ErrorAction Stop)
    $tByName = @{}; foreach ($s in $tSubs) { $tByName[[string](Get-Prop $s 'Name')] = $s }
    foreach ($p in $prep) {
        if (-not $tByName.ContainsKey($p.Name)) { Add-Row 'Subscription' $p.DisplayName 'MISSING' 'Not found on the target'; continue }
        $t = $tByName[$p.Name]
        $state = if (Get-Prop $t 'Enabled') { 'Enabled' } else { 'Disabled' }
        $detail = "To: $(Join-Names (Get-Prop $t 'ToRecipients')) | Channels: $(Join-Names (Get-Prop $t 'Actions'))"
        if ($p.ScopeIssues) { $detail += " | SCOPE ISSUES: $($p.ScopeIssues)" }
        Add-Row 'Subscription' $p.DisplayName "OK ($state)" $detail
    }

    $tChan = @(); try { $tChan = @(Get-SCOMNotificationChannel -ErrorAction Stop) } catch { }
    $tChanNames = @{}; foreach ($c in $tChan) { $tChanNames[[string](Get-Prop $c 'DisplayName')] = $c }
    foreach ($c in $srcChan) {
        if (-not $tChanNames.ContainsKey($c.DisplayName)) { Add-Row 'Channel' $c.DisplayName 'MISSING' $c.Type; continue }
        $note = $c.Type
        if ($c.Type -like 'Smtp*' -and $c.Authentication -and $c.Authentication -notmatch 'Anonymous') {
            $note += " | $($c.Authentication) auth: associate the Notification Action Account RunAs on the target"
        }
        if ($c.Type -like 'Command*' -and $c.ApplicationName) {
            $exists = Test-Path -LiteralPath $c.ApplicationName
            $note += " | $($c.ApplicationName) exists on this server: $exists (check every server in the Notifications resource pool)"
        }
        Add-Row 'Channel' $c.DisplayName 'OK' $note
    }

    $tSubr = @(); try { $tSubr = @(Get-SCOMNotificationSubscriber -ErrorAction Stop) } catch { }
    $tSubrNames = @{}; foreach ($r in $tSubr) { $tSubrNames[[string](Get-Prop $r 'Name')] = $r }
    foreach ($r in $srcSubr) {
        if ($tSubrNames.ContainsKey($r.Name)) {
            $count = @(Get-Prop $tSubrNames[$r.Name] 'Devices' | Where-Object { $_ }).Count
            $st = if ([string]$count -eq [string]$r.DeviceCount) { 'OK' } else { 'DIFFERENT' }
            Add-Row 'Subscriber' $r.Name $st "Devices source=$($r.DeviceCount) target=$count"
        }
        else { Add-Row 'Subscriber' $r.Name 'MISSING' '' }
    }

    $out = Join-Path $ImportDir "VerifyReport.$RunStamp.csv"
    @($rows) | Export-Csv -LiteralPath $out -NoTypeInformation -Encoding UTF8

    $bad = @($rows | Where-Object { $_.Status -in @('MISSING', 'DIFFERENT') })
    $enabledNow = @($tSubs | Where-Object { Get-Prop $_ 'Enabled' }).Count
    Write-Log "Subscriptions on target : $($tSubs.Count) ($enabledNow enabled)"
    Write-Log "Channels on target      : $($tChan.Count) (source: $($srcChan.Count))"
    Write-Log "Subscribers on target   : $($tSubr.Count) (source: $($srcSubr.Count))"
    Write-Log "Problems              : $($bad.Count)" $(if ($bad.Count) { 'ERROR' } else { 'OK' })
    foreach ($b in $bad) { Write-Log "  [$($b.Type)] $($b.Name): $($b.Status) $($b.Detail)" 'ERROR' }
    Write-Log "Report: $out"
}

###########################################################################
# STEP: ENABLE / DISABLE (target)
###########################################################################

function Select-Subscriptions([string[]]$Patterns) {
    $all = @(Get-SCOMNotificationSubscription -ErrorAction Stop)
    $picked = @($all | Where-Object {
        $dn = [string](Get-Prop $_ 'DisplayName')
        @($Patterns | Where-Object { $dn -like $_ }).Count -gt 0
    })
    foreach ($p in $Patterns) {
        if (-not @($all | Where-Object { [string](Get-Prop $_ 'DisplayName') -like $p }).Count) { Write-Log "No subscription matches '$p'." 'WARN' }
    }
    return $picked
}

function Invoke-Enable {
    Write-Banner 'ENABLE SUBSCRIPTIONS (TARGET)'
    if (-not $Name -and -not $AllEnabledOnSource) { throw 'Give -Name <display name> (wildcards OK) or -AllEnabledOnSource.' }
    $null = Connect-Scom
    $prep = @(Import-CsvRequired (Join-Path $PrepareDir 'PreparedSubscriptions.csv'))
    $prepByName = @{}; foreach ($p in $prep) { $prepByName[$p.Name] = $p }

    if ($AllEnabledOnSource) {
        $wanted = @($prep | Where-Object { [string](Get-Prop $_ 'EnabledOnSource') -eq 'True' -or [string](Get-Prop $_ 'EnabledOn2016') -eq 'True' } | ForEach-Object { $_.Name })
        $targets = @(Get-SCOMNotificationSubscription -ErrorAction Stop | Where-Object { $wanted -contains [string](Get-Prop $_ 'Name') })
    }
    else {
        $targets = @(Select-Subscriptions $Name)
    }

    $go = [System.Collections.Generic.List[object]]::new()
    foreach ($t in $targets) {
        $n = [string](Get-Prop $t 'Name'); $dn = [string](Get-Prop $t 'DisplayName')
        $p = $prepByName[$n]
        if (Get-Prop $t 'Enabled') { Write-Log "Already enabled: $dn"; continue }
        if ($p -and $p.EnableAllowed -ne 'True' -and -not $Force) {
            Write-Log "SKIPPED '$dn': scope issues ($($p.ScopeIssues)). Fix the scope in the console, or use -Force." 'WARN'
            continue
        }
        $go.Add($t)
    }
    if ($go.Count -eq 0) { Write-Log 'Nothing to enable.'; return }

    Write-Log "Will enable $($go.Count) subscription(s):"
    foreach ($t in $go) { Write-Log "  - $(Get-Prop $t 'DisplayName')" }
    if (-not (Confirm-Yes 'These will start sending notifications immediately.')) { return }

    foreach ($t in $go) {
        try {
            Enable-SCOMNotificationSubscription -Subscription $t -ErrorAction Stop
            Write-Log "Enabled: $(Get-Prop $t 'DisplayName')" 'OK'
        }
        catch { Write-Log "FAILED to enable $(Get-Prop $t 'DisplayName'): $(Get-ExceptionChain $_)" 'ERROR' }
    }
}

function Invoke-Disable {
    Write-Banner 'DISABLE SUBSCRIPTIONS (TARGET)'
    if (-not $Name -and -not $All) { throw 'Give -Name <display name> (wildcards OK) or -All.' }
    $null = Connect-Scom
    $targets = if ($All) { @(Get-SCOMNotificationSubscription -ErrorAction Stop) } else { @(Select-Subscriptions $Name) }
    $targets = @($targets | Where-Object { Get-Prop $_ 'Enabled' })
    if ($targets.Count -eq 0) { Write-Log 'Nothing enabled to disable.'; return }
    foreach ($t in $targets) {
        try {
            Disable-SCOMNotificationSubscription -Subscription $t -ErrorAction Stop
            Write-Log "Disabled: $(Get-Prop $t 'DisplayName')" 'OK'
        }
        catch { Write-Log "FAILED to disable $(Get-Prop $t 'DisplayName'): $(Get-ExceptionChain $_)" 'ERROR' }
    }
}

###########################################################################
# Main
###########################################################################

try {
    switch ($Step) {
        'Export'  { if ($SourceServer -and -not $RemoteChild) { Invoke-RemoteExport } else { Invoke-Export } }
        'Prepare' { Invoke-Prepare }
        'Import'  { Invoke-Import }
        'Verify'  { Invoke-Verify }
        'Enable'  { Invoke-Enable }
        'Disable' { Invoke-Disable }
    }
}
catch {
    Write-Log "STOPPED: $(Get-ExceptionChain $_)" 'ERROR'
    Add-Content -LiteralPath $NotifLogFile -Value "Stack: $($_.ScriptStackTrace)" -Encoding UTF8
    Write-Log "Log: $NotifLogFile"
    if ($RemoteChild) { throw }
    exit 1
}
Write-Log "Log: $NotifLogFile"
