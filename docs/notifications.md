# Notifications: channels, subscribers and subscriptions

`src/Migrate-ScomNotifications.ps1` moves a management group's notification setup to a **new, empty** management group. That covers every channel, subscriber and subscription. It's separate from the MP compiler and can run before or after it.

It was used on a production SCOM 2016 → 2025 migration of about 135 subscriptions.

## Approach

All notification config lives in one unsealed MP, `Microsoft.SystemCenter.Notifications.Internal`. The script moves that MP instead of recreating objects one by one through the `Add-SCOMNotification*` cmdlets. The cmdlets can't express much of what a subscription holds, so everything below is carried over unchanged:

- criteria (alert-property expressions)
- class and group scope
- To, CC and BCC recipients
- subscriber schedules
- HTML bodies and encodings
- SMTP port, authentication and secondary servers
- command channels

Every subscription arrives **disabled**. You switch them on deliberately, one at a time or all together, after you have checked them.

## Requirements

- **The target has no notification config yet.** Prepare blocks if the target already has any subscription, subscriber or channel, because importing would replace them.
- **Windows PowerShell 5.1 on each management server,** with that environment's own `OperationsManager` module. Export runs on the source and every other step on the target. A newer console can't connect to an older management group.
- **SCOM Administrator on the target.** Read access is enough on the source.
- **The MPs your subscriptions are scoped to already exist on the target.** Those are your custom groups and classes, so run the MP migration first if subscriptions point at custom groups.

## Run order

The default work folder is `C:\SCOMMigration\Notifications`. Change it with `-WorkFolder`, and use the same path on both servers.

| # | Where | Command | Changes SCOM? |
|---|---|---|---|
| 1 | Source MS | `.\Migrate-ScomNotifications.ps1 -Step Export` | No |
| – | | Copy `<WorkFolder>\Export` to the same path on the target MS | |
| 2 | Target MS | `.\Migrate-ScomNotifications.ps1 -Step Prepare` | No |
| 3 | Target MS | `.\Migrate-ScomNotifications.ps1 -Step Import` (type `YES`) | **Yes** |
| 4 | Target MS | Work through the follow-ups in `VerifyReport.*.csv` | |
| 5 | Target MS | `.\Migrate-ScomNotifications.ps1 -Step Enable -Name 'Ops - Critical'` (type `YES`) | **Yes** |
| 6 | Target MS | `.\Migrate-ScomNotifications.ps1 -Step Enable -AllEnabledOnSource` (type `YES`) | **Yes** |

**Emergency stop:** `.\Migrate-ScomNotifications.ps1 -Step Disable -All` turns every subscription off. `-Step Verify` can be run at any time.

The work folder ends up looking like this:

```
<WorkFolder>\
├── Export\     MP + SourceSubscriptions/Channels/Subscribers.csv + SourceGuidMap.csv + SourceSummary.json
├── Prepare\    prepared MP + PrepareSummary.json + PreparedSubscriptions.csv + ScopeCheck.csv
│               + ReferenceCheck.csv + NotInThisMp.csv (if any) + TargetBackup.*\
├── Import\     VerifyReport.*.csv + ImportErrors.*.txt (if the import failed)
└── Logs\       one log per step run
```

## What each step does

### Export (source, read only)

- Exports `Microsoft.SystemCenter.Notifications.Internal`.
- Writes an inventory through the cmdlets: `SourceSubscriptions.csv`, `SourceChannels.csv` and `SourceSubscribers.csv`.
- Resolves every GUID that appears in a subscription. That covers classes, groups, monitors, rules and individual objects. They're written to `SourceGuidMap.csv` while the source can still say what each one is.

### Prepare (target, read only)

Builds `Prepare\Microsoft.SystemCenter.Notifications.Internal.xml`:

- **Every subscription rule is set to `Enabled="false"`.** The original state is kept in `PreparedSubscriptions.csv` (`EnabledOnSource`).
- **The MP version is set above the target's own copy.** That copy is backed up first to `Prepare\TargetBackup.*`.
- **References are set to the versions installed on the target.** An unused reference the target doesn't have is removed. A *used* one blocks the import.
- **Per-server scope is re-pointed.** Each object is matched on the target by FullName and class. Class, group, monitor and rule IDs are derived from MP and element names, so they're the same on both sides once those MPs exist on the target.
- **Anything that can't be found or matched is listed in `ScopeCheck.csv`.** The affected subscriptions are marked `EnableAllowed = False`.
- **The verdict is `READY` or `BLOCKED`**, with the reasons in `PrepareSummary.json`. It also writes a SHA-256 of the prepared file, which Import checks.

Prepare blocks when:

- the target already has notification config;
- a used reference is missing on the target;
- the source inventory names a subscription that appears in the MP but wasn't recognised as a subscription rule, so it can't be guaranteed disabled;
- `SourceSubscriptions.csv` has far fewer rows than Export recorded, which suggests the file was damaged.

A subscription that the source lists but whose internal name doesn't appear in the MP at all is stored in another MP. That's a **warning**, not a blocker. It's listed in `NotInThisMp.csv`; see [Troubleshooting](#troubleshooting).

### Import (target)

- Refuses to run unless Prepare said `READY` and the prepared file is byte-for-byte unchanged.
- Re-checks that the target still has no subscriptions and that its copy of the MP is still the one Prepare saw.
- Asks you to type `YES`, then imports.
- If any subscription arrives enabled, it disables it on the spot.
- Runs Verify.

### Verify (target, read only)

Compares the target with the source inventory, subscription by subscription, channel by channel and subscriber by subscriber. Device counts are compared too. It flags the follow-ups below and writes `Import\VerifyReport.*.csv`.

### Enable / Disable (target)

- `-Name` takes display names, with wildcards.
- `-AllEnabledOnSource` enables only the subscriptions that were enabled on the source.
- Enable skips any subscription with scope issues unless you pass `-Force`. Fix its scope in the console first.
- Enable asks for `YES`. Disable doesn't ask, because it's the emergency stop.

## Follow-ups before enabling

`VerifyReport.*.csv` lists these:

- **SMTP channels with Windows Integrated authentication.** Associate a Run As account with the Notification Action Account profile on the target, or mail fails. Run As accounts are never migrated.
- **Command channels.** The program must exist on every management server in the Notifications resource pool.
- **Subscriptions with scope issues.** These scope a group, class or server that doesn't exist on the target. Re-point the scope in the console, then enable.

## Troubleshooting

**"The device is not ready" or "Can't write to work folder".** The work folder is on a drive that doesn't exist or can't be written to on that server. Pass `-WorkFolder` with a path on a drive that does, and use the same path on both servers.

**A subscription blocks with "referenced in this MP but weren't recognised as subscription rules".** The source inventory lists a subscription that the script can't find as a rule in the MP. If that subscription is unused (for example, it has no subscribers), you can drop it from the check. On the target:

```powershell
$csv = '<WorkFolder>\Export\SourceSubscriptions.csv'
Copy-Item $csv "$csv.bak"
(Import-Csv $csv) | ? DisplayName -notlike '<display name>*' | Export-Csv $csv -NoTypeInformation -Encoding UTF8
```

Keep the parentheses around `(Import-Csv $csv)`. Without them, Windows PowerShell 5.1 starts writing the file before it has finished reading it and empties it. Prepare now detects a damaged inventory and stops, but it's better not to need that.

If you can, please open an issue with the output of the following command (scrub names first). It shows how the subscription is stored, so the script can learn to recognise it:

```powershell
$n = (Import-Csv '<WorkFolder>\Export\SourceSubscriptions.csv.bak' | ? DisplayName -like '<display name>*').Name
Select-String -Path '<WorkFolder>\Export\Microsoft.SystemCenter.Notifications.Internal.xml' -Pattern $n -SimpleMatch
```

**A subscription is in `NotInThisMp.csv`.** It's stored in another MP, often a custom MP. To find which one, run this on the source:

```powershell
$s = Get-SCOMNotificationSubscription | ? DisplayName -like '<display name>*'
(Get-SCOMRule -Name $s.Name).GetManagementPack() | Select Name, DisplayName, Sealed
```

It comes over with that MP if the MP is migrated. Otherwise recreate it in the console.

**Import fails.** Nothing is changed. SCOM's full exception chain is in `Import\ImportErrors.*.txt`.

## Known limits

- **Empty target only.** Merging into a target that already has notification config isn't supported.
- **Where the config is assumed to live.** Channels and subscribers are assumed to be in the same MP as the subscriptions. That's where the console puts them. If Verify reports channels or subscribers as MISSING after the import, stop and open an issue.
- **Subscriptions stored in other MPs aren't moved by this script** (see above).
- **What isn't migrated:** Run As accounts and their profile associations, and command-channel programs on disk.
- **Offline testing:** `tests/Invoke-NotificationsTest.ps1` covers the steps and every blocker against a synthetic MP and stand-in cmdlets. Real-world validation so far is a single SCOM 2016 → 2025 migration.

## Before you share output

The CSVs and logs contain subscriber names, email addresses, phone numbers, SMTP servers and server names. Scrub them before you attach anything to a public issue.
