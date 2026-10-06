# SCOM-Migration

A PowerShell toolkit to assist in migrating management packs from one SCOM management group to another.

Move custom management packs from an old System Center Operations Manager (SCOM) management group to a new one. That means your groups, monitors, rules and overrides. The tool fixes what it can automatically and tells you exactly what it couldn't.

It was built for and used on a production **SCOM 2016 → SCOM 2025** migration of several hundred MPs. That is the only path tested end to end so far. Other source and target versions are on the [v4.0 roadmap](docs/roadmap-v4.md).

```
 OLD management group                      NEW management group  (you work here)
 ─────────────────────                     ─────────────────────
 export runs here with its own   ◄──────── Invoke-ScomMigrationStep.ps1 ExportSource
 OperationsManager module ──── remoting ──► Source\
                                           Invoke-ScomMigrationStep.ps1 ExportTarget
                                                       │
                 Invoke-ScomMigrationStep.ps1 Compile   (MPMigration.ps1 does the work)
                                 │
              READY / BLOCKED per MP  +  Import-Batch.<ver>.ps1
                                 │
                 Invoke-ScomMigrationStep.ps1 DryRun → Import → TestGroups
```

## What it does

- **Compiles every unsealed MP against the target's real MP set before anything is imported.**
  - Reference versions are raised to what the target actually has.
  - Overrides, categories, folder items, display strings and knowledge articles that point at something the target no longer has are removed.
  - Each removal is listed in `StrippedElements.csv`.
- **Gives every MP a verdict: `READY` or `BLOCKED`, with the exact reason.** A BLOCKED MP is one where a class, monitor, rule or view still depends on something missing from the target. Usually that is a Microsoft pack that was retired, such as SQL 2012 or IIS 2008.
- **Imports in dependency order.** One failure doesn't stop the run, and anything that depends on a failed MP is skipped instead of attempted.
- **Turns static groups into dynamic ones.** Explicit member GUIDs are meaningless in a new management group. They become NetBIOS-name rules (`^(APPSQL01|APPSQL02)$`, or a pattern) that populate as soon as the agents report in.
  - Health Service Watcher members are matched on display name.
  - Hosted classes are matched through their Windows Computer.
  - Nested groups are kept.
- **Re-points per-server overrides** ("disable this monitor on SERVER01"). The source object ID is resolved to its full name and the same object is found in the target. Overrides for servers that aren't in the target are removed and reported.
- **Treats sealed MPs as sealed.** They are imported from the original `.mp`/`.mpb` and never edited. An exported `.xml` of a sealed MP is never used as a stand-in.
- **Leaves a CSV trail for every step**, so you can review and hand off the work.

## What it doesn't do

- It doesn't rewrite MPs to use Microsoft's *replacement* packs. A custom monitor that targets a SQL 2012 class is reported as BLOCKED, not re-targeted. Element IDs changed between those packs, so a blind rewrite would produce MPs that import but monitor nothing.
- It doesn't migrate Run As accounts, maintenance schedules, dashboards outside MPs, or agent assignments. Notification channels, subscribers and subscriptions have their own script; see [Notifications](#notifications).
- It never writes to the source management group. The source side only calls `Get-*` and `Export-SCOMManagementPack`. With `ExportSource`, the only thing created on the source server is a temporary folder, which is removed afterwards.
- It doesn't import anything until you run the import step, and the import asks you to type `YES`.

## Requirements

| | |
|---|---|
| PowerShell | Windows PowerShell 5.1 on the management servers (PowerShell 7 works for the offline tests) |
| SCOM module | The `OperationsManager` module **of each environment**. A newer console can't connect to an older management group, so the source export runs on the old MS, either started from the new MS through PowerShell remoting (`ExportSource`) or run there by hand. |
| Remoting | For `ExportSource`: PowerShell remoting (WinRM) from the new MS to the old MS, as an administrator there. Check with `Test-WSMan <old MS>`. Without it, export on the old MS and copy the folder. |
| Rights | SCOM Administrator on the target for import. Read access on the source. |
| Sealed MPs | The original `.mp`/`.mpb` files for every in-scope sealed MP: vendor packs and your own sealed libraries |
| `.mp`/`.mpb` reading | The SCOM SDK assemblies, present on any management server |

## Quick start

Everything below runs on the **new** management server.

1. **Download** `src/` into a working folder, for example `C:\SCOMMigration`, and unblock the files:

   ```powershell
   Get-ChildItem *.ps1 | Unblock-File
   ```

2. **Write your manifest.** Start from [`examples/MigrationManifest.example.csv`](examples/MigrationManifest.example.csv) and save it as `MigrationManifest.csv` in the working folder. One row per MP, and only rows with `Migrate = Y` are moved. See [Manifest format](docs/step-reference.md#the-manifest).

3. **Export both management groups.** Install the current Microsoft packs you rely on in the new one first (Windows Server, SQL, IIS…).

   ```powershell
   .\Invoke-ScomMigrationStep.ps1 ExportSource -SourceServer OLDSCOM01 -SealedSearchPath '\\fileserver\MPs'
   .\Invoke-ScomMigrationStep.ps1 ExportTarget
   ```

   `ExportSource` runs the export **on** the old management server through PowerShell remoting, with its own SCOM module, and copies the result into `Source\`. `ExportTarget` fills `Target\`. To avoid typing the server and share every time, copy [`examples/Migration.settings.psd1`](examples/Migration.settings.psd1) next to the script.

   No remoting to the old server? Run `.\Export-ScomEnvironment.ps1 -Role Source -Manifest .\MigrationManifest.csv -OutputFolder C:\SCOMMigration\Source` there and copy the folder across. See [Exporting the source from the target server](docs/step-reference.md#exporting-the-source-from-the-target-server).

4. **Run the steps** from the working folder:

   ```powershell
   .\Invoke-ScomMigrationStep.ps1 Check          # pre-flight; changes nothing
   .\Invoke-ScomMigrationStep.ps1 FixReview      # tighten group patterns that would over-match
   .\Invoke-ScomMigrationStep.ps1 MapInstances   # only if you carry overrides
   .\Invoke-ScomMigrationStep.ps1 Compile        # READY/BLOCKED per MP; imports nothing
   .\Invoke-ScomMigrationStep.ps1 WhyBlocked
   .\Invoke-ScomMigrationStep.ps1 DryRun
   .\Invoke-ScomMigrationStep.ps1 Import         # asks you to type YES
   .\Invoke-ScomMigrationStep.ps1 TestGroups     # are the converted groups populating?
   ```

Every step is one word because the toolkit was built for a locked-down server where every command had to be typed by hand.

## Notifications

`src/Migrate-ScomNotifications.ps1` moves notification channels, subscribers and subscriptions to a target that has none yet. It moves the `Notifications.Internal` MP as a whole, so criteria, scope, CC/BCC, schedules and SMTP settings come over intact. Every subscription arrives **disabled**, and you enable them yourself when you're ready.

```powershell
.\Migrate-ScomNotifications.ps1 -Step Export -SourceServer OLDSCOM01   # from the target MS, through remoting
.\Migrate-ScomNotifications.ps1 -Step Prepare     # on the target MS; READY/BLOCKED, imports nothing
.\Migrate-ScomNotifications.ps1 -Step Import      # asks you to type YES; everything arrives disabled
.\Migrate-ScomNotifications.ps1 -Step Enable -Name 'Ops - Critical'
.\Migrate-ScomNotifications.ps1 -Step Disable -All   # emergency stop
```

See [docs/notifications.md](docs/notifications.md).

## Documentation

- [How it works](docs/how-it-works.md): the compile model, stripping, sealed vs. unsealed, READY/BLOCKED, group conversion, per-server overrides
- [Step reference](docs/step-reference.md): every step, every output file, the manifest and the direct `MPMigration.ps1` parameters
- [Troubleshooting](docs/troubleshooting.md)
- [Known limits](docs/known-limits.md)
- [Notifications](docs/notifications.md): channels, subscribers and subscriptions
- [Roadmap: v4.0, any supported version → any supported version](docs/roadmap-v4.md)

## Repository layout

```
src/        MPMigration.ps1, Export-ScomEnvironment.ps1, Invoke-ScomMigrationStep.ps1,
            Migrate-ScomNotifications.ps1
examples/   MigrationManifest.example.csv, Migration.settings.psd1
docs/       how-it-works, step-reference, troubleshooting, known-limits, notifications, roadmap-v4
tests/      offline tests: synthetic MPs + stand-in OperationsManager cmdlets
```

## Testing without SCOM

```powershell
pwsh ./tests/Invoke-SmokeTest.ps1           # MP migration
pwsh ./tests/Invoke-NotificationsTest.ps1   # notifications
```

This runs the full pipeline: source export, Check, FixReview, EnableOverrides, MapInstances, Compile, WhyBlocked, OverrideReport, DryRun and TestGroups. It uses synthetic MPs and mocked SCOM cmdlets, then checks the verdicts, stripping, group conversion and override re-pointing. It runs on Windows, Linux or macOS. See [CONTRIBUTING](CONTRIBUTING.md).

## Before you share output

`Collect` zips the result CSVs for review. Those files contain your **server names, group names and MP names**. Scrub them before you attach them to a public issue.

## License

[MIT](LICENSE). Use it at your own risk. Run `DryRun` first, and keep a backup of the target's unsealed MPs (`Export-ScomEnvironment.ps1 -Role Target` does exactly that).
