# SCOM-Migration

A PowerShell toolkit to assist in migrating management packs from one SCOM management group to another.

Move custom management packs from an old System Center Operations Manager (SCOM) management group to a new one. That means your groups, monitors, rules and overrides. The tool fixes what it can automatically and tells you exactly what it couldn't.

It was built for and used on a production **SCOM 2016 → SCOM 2025** migration of several hundred MPs. That is the only path tested end to end so far. Other source and target versions are on the [v4.0 roadmap](docs/roadmap-v4.md).

```
 OLD management group                      NEW management group
 ─────────────────────                     ─────────────────────
 Export-ScomEnvironment.ps1 -Role Source   Export-ScomEnvironment.ps1 -Role Target
            │                                          │
            └──────────── copy both folders ───────────┘
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
- It never writes to the source management group. The source side only calls `Get-*` and `Export-SCOMManagementPack`.
- It doesn't import anything until you run the import step, and the import asks you to type `YES`.

## Requirements

| | |
|---|---|
| PowerShell | Windows PowerShell 5.1 on the management servers (PowerShell 7 works for the offline tests) |
| SCOM module | The `OperationsManager` module **of each environment**. Run the source export on the old MS and everything else on the new MS. A newer console can't connect to an older management group. |
| Rights | SCOM Administrator on the target for import. Read access on the source. |
| Sealed MPs | The original `.mp`/`.mpb` files for every in-scope sealed MP: vendor packs and your own sealed libraries |
| `.mp`/`.mpb` reading | The SCOM SDK assemblies, present on any management server |

## Quick start

1. **Download** `src/` and copy it to a working folder on each management server. Then unblock the files:

   ```powershell
   Get-ChildItem *.ps1 | Unblock-File
   ```

2. **Write your manifest.** Start from [`examples/MigrationManifest.example.csv`](examples/MigrationManifest.example.csv). One row per MP, and only rows with `Migrate = Y` are moved. See [Manifest format](docs/step-reference.md#the-manifest).

3. **Export the source**, on the old management server:

   ```powershell
   .\Export-ScomEnvironment.ps1 -Role Source -Manifest .\MigrationManifest.csv `
       -OutputFolder D:\SCOMMigration\Source -SealedSearchPath '\\fileserver\MPs'
   ```

4. **Export the target**, on the new management server. Install the current Microsoft packs you rely on first (Windows Server, SQL, IIS…).

   ```powershell
   .\Export-ScomEnvironment.ps1 -Role Target -OutputFolder D:\SCOMMigration\Target
   ```

5. **Lay out the working folder on the new MS:**

   ```
   D:\SCOMMigration\
       Invoke-ScomMigrationStep.ps1  MPMigration.ps1  MigrationManifest.csv
       Source\          <- the whole source export folder
       Target\AllMPs\   <- from the target export
   ```

   For other folder names or version labels, copy [`examples/Migration.settings.psd1`](examples/Migration.settings.psd1) next to the script.

6. **Run the steps** from `D:\SCOMMigration`:

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
.\Migrate-ScomNotifications.ps1 -Step Export      # on the source MS
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
