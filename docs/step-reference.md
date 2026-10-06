# Step reference

- [The working folder](#the-working-folder)
- [The manifest](#the-manifest)
- [Export-ScomEnvironment.ps1](#export-scomenvironmentps1)
- [Exporting the source from the target server](#exporting-the-source-from-the-target-server)
- [Invoke-ScomMigrationStep.ps1 steps](#invoke-scommigrationstepps1)
- [Import-Batch.\<ver\>.ps1](#import-batchverps1)
- [MPMigration.ps1 run directly](#mpmigrationps1-run-directly)
- [Output files](#output-files)

## The working folder

`Invoke-ScomMigrationStep.ps1` works out every path from where it is saved:

```
<working folder>\
    Invoke-ScomMigrationStep.ps1
    MPMigration.ps1
    MigrationManifest.csv
    Migration.settings.psd1        (optional)
    Source\                        whole output folder of Export-ScomEnvironment.ps1 -Role Source
    Target\AllMPs\                 AllMPs folder of Export-ScomEnvironment.ps1 -Role Target
    Compile\                       created by Compile
    InstanceMap.csv                created by MapInstances
```

### Settings

| Parameter | Default | Meaning |
|---|---|---|
| `-ManagementServer` | `$env:COMPUTERNAME` | Target management server, for steps that connect |
| `-SourceFolder` | `Source` | Source export folder, relative to the script or absolute |
| `-TargetFolder` | `Target` | Folder that contains `AllMPs\` |
| `-SourceVersion` | `2016` | Label used in messages and reports |
| `-TargetVersion` | `2025` | Label; also names `Import-Batch.<TargetVersion>.ps1` |
| `-SourceServer` | — | `ExportSource`: the source management server to export through PowerShell remoting |
| `-SealedSearchPath` | — | `ExportSource`: folders or shares with original `.mp`/`.mpb` files, searched from this server |
| `-SourceCredential` | your account | `ExportSource`: credential for the remoting session (command line only) |

You can put the same keys in `Migration.settings.psd1` next to the script so nobody has to type them (see [`examples/`](../examples/Migration.settings.psd1)). A parameter given on the command line wins over the settings file.

The version labels **don't change any logic** in v3.x. The knowledge of which Microsoft packs replaced which is still written into the script. See the [roadmap](roadmap-v4.md).

## The manifest

`MigrationManifest.csv` decides what is in scope. Only rows with `Migrate = Y` are moved.

| Column | Required | Meaning |
|---|---|---|
| `Migrate` | yes | `Y` / `Yes` / `True` / `1` = move it. Anything else = leave it. |
| `ManagementPack` | yes | A name for the row. Shown in reports. |
| `MatchPattern` | no | Matched against the MP **ID**, then its **display name**. Wildcards `*` and `?` are allowed. One row can match several MPs. Defaults to `ManagementPack`. |
| `Action` | no | `MIGRATE` (default), `VENDOR_SEALED`, `OVERRIDES` or `SKIP`. `OVERRIDES` rows are the ones `EnableOverrides` switches on and `OverrideReport` summarises. |
| `Tier`, `Reason`, … | no | Your own notes. They're carried through but not used. |

Check `ManifestResolution.csv` (source export) or `ManifestMatch.csv` (compile) for rows marked `NO MATCH`.

## Export-ScomEnvironment.ps1

This script is read-only against SCOM. It runs with each environment's own `OperationsManager` module: either directly on that environment's management server, or, for the source, from the target server with `-SourceServer` (see [below](#exporting-the-source-from-the-target-server)).

| Parameter | Meaning |
|---|---|
| `-Role Source\|Target` | Required |
| `-OutputFolder` | Default `C:\SCOMMigration\<Role>-<timestamp>` |
| `-ManagementServer` | Default `localhost` |
| `-Credential` | Optional. With `-SourceServer`, the credential for the remoting session. |
| `-SourceServer` | Source role. Run the export on this server through PowerShell remoting and copy the result into `-OutputFolder` here. |
| `-Manifest` | Source role. Limits sealed-original searching and group/override analysis to in-scope MPs, and writes `ManifestResolution.csv`. |
| `-SealedSearchPath` | Source role. Extra folders and shares to search for original `.mp`/`.mpb` files. |
| `-GroupPatternMode` | Source role. `Tight` (default), `Wildcard` or `Exact`. See [How it works](how-it-works.md#static-groups--dynamic-groups). |
| `-SkipGroupConversion` | Source role. Don't analyse static groups. |

What each role writes:

- **Source:**
  - `AllMPs\`
  - `SourceInventory.csv`
  - `ManifestResolution.csv`
  - `SealedOriginals\`
  - `SealedOriginalsMissing.csv`
  - `SealedOriginalsFound.csv` (which file was used for each sealed MP, and where it was found)
  - `SealedDependencies\` (vendor sealed MPs the in-scope MPs use)
  - `SealedDependencies_Microsoft_optional\`
  - `GroupConversion.csv`
  - `StaticGroupMembers.csv`
  - `OverrideInstances.csv`
  - `Export-Source.log` (with `-SourceServer`, also `Export-Source.remote-run.log` from the target side)
- **Target:**
  - `AllMPs\`
  - `TargetInventory.csv`
  - `Export-Target.log`

## Exporting the source from the target server

A newer `OperationsManager` module can't connect to an older management group. With `-SourceServer` (or the `ExportSource` step) the export still runs **on** the source management server with the source's own module, but you start it from the target server:

```powershell
.\Invoke-ScomMigrationStep.ps1 ExportSource -SourceServer OLDSCOM01 -SealedSearchPath '\\fileserver\MPs'
.\Invoke-ScomMigrationStep.ps1 ExportTarget
```

What happens:

1. A PowerShell remoting session is opened to the source server.
2. The export script's text is sent over the session and run there as a script block. Nothing is copied to the source except the manifest, and the source's execution policy doesn't apply.
3. It writes to a temporary folder under the source's `%TEMP%`, which is copied back to `Source\` here and then deleted.
4. The source's own install folders are searched there for sealed originals. `-SealedSearchPath` is searched **from the target server** afterwards. A remoting session can't pass your credentials on to a third server (the "double hop"), so shares are read from here.
5. An existing `Source\` is kept as `Source.previous-<timestamp>`, and originals you dropped into its `SealedOriginals\` are found again.

Requirements:

- PowerShell remoting (WinRM, TCP 5985) from the target to the source server. Check with `Test-WSMan <server>`.
- Your account (or `-SourceCredential`) is an administrator on the source server and can read the source management group.
- Space under `%TEMP%` on the source server for one export of every MP.

Quick test from the target server:

```powershell
Invoke-Command -ComputerName OLDSCOM01 {
    Import-Module OperationsManager
    New-SCOMManagementGroupConnection -ComputerName localhost
    (Get-SCOMManagementGroupConnection).ManagementGroupName
}
```

If remoting isn't allowed, run `Export-ScomEnvironment.ps1 -Role Source` on the source server and copy the folder across, as before.

## Invoke-ScomMigrationStep.ps1

Each step is one word, and the steps are listed in their usual order. Steps that touch SCOM connect to `-ManagementServer`, which is the **target**. The only exception is `ExportSource`, which reads the source through PowerShell remoting.

| Step | Touches SCOM | What it does | Writes |
|---|---|---|---|
| `ExportSource` | read (source) | Runs `Export-ScomEnvironment.ps1 -Role Source` on `-SourceServer` through PowerShell remoting with `MigrationManifest.csv`, and copies the result here. Keeps any previous export as `Source.previous-<timestamp>`. | `Source\` |
| `ExportTarget` | read | Runs `Export-ScomEnvironment.ps1 -Role Target` against this management group. Keeps any previous export as `Target.previous-<timestamp>`. | `Target\` |
| `Check` | no | Pre-flight. Counts exported MPs, sealed originals and dependencies. Lists manifest rows with no match, in-scope sealed MPs with no original file, and REVIEW group rules. It flags risky REVIEW groups (names containing Disable, NonProd, Test, Maint…), where over-matching would actually hurt. | `Check-Report.txt` |
| `FixReview` | no | Rewrites REVIEW group rules to exact current member names, and rules whose members are Health Service Watchers to a display-name match. It keeps the original as `GroupConversion.original.csv`. Safe to run again: if a fresh export replaces `GroupConversion.csv`, the new file is used. | `Source\GroupConversion.csv` |
| `EnableOverrides` | no | Optional. Sets `Migrate = Y` on every `Action = OVERRIDES` row. | `MigrationManifest.csv` (backup: `MigrationManifest.before-overrides.csv`) |
| `MapInstances` | read | For each object an override targets by ID, finds the same object in the target (by FullName, otherwise by an unambiguous display name). | `InstanceMap.csv` |
| `Compile` | no | Runs `MPMigration.ps1` with every input it finds. | `Compile\` (see below) |
| `WhyBlocked` | no | One screen: BLOCKED MPs grouped by cause, plus the groups that were dropped so their MP could import. | `WhyBlocked-*.txt` |
| `OverrideReport` | no | Override packs: READY/BLOCKED, overrides kept, re-pointed and removed by reason, and the missing MPs that account for most removals. | — |
| `DryRun` | read | `Import-Batch.<ver>.ps1 -WhatIfOnly` | `Compile\ImportResults_*.csv` |
| `Import` | **write** | Imports READY rows in order. Asks you to type `YES` after showing the management group name. | `Compile\ImportResults_*.csv`, `ImportErrors_*.txt` |
| `Reimport` | **write** | Same as `Import`, but also overwrites **unsealed** MPs already installed at the same version. Use it after a re-compile, for example when more overrides have been mapped. | same |
| `TestGroups` | read | For every converted group: members now, target agents the pattern matches, and source member names. Status is WORKING / WAITING / NOT CALCULATED YET / BROKEN. | `GroupTest-*.csv` |
| `Rollback` | **write** | Removes every MP the import runs recorded as `Imported`, newest first, in up to three passes. Asks you to type `YES`. | `Rollback-*.csv` |
| `Collect` | no | Zips the result files for review. **They contain your server and MP names.** | `MigrationResults-*.zip` |

A typical first run is:

```
ExportSource → ExportTarget → Check → FixReview → (EnableOverrides) → (MapInstances) → Compile → WhyBlocked → DryRun → Import → TestGroups
```

A re-export replaces `Source\GroupConversion.csv`. Your edits stay in `Source.previous-*\`: run `FixReview` again or copy them across, and run `MapInstances` again.

After you fix something (add a sealed original, install a Microsoft pack and re-export the target, edit `GroupConversion.csv`), run `Compile` again, then `DryRun`, then `Import` or `Reimport`.

## Import-Batch.\<ver\>.ps1

`Compile` writes this script into `Compile\`. It reads `BatchImportManifest.csv` next to it.

| Parameter | Meaning |
|---|---|
| `-ManagementServer` | Target MS (default `localhost`) |
| `-WhatIfOnly` | Dry run |
| `-IncludeBlocked` | Also attempts BLOCKED rows, to capture SCOM's own error text |
| `-Reimport` | Overwrites unsealed MPs already present at the same version |
| `-Only <pattern>` | Only MP IDs matching the wildcard, e.g. `-Only "Contoso.App*"` |
| `-ExpectedManagementGroup <name>` | Refuses to run against any other management group. You aren't asked to type YES. |
| `-Force` | Skips the YES prompt |

Per MP, the result is `Imported`, `WouldImport`, `AlreadyPresent` (same or newer version installed), `SkippedBlocked`, `SkippedDependencyFailed` or `Failed`. For every failure, SCOM's full exception chain is written to `ImportErrors_*.txt`.

## MPMigration.ps1 run directly

The step runner covers the normal path. You can also run the compiler yourself:

```powershell
.\MPMigration.ps1 -NonInteractive `
    -InputPath D:\SCOMMigration\Source\AllMPs, D:\SCOMMigration\Source\SealedOriginals `
    -SourceInventory D:\SCOMMigration\Source\SourceInventory.csv `
    -SourceRepositoryFolder D:\SCOMMigration\Source\SealedDependencies `
    -GroupConversionFile D:\SCOMMigration\Source\GroupConversion.csv `
    -InstanceMapFile D:\SCOMMigration\InstanceMap.csv `
    -RepositoryFolder D:\SCOMMigration\Target\AllMPs `
    -Manifest D:\SCOMMigration\MigrationManifest.csv `
    -OutputFolder D:\SCOMMigration\Compile `
    -SourceVersion 2016 -TargetVersion 2025
```

If you leave out `-NonInteractive` and the paths, it opens file and folder pickers. `Get-Help .\MPMigration.ps1 -Full` documents every parameter. The ones you're most likely to need:

| Parameter | Use |
|---|---|
| `-KeepBlockedGroups` | Block the whole MP instead of dropping a group whose member class is missing |
| `-ExcludeMPIDs` | Treat these missing dependencies as known gaps, so they aren't searched for again |
| `-AuditOnly` | Score a whole source environment against the target (batch-internal references don't count as resolved) |
| `-Strict` | Treat ERROR diagnostics as fatal |
| `-AllowIDRewrite` / `-ForceRewriteCoreLibraries` | Riskier rewrites. Read the help first. |
| `-CheckCatalog` / `-CatalogOnly` | Compare the target against Microsoft's MP catalog on Microsoft Learn (needs internet access) |
| `-LiveImport` | Import auto-resolved dependencies during compile. Off by default, and the step runner never uses it. |

## Output files

| File | From | Contents |
|---|---|---|
| `BatchImportManifest.csv` | Compile | One row per in-scope MP, in import order. Columns include `Verdict`, `BlockingReasons`, `BatchDependsOn`, `ImportPath`, `OverridesKept`, `InstanceRemapped`, `InstanceUnmapped` |
| `CandidateMPs\*.xml` | Compile | What will be imported (unsealed). `_analysis_only_do_not_import\` holds XML of sealed MPs, kept for reference only. |
| `CandidateMPs\*.Stripped.csv`, `*.SemanticRecommendations.csv` | Compile | Per-MP removals; likely successor packs for missing references (a hint only) |
| `StrippedElements.csv` | Compile | Every element removed, with the reason |
| `MissingDependencies.csv` | Compile | Every MP referenced but not found, who needs it, and where to get it |
| `GroupConversionResults.csv` | Compile | Every static membership rule: converted / left / dropped |
| `StaticGroupMembership.csv` | Compile | Groups that still have static members after conversion |
| `ManifestMatch.csv` | Compile | Manifest row → MP ID(s) |
| `MPCompiler.*.log` | Compile | Full log; the BOTTOM LINE summary is at the end |
| `ImportResults_*.csv`, `ImportErrors_*.txt` | Import | Per-MP result; SCOM exception chains |
| `InstanceMap.csv` | MapInstances | Source object GUID → target object GUID |
| `GroupTest-*.csv` | TestGroups | Per converted group rule: status, members, matching agents |
