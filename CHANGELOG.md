# Changelog

## 3.40.0 (first public release)

This is the code used for a production SCOM 2016 → 2025 migration, prepared for publishing. The compile and import logic is identical to the production build. A regression run on the same inputs produced identical verdicts, candidate MPs and group conversions.

### Publishing changes
- Renamed the step runner to `Invoke-ScomMigrationStep.ps1`.
- The runner takes `-SourceFolder`, `-TargetFolder`, `-SourceVersion` and `-TargetVersion`, and can read them from an optional `Migration.settings.psd1`. The defaults are `Source`, `Target`, `2016` and `2025`.
- `Compile` passes the version labels through to `MPMigration.ps1`, and the import script is named `Import-Batch.<TargetVersion>.ps1`.
- Renamed CSV columns:
  - `ExtraIn2016` → `ExtraInSource` and `MatchesIn2016` → `MatchesInSource` (in `GroupConversion.csv`). The runner still accepts files that use the old names.
  - `MembersIn2025` → `MembersInTarget`, `MatchingAgentsIn2025` → `MatchingAgentsInTarget`, `MemberNames2016` → `MemberNamesSource` and `MatchingNames2025` → `MatchingNamesTarget` (in `GroupTest-*.csv`).
- Messages, prompts and help no longer assume 2016/2025 and no longer name any organisation.
- Added the offline smoke test (`tests/`), the example manifest and the docs.

### Compiler (MPMigration.ps1)
- Override validation reads `Monitoring/Overrides`. Earlier builds read a path that doesn't exist, so no override was ever validated.
- A reference whose exact ID exists in the target is no longer version-rewritten by the family pattern table.
- Dead-element stripping:
  - Every override, category and folder item is checked against the actual element list of the MP it points at.
  - Display strings and knowledge articles for removed elements are removed with them.
  - A `<Reference>` is removed only when nothing else uses it.
- The READY/BLOCKED verdict is computed from the final candidate, with exact blocking reasons.
- Sealed MPs are imported from their original `.mp`/`.mpb`, never from re-emitted XML. Exported `.xml` copies of sealed MPs are no longer pulled in as dependencies.
- Import runs in dependency order. `BatchDependsOn` comes from the final references, dependents of a failed MP are cascade-skipped, and the run writes `ImportResults_*.csv` and `ImportErrors_*.txt` (SCOM's full exception chain).
- New parameters: `-Manifest`, `-SourceInventory`, `-NonInteractive`.
- XML loading is encoding-safe, and `.mpb` bundles are supported.
- Static → dynamic group conversion (`-GroupConversionFile`):
  - NetBIOS-name rules for Windows computers.
  - `HostProperty` rules for hosted classes, in the same form the console wizard writes, so the groups stay editable.
  - Display-name rules for Health Service Watchers.
  - Nested subgroups are kept; `ExcludeList` becomes `DoesNotMatchRegularExpression`.
- Groups whose member class is missing in the target are dropped, with their dependents, so the rest of the MP imports (`-KeepBlockedGroups` turns this off).
- Per-object overrides (`ContextInstance`) are re-pointed through `-InstanceMapFile`; unmapped ones are removed and reported.
- The generated import script confirms the target management group (type `YES`, or use `-ExpectedManagementGroup`) and supports `-Reimport` and `-Only`.

### Export (Export-ScomEnvironment.ps1)
- Writes source/target inventories with sealed status and key token.
- Resolves manifest rows to MP IDs.
- Finds sealed originals.
- Resolves static group members while the source still knows them, proposes and tests NetBIOS patterns (`-GroupPatternMode Tight|Wildcard|Exact`), and takes host computer names from `FullName`.
- Resolves override `ContextInstance` objects.

### Step runner
- Steps: `Check`, `FixReview`, `EnableOverrides`, `MapInstances`, `Compile`, `WhyBlocked`, `OverrideReport`, `DryRun`, `Import`, `Reimport`, `TestGroups`, `Rollback`, `Collect`.

## 3.32 and earlier
Single-MP and batch compiler:
- dependency auto-resolution from a source repository or a live source management server
- topological ordering with cycle detection
- core-library minimum-version resolution
- dead-override stripping
- static group detection
- Microsoft MP catalog gap check (`-CheckCatalog`)
- `-AuditOnly`, `-LiveImport`, `-ExcludeMPIDs`, `-ForceRewriteCoreLibraries`
- plain-English bottom-line summary
