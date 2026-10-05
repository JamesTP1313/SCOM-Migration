# Known limits

Read this page before you rely on the toolkit for a migration of your own.

## Versions

- **Tested path: SCOM 2016 → SCOM 2025.** That is the only pair used in production so far. `-SourceVersion` and `-TargetVersion` only change labels in v3.x. The knowledge of which Microsoft packs were retired and which packs are safe to version-bump is written into `MPMigration.ps1` (`$FamilyBehavior`, `$ReferencePatterns`, `$DeprecationNotices`) and reflects the 2016 → 2025 landscape. Other pairs will probably work for custom MPs, but the hints and families may be wrong. The plan to fix this is the [v4.0 roadmap](roadmap-v4.md).
- **One management group to one management group.** Merging several source groups into one target isn't modelled. Element-ID collisions between them would only show up at import.
- The sealed-original search looks in the SCOM 2016 install path. For other versions, pass the right folders with `-SealedSearchPath`.

## Validation

- **The compiler isn't SCOM's verifier.** It checks references, element existence and the structures it understands (overrides, groups, categories, folders, display strings, knowledge, dependencies). It doesn't run the full MP schema or the type checks that `Import-SCOMManagementPack` runs. A READY MP can still be rejected. When that happens, `ImportErrors_*.txt` has SCOM's exact reason, and nothing that depends on the rejected MP is attempted.
- **Development and regression testing are offline.** They use synthetic MPs and stand-in `OperationsManager` cmdlets (`tests/`). Real-world validation has come from a single production migration.
- **`.mpb` bundles are read on a best-effort basis** through the SCOM SDK Packaging assembly. A bundle that can't be read is logged and skipped. Resources inside bundles (images, scripts) aren't inspected.

## What isn't migrated

- Run As accounts and profiles; connectors; maintenance schedules; user roles; agent assignment and management-server failover; and Operations Console dashboards or reports that live outside MPs.
- Data: alerts, performance history, state history and the data warehouse.
- Notification channels, subscribers and subscriptions are handled by a separate script. Its limits are in [Notifications](notifications.md#known-limits).

## Rewrites the toolkit won't do

- **Re-pointing to replacement Microsoft packs.** A monitor, rule or class that targets a retired pack (for example SQL 2012) stays BLOCKED, because the successor packs use different element IDs. `-AllowIDRewrite` exists but isn't recommended: a re-pointed reference rarely resolves. Overrides on retired packs are removed, not translated.
- **Version-specific per-server overrides** can't be re-pointed when the object's class is gone in the target. For example, an override on a SQL 2012 database object of a specific server has nowhere to go and is removed.
- **Core library version bumps** (System, SystemCenter, IIS common library) are off by default. `-ForceRewriteCoreLibraries` turns them on, at your own risk.

## Groups

- Pattern matching is a .NET regular expression run by SCOM, and **case-sensitive**. NetBIOS names are normally upper-case. Watcher display names are matched in both the upper-case and lower-case forms of each listed name. Mixed-case names aren't covered.
- `HostProperty` conversion only works for classes **hosted on a Windows computer**. Members of unhosted or non-Windows classes (network devices, UNIX/Linux computers, distributed applications) are left static and reported in `StaticGroupMembership.csv`.
- `Tight` and `Wildcard` patterns are based on the name stem and can catch servers added later, which is the point of a dynamic group. Run `FixReview` to switch REVIEW rules to exact names, and check `GroupConversion.csv` before compiling.
- After import, a converted group stays empty until its agents report to the target and group calculation runs.

## Operational

- Windows PowerShell 5.1 is required on the management servers. PowerShell 7 can run the offline tests but can't load the SCOM 2016-era SDK.
- `Collect` output contains server, group and MP names. Scrub it before sharing.
- `CandidateMPs\*.SemanticRecommendations.csv` is a hint whose ordering can vary between runs when several successor packs score equally. It doesn't affect verdicts, candidates or import.
