# How it works

This page explains what the toolkit does to your management packs, and why. You need it when a verdict surprises you.

## The problem

A management pack (MP) is an XML document that refers to elements in other MPs by `Alias!ElementID`. Each referenced MP is declared once in the manifest:

```xml
<Reference Alias="SQL12D">
  <ID>Microsoft.SQLServer.2012.Discovery</ID>
  <Version>6.7.2.0</Version>
  <PublicKeyToken>31bf3856ad364e35</PublicKeyToken>
</Reference>
```

When SCOM imports an MP, it checks every one of those references against what is installed **in the new management group**. Between SCOM releases a lot changes:

- Microsoft retires whole packs, such as SQL 2008–2014, IIS 7/8 and Windows Server 2012. Their replacements use new IDs and new element names.
- Packs that survive are re-versioned, and some elements inside them are removed or renamed.
- Every object, including every server, gets a **new GUID** in the new management group.
- A sealed MP's signature can't be recreated, so only the original file will do.

If you import a 2016-era MP as it is, you get one of three outcomes. SCOM refuses it, or it imports and monitors nothing (static groups, per-server overrides), or it imports half your batch before it hits the first broken dependency. The toolkit finds these cases *before* import.

## The compile model

`MPMigration.ps1` treats a migration like a compiler treats source code. It takes the full set of source MPs, the full set of target MPs and the manifest of what to move. For each MP it produces either an importable candidate or a precise error.

```
 Source\AllMPs  ─┐                                    ┌─ CandidateMPs\*.xml
 SealedOriginals ┤                                    ├─ BatchImportManifest.csv (READY/BLOCKED)
 SourceInventory ┼──► MPMigration.ps1 (Compile) ──────┼─ StrippedElements.csv
 Target\AllMPs  ─┤                                    ├─ GroupConversionResults.csv
 Manifest        ┤                                    ├─ MissingDependencies.csv
 GroupConversion ┤                                    └─ Import-Batch.<ver>.ps1
 InstanceMap    ─┘
```

1. **Ingest.**
   - Each source MP is read: `.xml` directly; `.mp` and `.mpb` through the SCOM SDK.
   - Only the MPs whose manifest row says `Migrate = Y` are kept. `ManifestMatch.csv` shows what each row matched.
2. **Index the target.**
   - Every MP in `Target\AllMPs` is indexed: its ID and version, and the ID of every class, relationship, monitor, rule, discovery, module type and so on that it defines.
   - This is the symbol table everything else is checked against.
3. **Resolve dependencies.** Each reference in each in-scope MP is classified:
   - **In the target**: the same ID at an equal or higher version. Because SCOM references are *minimum* versions, any newer version satisfies one. For example, Windows Library 7.5 → 10.19.
   - **In this batch**: another MP being migrated in the same run.
   - **A sealed dependency from the source** (`SealedDependencies\`). It's pulled into the batch and goes through the same checks.
   - **Missing**: `NotInTargetRepo` / `NotInstalledInTarget`.
4. **Strip what is dead** (see below).
5. **Decide the verdict** for the MP: READY or BLOCKED.
6. **Write the candidate**, with reference versions raised to what the target has, then **order the batch** topologically. A dependency cycle is reported, not silently ordered.
7. **Generate `Import-Batch.<ver>.ps1`** and `BatchImportManifest.csv`. `BatchDependsOn` comes from the *final* references, after stripping, so an MP is never held back by a dependency it no longer has.

The compile never connects to SCOM and never imports anything. You can re-run it as often as you like.

## Stripping: what is removed, and when

Some elements only *decorate* other elements: overrides, categories, folder items, display strings and knowledge articles. If the thing they point at no longer exists, they can be removed without changing what the MP monitors. The compiler removes them:

| Element | Removed when |
|---|---|
| Override (`*PropertyOverride`, `*ConfigurationOverride`, `DiagnosticPropertyOverride`…) | Its monitor/rule/discovery, or its context class, is in an MP the target doesn't have, or no longer exists in the target's version of that MP |
| Override with `ContextInstance` | The specific source object can't be found in the target (see [per-server overrides](#per-server-overrides)) |
| Category, folder item | Its target element is gone, as above |
| Display string, knowledge article | The element it describes was removed |

A **reference** is removed only when *nothing* left in the MP still uses its alias. If a monitor, rule, class, discovery, view or group still uses it, the reference stays and the MP is BLOCKED. The compiler never removes an element that does real monitoring just to make an import succeed.

Every removal is a row in `StrippedElements.csv` (MP, kind, element ID, reason). Per-MP copies go in `CandidateMPs\<MP>.Stripped.csv`.

### Groups that can't exist in the target

Suppose a group's membership rule targets a class from an MP the target doesn't have, such as a third-party pack you aren't bringing across. By default the compiler **drops that group** together with everything that points at it (its discovery, display strings, and the overrides and views that target it), so the rest of the MP can import. Dropped groups appear in `StrippedElements.csv` as `Group (DROPPED)` and in the `WhyBlocked` summary. `-KeepBlockedGroups` turns this off, so the whole MP is BLOCKED instead.

## Sealed vs. unsealed

| | Unsealed MP | Sealed MP (`.mp` / `.mpb`) |
|---|---|---|
| Who authored it | You: overrides, groups, custom monitoring | Microsoft, vendors, or your own libraries |
| Can other MPs reference it? | No | Yes; it has a `PublicKeyToken` |
| What the toolkit imports | A rewritten candidate `.xml` | **The original file, unchanged** |

`Export-SCOMManagementPack` writes a sealed MP out as plain `.xml`. That XML can't be imported in its place: the signature is gone, so SCOM would treat it as a different, unsealed MP, and every MP that references it by key token would fail. That's why:

- `SourceInventory.csv` records which source MPs were sealed.
- Any in-scope sealed MP supplied only as `.xml` is **BLOCKED** ("SEALED in the source environment but only an exported .xml was supplied"). Anything that depends on it is BLOCKED with `BatchMemberNeedsSealedOriginal`.
- The source export searches `-SealedSearchPath` and the usual install folders for the original files. It copies the in-scope ones to `SealedOriginals\` and lists the rest in `SealedOriginalsMissing.csv`.

## READY / BLOCKED

An MP is **READY** when every reference it still has after stripping resolves to the target, to another READY member of the batch, or to a sealed original that's being imported. Otherwise it's **BLOCKED**, and `BlockingReasons` says exactly why, for example:

```
Needs Microsoft.SQLServer.2012.Discovery v6.7.2.0 [NotInTargetRepo] -- still used by 1 element(s): UnitMonitor 'Contoso.EHR.DBMon'
Element 'Microsoft.Windows.Server.2016.Monitoring!...RemovedClass' does not exist in the 2025 version of that MP -- used by 1 element(s): Rule 'Contoso.R1'
SEALED in the source environment but only an exported .xml was supplied -- add the original .mp/.mpb ...
Needs Contoso.Custom.Library v2.0.0.0 [BatchMemberNeedsSealedOriginal] -- ...
```

`Invoke-ScomMigrationStep.ps1 WhyBlocked` groups these into one screen.

When the import script runs, it imports READY rows in order and skips BLOCKED ones. It also **cascade-skips**: if an MP fails, everything that depends on it is recorded as `SkippedDependencyFailed` rather than attempted. `-IncludeBlocked` tries the BLOCKED rows too, which is useful when you want SCOM's own error text for them.

### Why BLOCKED MPs aren't re-pointed automatically

It is tempting to rewrite `Microsoft.SQLServer.2012.Discovery` to `Microsoft.SQLServer.Windows.Discovery`. But the replacement packs renamed their classes and monitors. `Microsoft.SQLServer.2012.Database` has no same-named counterpart, so a rewritten reference would point at elements that don't exist. The compiler can do this rewrite (`-AllowIDRewrite`), and it's off by default because it rarely produces a working MP. `CandidateMPs\<MP>.SemanticRecommendations.csv` lists the likely successor pack as a hint for whoever re-authors the monitor.

## Static groups → dynamic groups

A static group stores its members as source object GUIDs:

```xml
<MembershipRule>
  <MonitoringClass>$MPElement[Name="Windows!Microsoft.Windows.Computer"]$</MonitoringClass>
  <RelationshipClass>$MPElement[Name="SC!Microsoft.SystemCenter.ComputerGroupContainsComputer"]$</RelationshipClass>
  <IncludeList>
    <MonitoringObjectId>11111111-...</MonitoringObjectId>
    <MonitoringObjectId>66666666-...</MonitoringObjectId>
  </IncludeList>
</MembershipRule>
```

Those GUIDs don't exist in the new management group, so the group imports cleanly and stays **empty forever**. The conversion runs in three stages.

**1. Source export.** `Export-ScomEnvironment.ps1 -Role Source` does the following while the source still knows the objects:

- It resolves every member GUID to its server name (`StaticGroupMembers.csv`).
- It proposes a NetBIOS-name pattern per rule (`GroupConversion.csv`). There are three modes:
  - `Tight` (default): name stem + digits, e.g. `^(APPSQL[0-9]+)$`.
  - `Wildcard`: e.g. `APPSQL*`.
  - `Exact`: the listed names only.
- It tests each pattern against **every** Windows computer in the source. A pattern that would also catch servers that aren't in the group today is marked `REVIEW`, with the extra names listed.

**2. FixReview.** `Invoke-ScomMigrationStep.ps1 FixReview` rewrites every REVIEW rule to the exact current member names. It also switches rules whose members are Health Service Watchers to a display-name match, because a watcher isn't hosted on the computer it watches. You can still edit `GroupConversion.csv` by hand: change `Pattern`, or set `Convert = N`.

**3. Compile.** For each rule with `Convert = Y`, the IncludeList is replaced with a dynamic expression of the same form the console's group wizard writes, so the group stays **editable in the console**:

| Member class | Expression |
|---|---|
| `Microsoft.Windows.Computer` (and subclasses) | `Property` `NetbiosComputerName` matches the pattern |
| Any class hosted on a Windows computer (OS, logical disk, SQL instance…) | `HostProperty` → `Microsoft.Windows.Computer` / `NetbiosComputerName` |
| Health Service Watcher | `Property` `DisplayName` matches `^(NAME\|name)(\..*)?$` |
| A nested static subgroup (singleton group class) | The IncludeList is removed; the containment rule stays |

An `ExcludeList` becomes an `And` with `DoesNotMatchRegularExpression` on the excluded names.

`GroupConversionResults.csv` lists every rule that was converted, left alone or dropped. After import, `TestGroups` reports for each converted group whether it is **WORKING**, **WAITING** (none of its servers report to the target yet), **NOT CALCULATED YET** (the servers are there but group calculation hasn't run), or **BROKEN**.

Group membership populates only as agents report to the new management group, and group calculation runs on its own schedule. An empty group right after import is normal.

## Per-server overrides

"Disable this monitor on SERVER01" is stored as an override with a `ContextInstance` GUID:

```xml
<MonitorPropertyOverride ID="..." Context="Windows!Microsoft.Windows.Computer"
    ContextInstance="eeeeeeee-2222-..." Monitor="..." Property="Enabled">
```

That GUID points at nothing in the new management group. The toolkit re-points it in three stages:

1. **Source export.** Every `ContextInstance` GUID in an in-scope MP is resolved to the object's class, `FullName` (class + key, e.g. `Microsoft.Windows.Computer:server01.contoso.com`) and display name, and written to `OverrideInstances.csv`.
2. **`MapInstances`** runs on the target and is read-only. For each object it looks up the same class in the target and matches by `FullName`. If that fails, it matches by display name, but only when that name is unambiguous. The result is `InstanceMap.csv` (old GUID → new GUID, and how it was matched).
3. **Compile.** Mapped overrides get the new GUID. Unmapped ones are removed and listed in `StrippedElements.csv`: the server is retired, isn't an agent in the target yet, or its object type no longer exists. Removing them is deliberate. An override aimed at a GUID that doesn't exist imports without error and silently does nothing.

If you add agents to the target later, run `MapInstances` → `Compile` → `Reimport` again to pick up their overrides.

`OverrideReport` summarizes all of this for the override packs: how many overrides were kept, how many were re-pointed, how many were removed and why, and which missing MPs account for most of the removals.
