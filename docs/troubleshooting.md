# Troubleshooting

These are the problems hit during real runs, roughly in the order you'll meet them.

## Running the scripts

**"cannot be loaded … not digitally signed" / "running scripts is disabled"**
Files downloaded from the internet are blocked. In the working folder:

```powershell
Get-ChildItem *.ps1 | Unblock-File
Set-ExecutionPolicy -Scope Process Bypass
```

**A folder or file picker pops up**
You ran `MPMigration.ps1` directly without `-NonInteractive` and the paths. Use `Invoke-ScomMigrationStep.ps1 Compile` instead, or pass every path (see [Step reference](step-reference.md#mpmigrationps1-run-directly)).

**`Cannot find path …\Source\AllMPs` / "Need … run Compile first"**
The working folder layout doesn't match. Run `Check`: it prints what it found. Remember that `Target` must *contain* `AllMPs\`. If your folders have other names, use `Migration.settings.psd1`.

**A typed command fails with a strange parameter error**
Steps are positional words: `.\Invoke-ScomMigrationStep.ps1 Check`, not `-Check`. On a locked-down server where commands are typed by hand, prefer the step words over long `Import-Csv | Where-Object …` queries. `WhyBlocked`, `OverrideReport` and `TestGroups` exist so you don't have to type those queries.

**The output scrolls off the screen**
Every summary step writes its full detail to a file and prints only one screen. The file path is printed on the last line.

**"Exception setting 'ExtraInSource'"**
The source export came from a toolkit version that used the old column names (`ExtraIn2016`, `MatchesIn2016`). v3.40 and later accept both. Use the current `Invoke-ScomMigrationStep.ps1`, or re-export the source.

## Export

**`Get-SCOMManagementPack` fails, or connects to the wrong management group**
Run the export **on** each environment's own management server with its own console and module. A newer `OperationsManager` module can't connect to an older management group.

**Lots of rows in `SealedOriginalsMissing.csv`**
These are the original `.mp`/`.mpb` files for in-scope sealed MPs. Look on the old install media, the vendor's download page, the `System Center Management Packs` folder on old management servers, or your team's MP share. Drop them in `Source\SealedOriginals\` (sub-folders are fine) and run `Compile` again. Until then, those MPs and everything built on them are BLOCKED.

**A manifest row shows `NO MATCH`**
The `MatchPattern` doesn't match any MP ID or display name. Display names often differ from what people call the pack. Look the MP up in `SourceInventory.csv` and put its ID or a wildcard in `MatchPattern`.

## Compile

**The target folder "only contains N MP files"**
A complete target export normally has 100+ MPs. If you pointed at a partial folder, nearly everything resolves as missing. Re-run `Export-ScomEnvironment.ps1 -Role Target`.

**Most MPs are BLOCKED on Microsoft packs (`NotInTargetRepo`)**
Install the current Microsoft packs on the target first: Windows Server, SQL Server, IIS, Cluster, DNS, and so on. Then re-export the target and run `Compile` again. `MissingDependencies.csv` lists what is missing and how many MPs need each one.

**BLOCKED on SQL 2008/2012/2014, IIS 7/8, or Windows Server 2008/2012 packs**
These packs were retired, and their successors use different element IDs. Overrides on them are stripped automatically. A *monitor*, *rule* or *class* that targets them can't be carried over and has to be re-authored against the new pack. `CandidateMPs\<MP>.SemanticRecommendations.csv` names the likely successor pack.

**"SEALED in the source environment but only an exported .xml was supplied"**
See `SealedOriginalsMissing.csv` above. Never import the `.xml` export of a sealed MP.

**"Element 'X!Y' does not exist in the \<ver\> version of that MP"**
The pack is present, but that monitor, rule or class was removed or renamed in the newer version. If only overrides used it, they were stripped already. This message means a real element still depends on it.

## Groups

**A converted group is empty right after import**
This is normal. Membership is calculated on a schedule, and only for agents already reporting to the new management group. Run `TestGroups`:

- `WAITING`: none of the group's servers report to the target yet. Move the agents.
- `NOT CALCULATED YET`: the servers are there. Give it time (up to roughly an hour in a large management group), or restart the management server's Health Service.
- `BROKEN - bad server name`: see the next item.

**A group pattern shows `^(MICROSOFT)$` or `^(WINDOWS)$`**
That rule's members were hosted objects (for example an operating system object), whose *display name* is a description such as "Microsoft Windows Server 2016 Standard", not a server name. Current exports take the host computer name from the object's `FullName`. Re-export the source with the current `Export-ScomEnvironment.ps1`, then run `FixReview` and `Compile`. `FixReview` prints "Rules still showing a bad name: 0" when it is clean.

**Health Service Watcher groups are empty**
A watcher isn't hosted on the computer it watches, so a NetBIOS-name rule can never match it. `FixReview` switches these rules to a display-name match (the agent FQDN). If `StaticGroupMembers.csv` was missing when you ran `FixReview`, re-export the source and run it again.

**Import errors mentioning `HostProperty` or `MonitoringClass` in a group expression**
A pre-3.40 compile wrote an invalid `HostProperty` form. Re-compile with the current `MPMigration.ps1` and `Reimport`.

**A group shows a red "!" in the console wizard**
The group contains another group (a nested subgroup rule). The console wizard can't show that kind of rule. The group still works, and its regex rules can still be edited in the wizard.

**A group catches servers it shouldn't**
In `Source\GroupConversion.csv`, set an exact `Pattern` (`^(SERVER01|SERVER02)$`) or `Convert = N`, then run `Compile` and `Reimport`. A pattern you change in the console after import is kept until the MP is re-imported.

## Overrides

**`OverrideReport` says "Per-server overrides NOT mapped"**
Run `MapInstances` before `Compile`. Without `InstanceMap.csv`, overrides aimed at specific objects can't be re-pointed.

**Many overrides were removed as "server/object not in target"**
Those servers aren't agents in the new management group yet. After you move more agents, run `MapInstances` → `Compile` → `Reimport`.

## Import

**`AlreadyPresent` for an MP you changed**
The same version is already installed, so the import skips it. Use `Reimport`, which overwrites unsealed MPs at the same version. Sealed MPs are never overwritten at the same version.

**`SkippedDependencyFailed`**
An MP this one depends on failed earlier in the same run. Fix the first failure (see `ImportErrors_*.txt`) and import again.

**SCOM rejects an MP that compiled READY**
`ImportErrors_*.txt` has SCOM's full exception chain. Common causes are a duplicate element ID that is already defined in another installed MP, a display string or knowledge article pointing at something the compiler doesn't index, or schema rules the compiler doesn't check. Open an issue with the (scrubbed) error and the MP's structure.

## Rollback

`Rollback` removes only what the recorded import runs say they **imported**, newest first. An MP that something else now depends on can't be removed until that dependent is removed. Rollback retries up to three passes and logs anything it couldn't remove.
