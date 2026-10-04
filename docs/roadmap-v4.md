# Roadmap: v4.0, any supported SCOM to any supported SCOM

**Goal:** the same toolkit, with no code changes, migrates custom MPs from **SCOM 2012 R2, 2016, 2019 or 2022** to **SCOM 2019, 2022 or 2025**, and says clearly when a pair isn't supported.

**Non-goal:** changing how the compile works. The compile model (strip, verdict, order, convert groups, re-point overrides) is version-independent, and it has been proven on a production migration. v4.0 moves the version-*specific* knowledge out of the code and makes the version-*detection* automatic.

## 1. What is version-specific today

| Where | What | Problem |
|---|---|---|
| `MPMigration.ps1` `$FamilyBehavior` | Which MP families are Unified / PerVersion / ExactMatchOnly | Written against the 2025-era catalog. For example, "SQL covers 2014–2025+" isn't true for a 2019 target with an old SQL MP. |
| `MPMigration.ps1` `$ReferencePatterns` | Retired pack → successor pack (IIS 2008 → unified IIS, Server 2012 → Base OS, AD 2008 → ADDS, SQL 2012 → …) | One table for every pair, so it can't say "on a 2019 target the Windows Server 2016 MP is still current". |
| `MPMigration.ps1` `$DeprecationNotices` | End-of-support notices | Dated text, written into the code |
| `MPMigration.ps1` `-CatalogView 'sc-om-2025'` | Microsoft Learn catalog view | Should follow the target version |
| `Export-ScomEnvironment.ps1` | Sealed-original search path `Microsoft System Center 2016\Operations Manager` | Only right for 2016 |
| `Invoke-ScomMigrationStep.ps1` | `-SourceVersion 2016` / `-TargetVersion 2025` defaults | Labels only. They should be detected, not typed. |
| SDK loading | Assumes the local MS's `Microsoft.EnterpriseManagement.*` assemblies | Has to work with whichever version is installed. `.mpb` reading differs slightly between versions. |

## 2. Detect versions at export time

`Export-ScomEnvironment.ps1` records the environment it ran against in a new `Environment.json` next to the inventory:

```json
{
  "Role": "Source",
  "ManagementGroup": "MG_OLD",
  "ScomVersion": "2016",
  "ScomBuild": "7.2.12335.0",
  "UpdateRollup": "UR10",
  "SdkVersion": "7.2.11719.0",
  "ExportedAt": "2026-10-04T09:58:00Z",
  "ToolkitVersion": "4.0.0"
}
```

Detection order (first one that works wins):

1. The build version of the management server's `HealthService` / `Microsoft.EnterpriseManagement.Core.dll`, mapped to a release (for example 7.1 = 2012 R2, 7.2 = 2016, 10.19 = 2019, 10.22 = 2022). The mapping lives in the data file.
2. The version of the `Microsoft.SystemCenter.Library` MP installed in the management group. It's present everywhere and its major/minor tracks the release.
3. `-ScomVersion` given by hand: a last resort, recorded as `"DetectedBy": "manual"`.

`MPMigration.ps1` and the step runner read both `Environment.json` files. If they're missing (a v3 export), they fall back to the labels. `-SourceVersion` and `-TargetVersion` stay available as overrides. Before any work, `Check` prints the detected pair and whether it's supported.

## 3. Move the knowledge into a data file

There will be one versioned file, `data/mp-knowledge.psd1`. It's a PowerShell data file, so it loads safely with `Import-PowerShellDataFile` on 5.1 and needs no JSON quirks. Contributors can update it **without touching code**. The values in this sketch are illustrative; each one gets verified and sourced in the real file.

```powershell
@{
    SchemaVersion = 1
    ScomReleases  = @{
        '2012R2' = @{ BuildPrefix = '7.1';   SystemCenterLibrary = '7.1.10226' }
        '2016'   = @{ BuildPrefix = '7.2';   SystemCenterLibrary = '7.2.11719' }
        '2019'   = @{ BuildPrefix = '10.19'; SystemCenterLibrary = '10.19.10050' }
        '2022'   = @{ BuildPrefix = '10.22'; SystemCenterLibrary = '10.22.10056' }
        '2025'   = @{ BuildPrefix = '10.25'; SystemCenterLibrary = '10.25.x' }
    }
    SupportedPairs = @(
        @{ Source = '2012R2'; Target = '2019', '2022'      ; Status = 'Supported' }
        @{ Source = '2016'  ; Target = '2019', '2022', '2025'; Status = 'Supported' }
        @{ Source = '2019'  ; Target = '2022', '2025'      ; Status = 'Supported' }
        @{ Source = '2022'  ; Target = '2025'              ; Status = 'Supported' }
        @{ Source = '2012R2'; Target = '2025'              ; Status = 'Experimental' }  # widest gap; most retirements
    )
    Families = @(
        @{ Name = 'Windows'; Behavior = 'Unified'
           Retired = @(
             @{ Id = 'Microsoft.Windows.Server.2012.Monitoring'; RetiredIn = '2019'; Successor = 'Microsoft.Windows.Server.Monitoring' }
           ) }
        # ...
    )
    Deprecations = @(
        @{ IdPrefix = 'Microsoft.SQLServer.Reporting'; EndOfSupport = '2027-01'; Note = '...'; Source = 'https://...' }
    )
    SealedSearchPaths = @(
        '%ProgramFiles%\Microsoft System Center {Version}\Operations Manager'
        '%ProgramFiles%\Microsoft System Center\Operations Manager'
        '%ProgramFiles(x86)%\System Center Management Packs'
    )
    CatalogView = @{ '2019' = 'sc-om-2019'; '2022' = 'sc-om-2022'; '2025' = 'sc-om-2025' }
}
```

Rules:

- Every successor mapping is **keyed by target version**. A retirement only applies when `RetiredIn <= Target`. That's how "the Windows Server 2016 MP is still current on 2019" is expressed.
- Behaviour must not change for 2016 → 2025: the v3.40 tables are converted one-for-one, and the smoke test proves the output is identical.
- `Get-Help`-style comments in the file say where each fact came from, so reviewers can verify it.
- Use ordered collections throughout. v3 iterates `@{}` hashtables, which is why the `SemanticRecommendations` hint can vary between runs on PowerShell 7.
- The knowledge stays **advisory**. The verdict comes from what the target export actually contains, never from the data file. A wrong entry can make a hint wrong, but it can't make a BLOCKED MP import.

## 4. Optional: element-ID maps for successor packs

This is where v4 can go beyond "BLOCKED: re-author it". `data/element-maps/<from-mp>__<to-mp>.csv` would hold verified element mappings:

```
FromElement,ToElement,Kind,Confidence,VerifiedBy
Microsoft.Windows.Server.2012.LogicalDisk,Microsoft.Windows.Server.10.0.LogicalDisk,Class,High,<github handle>
```

- Used only behind an explicit switch (`-UseElementMaps`), only for overrides and categories at first, and every rewrite is logged.
- Maps are contributed and reviewed one pack pair at a time. No map means today's behaviour (strip and report).

## 5. Support matrix to validate

| Source ↓ / Target → | 2019 | 2022 | 2025 |
|---|---|---|---|
| 2012 R2 | build & test | build & test | experimental |
| 2016 | build & test | build & test | **proven (v3.40)** |
| 2019 | — | build & test | build & test |
| 2022 | — | — | build & test |

To mark a pair "Supported", it needs:

1. Smoke-test fixtures for that pair: a target fixture set built from the real Microsoft library MP versions of that release.
2. At least one real run, reported by a contributor, with `Check`, `WhyBlocked`, `OverrideReport` and import results (scrubbed).

SDK notes:

- Each export runs on its own environment's management server with its own module, as it does today, so no cross-version SDK calls are needed.
- The compile runs on the target MS. `.mp` reading uses the target's SDK, which reads older sealed MPs. That's confirmed for 2016 MPs on 2025 and still needs confirming for 2012 R2 MPs on 2019.
- 2012 R2 management servers may only have PowerShell 4.0. The export script must stay 4.0-compatible, or the documentation must say to install WMF 5.1.

## 6. Engineering work

| # | Item | Notes |
|---|---|---|
| 1 | `Environment.json` + detection in the export | §2 |
| 2 | `data/mp-knowledge.psd1` + loader; v3 tables converted 1:1 | §3. Byte-identical output on the existing smoke test. |
| 3 | Pair check in `Check` / `Compile` | Unsupported pair → clear message, `-Force` to continue |
| 4 | Pester 5 tests replacing `Invoke-SmokeTest.ps1` assertions | Keep the smoke test as an end-to-end test; add unit tests for stripping, the verdict, group XML and instance mapping |
| 5 | Fixture sets per target release | `tests/fixtures/target-2019`, `-2022`, `-2025` |
| 6 | CI | GitHub Actions on `windows-latest` (Windows PowerShell 5.1 + pwsh 7) and `ubuntu-latest` (pwsh 7); PSScriptAnalyzer |
| 7 | Split `MPMigration.ps1` (≈5,800 lines) into a module | `ScomMigration.psm1` + thin scripts. Do this after 1–6, so behaviour is pinned by tests first. |
| 8 | `-UseElementMaps` | §4, optional |
| 9 | Signed release zip + PowerShell Gallery package | Many SCOM servers require signed scripts |

Items 1 to 3 are the minimum for v4.0. Items 4 to 6 come before anyone calls a new pair "Supported". Items 7 to 9 can come in 4.x.
