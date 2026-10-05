# Contributing

Issues and pull requests are welcome. This toolkit touches production monitoring, so the bar is **don't change what it does without a test that shows it**.

## Reporting a problem

1. Say which source and target SCOM versions (and update rollups) you used, plus the toolkit version.
2. Say which step you ran and paste the one-screen summary it printed.
3. Attach the relevant rows of `BatchImportManifest.csv`, `StrippedElements.csv` or `ImportErrors_*.txt`.
4. **Scrub them first.** Replace server, group, MP and domain names. Never attach a `Collect` zip or an export folder to a public issue.

When the problem is a specific MP, the most useful thing you can send is a *minimal* MP XML that reproduces it, with names changed.

## Running the tests

No SCOM is needed:

```powershell
pwsh ./tests/Invoke-SmokeTest.ps1               # PowerShell 7, any OS
powershell -File .\tests\Invoke-SmokeTest.ps1   # Windows PowerShell 5.1
pwsh ./tests/Invoke-NotificationsTest.ps1       # notifications script
```

The smoke test:

- builds a working folder from `tests/fixtures/`;
- runs the source export with the stand-in module in `tests/mocks/source`;
- runs every non-destructive step with `tests/mocks/target`;
- asserts the verdicts, stripping, group conversion and override re-pointing.

Every PR must keep both tests green. A PR that changes behaviour must add a fixture MP and an assertion that cover the change.

## Code rules

The scripts run on locked-down management servers with **Windows PowerShell 5.1** and `Set-StrictMode -Version Latest`. They must keep working there.

- Don't use PowerShell 7-only syntax: no `??`, `?.`, `&&`/`||` pipeline chains, ternary `? :`, or `-Parallel`.
- Watch PS 5.1 array behaviour:
  - Wrap results whose count you need in `@(...)`.
  - Use `.ToArray()` on generic lists before you index them or pipe them.
  - Don't wrap a `List[T]` that a function returned in `@()`.
- Use `Out-String -Width 250` when you capture tables.
- Don't add aliases or one-letter function names. `R`, for example, is already `Invoke-History`.
- Don't add new module dependencies on the management server.
- Keep the source side **read-only**: `Get-*` and `Export-SCOMManagementPack` only.
- Anything that writes to SCOM must ask the operator to type `YES`, and must accept `-WhatIfOnly`/`-Force`-style switches.
- Every step must print a one-screen summary and write the detail to a file. Operators often can't scroll or copy from the console.
- Version-specific knowledge belongs in data, not code. This is the direction of v4.0; see `docs/roadmap-v4.md`.

## Line endings

`.gitattributes` keeps `.ps1`, `.psm1`, `.psd1` and `.csv` files as CRLF. Save scripts as UTF-8 with a BOM; Windows PowerShell 5.1 misreads non-ASCII characters in BOM-less files.
