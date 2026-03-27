[Back to README](../README.md)

# GitHub Publishing

## Publication Baseline

Before pushing the repository publicly:

1. Verify the workbook opens and macros compile successfully.
2. Run the automated smoke test:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_excel_smoke_test.ps1 -RequireLocalizationModule -RequireStructuredTables -RequireLocalizationSheet
```

3. Confirm that `CreateLetter.xlsm.modules/` is synchronized with `CreateLetter.xlsm`.
4. Confirm that local-only folders are not staged:
   - `filesarchive/`
   - `Backups/`
   - `_tmp_export/`

## Public Repository Hygiene

- Template files are expected to be named:
  - `LetterTemplate.docx`
  - `LetterTemplateFOU.docx`
- Exported VBA source is the text source of truth.
- `.frx`, `.xlsm`, and `.docx` are binary artifacts and should not be line-normalized.
- Workspace/editor configuration must remain UTF-8-oriented to avoid reintroducing encoding issues.

## Recommended Release Checklist

- Update `README.md` if setup steps or file names change.
- Update `AGENTS.md` and `.ai-factory/*` when structure or workflow changes.
- Keep user-facing strings in localization data rather than scattering them through VBA logic.
- Prefer additive compatibility aliases when renaming workbook-facing macros.
