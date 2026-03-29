[Back to README](../README.md)

# GitHub Publishing

## Publication Baseline

Before pushing the repository publicly:

1. Verify the workbook opens and macros compile successfully.
2. Run the automated smoke test:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_excel_smoke_test.ps1 -RequireLocalizationModule -RequireStructuredTables -RequireLocalizationSheet
```

3. Confirm that `CreateLetter.xlsm.modules/` and `CreateLetter.xlsm.document-modules/` are synchronized with the local `CreateLetter.xlsm`.
4. Confirm that local-only folders and binaries are not staged:
   - `filesarchive/`
   - `Backups/`
   - `_tmp_export/`
   - `.codex/`
   - `.agents/`
   - `CreateLetter.xlsm`
   - `LetterTemplate.docx`
   - `LetterTemplateFOU.docx`
   - `CreateLetter.xlsm.modules/*.frx`

## Public Repository Hygiene

- Local template files are expected to be named:
  - `LetterTemplate.docx`
  - `LetterTemplateFOU.docx`
- Exported VBA source is the text source of truth.
- `.frx`, `.xlsm`, and `.docx` are binary artifacts and are intentionally excluded from the public repository.
- Workspace/editor configuration must remain UTF-8-oriented to avoid reintroducing encoding issues.

## Recommended Release Checklist

- Update `README.md` if setup steps or file names change.
- Update `AGENTS.md` and `.ai-factory/*` when structure or workflow changes.
- Keep user-facing strings in localization data rather than scattering them through VBA logic.
- Prefer additive compatibility aliases when renaming workbook-facing macros.

## GitHub Automation

- The repository includes a GitHub Actions workflow that validates:
  - required source/docs files are present;
  - forbidden binary/runtime/local AI artifacts are not tracked;
  - public docs do not drift back to legacy template names or Windows-1251 guidance.
- Excel COM smoke tests remain local-only because the public repository does not include the runtime workbook binary.
