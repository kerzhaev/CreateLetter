[← Previous Page](configuration.md) · [Back to README](../README.md) · [Next Page →](maintenance.md)

# Development Workflow

## AI Factory Delivery Policy

All project work now follows the AI Factory pipeline:

1. Run `$aif-plan` for a single feature stage.
2. Create a feature branch named `pisces/<feature-name>`.
3. Create a restore point before editing repo-tracked files.
4. Run `$aif-implement` for the scoped stage only.
5. Perform a manual Excel smoke test.
6. Run `$aif-verify`.
7. Merge into `main` only after the stage passes.
8. Push to GitHub after merge.

Do not bundle multiple migration stages into one implementation cycle.

## Branching Rules

- One migration stage per branch.
- Branch names must use the `pisces/` prefix.
- Keep `main` as the stable integration branch.
- If a stage fails the smoke test, restore from the snapshot and fix it in the same feature branch or discard the branch entirely.

## Restore Point Workflow

Before any change:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\create_restore_point.ps1 -Label feature-name
```

Each restore point must contain:

- `CreateLetter.xlsm`
- `CreateLetter.xlsm.modules/`
- `CreateLetter.xlsm.document-modules/`

Restore points are stored under `filesarchive/restore-point-<label>-<timestamp>/`.

## Manual VbaModuleManager Workflow

This project uses source-managed VBA text artifacts plus Excel COM automation.

- Do not attach import/export to `Workbook_Open`.
- Do not attach export to `Workbook_BeforeSave`.
- Do not attach cleanup to `Workbook_Close`.
- Run module import/export manually as part of the developer workflow.

Developer contract:

1. Prefer exporting and synchronizing VBA automatically before asking for manual import:

```powershell
python .\scripts\export_vba_to_modules.py .\CreateLetter.xlsm .\CreateLetter.xlsm.modules .\CreateLetter.xlsm.document-modules
```

2. Synchronize modules/forms/document modules back into the workbook before testing VBA changes in Excel:

```powershell
python .\scripts\sync_vba_from_modules.py .\CreateLetter.xlsm .\CreateLetter.xlsm.modules .\CreateLetter.xlsm.document-modules
```

3. If automatic export/sync fails because of COM/VBProject access or a workbook-specific import edge case, fall back to the manual modified `VbaModuleManager` only for `CreateLetter.xlsm.modules/`.
4. Export modules back to `CreateLetter.xlsm.modules/` and document modules back to `CreateLetter.xlsm.document-modules/` after validated workbook-side changes.
5. Verify that workbook runtime behavior does not depend on automatic source-management hooks.

Document-module note:

- Workbook and worksheet document modules are part of the source-managed contour.
- Keep `.cls` exports for `ЭтаКнига`, `Лист1`, `Лист2`, `Лист3`, `Лист4` in `CreateLetter.xlsm.document-modules/`.
- Changes made only in hidden workbook/sheet code must be exported back into the repository before the feature is considered complete.
- Do not point manual `VbaModuleManager` import at `CreateLetter.xlsm.document-modules/`; those files must be updated through COM sync so existing document modules are edited in place.

Class-module note:

- `.cls` files must exist as real VBA class components, not as standard modules with class text pasted into them.
- If a new class module is introduced, either let the sync helper create/update a real class component or create it manually via `Insert -> Class Module`, set its `(Name)`, paste only the class body, then export it back to `CreateLetter.xlsm.modules/`.
- Do not paste `VERSION 1.0 CLASS`, `BEGIN/END`, `MultiUse`, or `Attribute VB_*` lines into the VBE code pane.

## Workbook Schema Automation

When a feature changes workbook structure itself, prefer a repeatable script over a one-off manual edit.

Current helper:

```powershell
python .\scripts\ensure_workbook_tables.py .\CreateLetter.xlsm
```

This helper creates `tblAddresses` and `tblLetters` if they are missing and leaves existing data intact.

Localization bootstrap helper:

```powershell
python .\scripts\ensure_localization_sheet.py .\CreateLetter.xlsm
```

This helper materializes the workbook `Localization` sheet from built-in translations in `ModuleLocalization.bas`.

## Source of Truth and Encoding Policy

- `CreateLetter.xlsm.modules/` is the source of truth for standard VBA modules, class modules, and forms.
- `CreateLetter.xlsm.document-modules/` is the source of truth for workbook and worksheet document modules.
- Workbook behavior must remain stable even if source-management tooling is not invoked by end users.
- UTF-8 is the target baseline for text artifacts consumed by Git and AI agents.
- If an import/export step produces ANSI or CP1251-only artifacts, stop the migration and stabilize source-management before changing business logic.

## Smoke Test Gate

Run this minimum smoke test after each feature stage:

- Workbook opens without macro errors.
- Letter creator form loads.
- Address lookup works.
- Attachment selection works.
- Word generation works.
- Letter history save/load works.
- Backup and audit behavior still works.

Automation helper:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_excel_smoke_test.ps1
```

Use `-RequireLocalizationModule` after importing updated modules into the workbook for localization stages.
Use `-RequireStructuredTables` after workbook schema stages that depend on `tblAddresses` / `tblLetters`.
Use `-RequireLocalizationSheet` after workbook-backed localization becomes part of the expected schema.
Use `-RequireRibbonCustomization` after source-managed Ribbon changes or package customization work.
Use `-RequireAddressGroupColumn` after address-schema stages that depend on the optional `AddressGroup` column in `tblAddresses`.
The smoke harness also verifies that source files exist for all workbook and worksheet document modules in `CreateLetter.xlsm.document-modules/`.

## Reusable COM Pattern

This repository now includes a reusable Excel/VBA automation playbook:

- [Excel VBA COM Playbook](excel-vba-com-playbook.md)

Use it as the baseline for future workbook projects when you want:

- source-managed VBA;
- COM-based sync into `.xlsm`;
- schema/bootstrap scripts;
- repeatable smoke tests;
- less reliance on manual VBE-only workflows.

## See Also

- [Configuration](configuration.md) - Workbook, template, and MCP settings
- [Maintenance](maintenance.md) - Recovery and safe update checklist
- [Architecture](architecture.md) - Module boundaries and migration constraints
- Create restore points with `powershell -ExecutionPolicy Bypass -File .\scripts\create_restore_point.ps1 -Label "<feature-name>"`.
- The `Addresses` schema supports an optional `AddressGroup` field for scenarios where different recipients share one postal address but must stay separate as named addressees.
