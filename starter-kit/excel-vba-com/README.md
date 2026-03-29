# Excel VBA COM Starter Kit

Reusable starter kit for source-managed Excel VBA projects.

This package is extracted from the working automation pattern used in CreateLetter and is intended to be copied into a new workbook repository as the initial baseline.

## What is included

```text
starter-kit/excel-vba-com/
├── .gitignore
├── customUI/
│   └── customUI.xml
├── scripts/
│   ├── apply_custom_ui.py
│   ├── create_restore_point.ps1
│   ├── export_vba_to_modules.py
│   ├── run_excel_smoke_test.ps1
│   └── sync_vba_from_modules.py
└── starter-checklist.md
```

## Intended repository layout

Copy these files into a new repository that roughly looks like this:

```text
.
├── Workbook.xlsm
├── Workbook.xlsm.modules/
├── scripts/
├── customUI/
└── filesarchive/
```

## First-time setup

1. Copy the starter kit files into the new repository.
2. Rename `Workbook.xlsm` references inside the scripts if your workbook file uses a different name.
3. Create the exported source folder, for example:

```text
Workbook.xlsm.modules/
```

4. Adjust the smoke script defaults:
   - required worksheet names
   - required table names
   - optional Ribbon requirement
5. Adjust `customUI/customUI.xml` if you want different Ribbon buttons.
6. Keep the workbook binary out of public git if you want a source-only repository.

## Recommended workflow

1. Edit `.bas/.frm/.cls` files in the source folder.
2. If the workbook was edited directly in Excel/VBE, export first:

```powershell
python .\scripts\export_vba_to_modules.py .\Workbook.xlsm .\Workbook.xlsm.modules
```

3. Run:

```powershell
python .\scripts\sync_vba_from_modules.py .\Workbook.xlsm .\Workbook.xlsm.modules
```

4. If you use Ribbon XML, run:

```powershell
python .\scripts\apply_custom_ui.py .\Workbook.xlsm
```

5. Run smoke:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\run_excel_smoke_test.ps1 -WorkbookPath .\Workbook.xlsm
```

6. If a change is risky, create a restore point first:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\create_restore_point.ps1 -Label feature-name
```

## Important limitations

- `.cls` files must be real class modules inside `VBProject`.
- workbook and worksheet document modules should also be exported and tracked as `.cls` files.
- `Attribute VB_*` lines are export metadata only; do not paste them into the VBE code pane.
- `.frm` files can depend on `.frx` resources; keep the matching `.frx` next to the form when needed.
- Ribbon package checks should inspect a temporary workbook copy, not a COM-opened locked file.

## Best use

This starter kit is strongest when you want:

- VBA source under git review;
- automated workbook sync through Excel COM;
- repeatable smoke tests;
- less reliance on manual VBE-only work.

For a fuller explanation of the approach, see:

- [../../docs/excel-vba-com-playbook.md](C:/Users/Nachfin/Desktop/Projets/CreateLetter/docs/excel-vba-com-playbook.md)
