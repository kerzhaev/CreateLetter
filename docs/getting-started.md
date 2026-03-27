[Back to README](../README.md) · [Next Page ->](workflow.md)

# Getting Started

## Prerequisites

- Microsoft Excel with VBA macros enabled
- Access to `CreateLetter.xlsm`
- Local template files:
  - `LetterTemplate.docx`
  - `LetterTemplateFOU.docx`

## Project Files

- `CreateLetter.xlsm`: runtime workbook
- `CreateLetter.xlsm.modules/`: source-managed VBA modules and forms
- `filesarchive/`: workbook archive snapshots

## First Run

1. Open `CreateLetter.xlsm` in Excel.
2. Enable macros.
3. Run initialization macro if needed (`InitializeAllWorksheets` in `mdlInicialize`).
4. Open the main letter form (`frmLetterCreator`) from workbook UI/macro launcher.

## Verify Setup

- Workbook contains required sheets (addresses, letters, settings).
- Executor list is available in settings sheet.
- Templates are present and reachable from your workflow.

## Troubleshooting

- Macro disabled: re-open workbook and enable macros.
- Missing worksheet data: run initialization/reset helpers in `mdlInicialize.bas`.
- Invalid phone formatting: verify input and validation logic in `ModuleMain.bas`.

## See Also

- [Workflow](workflow.md) - User flow and form behavior
- [Configuration](configuration.md) - Sheet and template configuration
- [Maintenance](maintenance.md) - Safe update and backup process
