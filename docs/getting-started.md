[Back to README](../README.md) · [Next Page ->](workflow.md)

# Getting Started

## Prerequisites

- Microsoft Excel with VBA macros enabled
- Access to a local `CreateLetter.xlsm` runtime workbook
- Local template files:
  - `LetterTemplate.docx`
  - `LetterTemplateFOU.docx`

## Project Files

- `CreateLetter.xlsm.modules/`: source-managed VBA modules and forms
- `scripts/`: local automation helpers

## First Run

1. Provision a local `CreateLetter.xlsm` workbook and the required templates.
2. Sync the repository VBA into the workbook with `scripts/sync_vba_from_modules.py`.
3. Open the workbook in Excel.
4. Enable macros.
5. Run initialization macro if needed (`InitializeAllWorksheets` in `mdlInicialize`).
6. Open the main letter form (`frmLetterCreator`) from workbook UI/macro launcher.

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
