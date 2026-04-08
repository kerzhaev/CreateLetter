# Starter Checklist

Use this checklist when copying the starter kit into a new Excel/VBA repository.

## Repository baseline

- workbook binary exists, for example `Workbook.xlsm`
- exported source folder exists, for example `Workbook.xlsm.modules/`
- exported document-modules folder exists, for example `Workbook.xlsm.document-modules/`
- `filesarchive/` exists or is allowed to be created
- `scripts/` exists
- `customUI/` exists if Ribbon is needed

## Before first sync

- confirm Excel `VBProject` access is enabled
- confirm `python` is available
- confirm `pywin32` is installed
- confirm workbook opens locally
- confirm workbook is not protected against VBA project edits

## Script customization

- replace default workbook name if it is not `Workbook.xlsm`
- update required worksheet names in the smoke script
- update required `ListObject` names in the smoke script
- update Ribbon callbacks and labels in `customUI.xml`
- update any project-specific naming in restore-point output
- decide which workbook and sheet document modules must be exported as tracked `.cls` files into `Workbook.xlsm.document-modules/`
- decide whether `sync_and_smoke.cmd` and `export_and_smoke.cmd` should be the default daily workflow entrypoints for the team
- decide whether `repair_workbook.cmd` should be documented as the team's standard recovery path after bad manual imports

## Early smoke gate

- workbook opens through COM
- expected sheets exist
- expected tables exist
- at least one known VBA public entry point can be inspected
- Ribbon package check passes if customUI is used
- workbook and sheet document modules export and sync cleanly

## Known pitfalls

- class modules imported like standard modules
- hidden workbook/sheet code not exported back to source
- hidden workbook/sheet code mixed into the same folder used by a generic manual module importer
- `Attribute VB_Name` injected into VBE text
- workbook file locked during ZIP inspection
- `.frm` without matching `.frx`
- legacy codepage exports breaking UTF-8 assumptions
