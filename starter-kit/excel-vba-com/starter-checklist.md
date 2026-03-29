# Starter Checklist

Use this checklist when copying the starter kit into a new Excel/VBA repository.

## Repository baseline

- workbook binary exists, for example `Workbook.xlsm`
- exported source folder exists, for example `Workbook.xlsm.modules/`
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

## Early smoke gate

- workbook opens through COM
- expected sheets exist
- expected tables exist
- at least one known VBA public entry point can be inspected
- Ribbon package check passes if customUI is used

## Known pitfalls

- class modules imported like standard modules
- `Attribute VB_Name` injected into VBE text
- workbook file locked during ZIP inspection
- `.frm` without matching `.frx`
- legacy codepage exports breaking UTF-8 assumptions
