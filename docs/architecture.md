[← Previous Page](workflow.md) · [Back to README](../README.md) · [Next Page →](configuration.md)

# Architecture

## Pattern

CreateLetter follows a **modular monolith** pattern inside one Excel workbook.

## Modules

- UI Layer:
  - `frmLetterCreator.frm`
  - `frmLetterHistory.frm`
- Core Logic:
  - `ModuleMain.bas`
  - `ModuleDates.bas`
- Initialization:
  - `mdlInicialize.bas`
- Cross-cutting:
  - `ModuleBackup.bas`
  - `ModuleAuditLogger.bas`
  - `ModuleCache.bas`
  - `ModuleLocalization.bas`

## Data Layer

Workbook worksheets act as persistence:
- addresses
- letters
- settings

## Dependency Rules

- UserForms can call core modules.
- Core modules should not depend on specific form control state.
- Backup/audit helpers should remain reusable and side-effect explicit.

## Source of Truth

- Runtime artifact: `CreateLetter.xlsm`
- Source-managed VBA: `CreateLetter.xlsm.modules/`
- Developer sync mode: manual `VbaModuleManager` invocation only
- Localization foundation: `ModuleLocalization.bas` with built-in fallback values and optional workbook-backed overrides

## Migration Constraints

- Do not add workbook open/save/close hooks for module import/export.
- Keep runtime user behavior independent from source-management tooling.
- Migrate internal identifiers toward English ASCII in small feature stages only.
- Move user-facing Russian text toward localization data instead of new hardcoded literals.

For full architecture rationale and examples, see `.ai-factory/ARCHITECTURE.md`.

## See Also

- [Workflow](workflow.md) - Operational flow through forms and modules
- [Configuration](configuration.md) - Data and environment settings
- [Development Workflow](development-workflow.md) - AI Factory and branch-level delivery process
- [Maintenance](maintenance.md) - Safe change management process
