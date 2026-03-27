# Project Rules

> Short, actionable rules and conventions for this project. Loaded automatically by $aif-implement.

## Rules

- Start every implementation change from a dedicated `pisces/<feature-name>` branch.
- Create a restore point with both workbook and exported modules before any repo-tracked edit.
- Implement migration work in small feature stages; never combine multiple major migration steps in one change set.
- Use AI Factory as the only delivery pipeline: `$aif-plan`, `$aif-implement`, `$aif-verify`, then merge.
- Do not wire `VbaModuleManager` to workbook open, save, or close events; keep import/export manual in the developer workflow.
- Treat `CreateLetter.xlsm.modules/` as the source of truth for VBA text artifacts.
- New internal identifiers, module names, constants, sheet keys, and placeholder keys must use English ASCII only.
- User-facing Russian text must come from localization data, not from new hardcoded literals in business logic.
- Do not start the next migration stage until the previous stage passes a manual smoke test in Excel.
- Every VBA module or form edited during an iteration must have its top header metadata updated: keep the module/form description accurate, bump the visible version number, and refresh the change date.
- Before asking the user for a manual VBA import, first try synchronizing the workbook automatically through the local COM tooling and run the smoke tests yourself.
- Workbook schema changes must be performed through a repeatable local script or explicit COM automation step, then verified automatically before asking the user to inspect anything.
- If a feature changes workbook structure, extend the smoke harness to assert the new structure so the same regression is caught automatically next time.
- When code already contains a workbook-backed fallback path, prefer bootstrapping the corresponding workbook sheet/table through automation instead of hardcoding more defaults in VBA.
- When a new workbook-backed localization or schema artifact becomes part of the baseline, add an explicit smoke-test flag for it instead of relying on ad-hoc manual checks.
- Use `scripts/create_restore_point.ps1 -Label "<feature-name>"` for restore points; do not substitute other parameter names.
- Do not drive business/status logic from localized English literals like `"RECEIVED"`; use neutral helpers or data-derived checks so Russian and English localizations behave the same.
- When localizing a workflow, include fallback document-generation text and runtime `MsgBox` paths in `ModuleMain`, not only form captions/tooltips.
- Combo-box display labels may be localized, but persisted workbook values and business-logic comparisons must use stable English ASCII internal keys.
- When migrating public VBA identifiers, add English-safe aliases first and keep legacy entry points as thin compatibility wrappers until workbook macros/buttons are switched over.
