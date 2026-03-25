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
