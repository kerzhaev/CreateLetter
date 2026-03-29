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
- Internal storage keys may live in workbook tables, but every user-facing surface such as history lists, exports, and search hints must convert them back to display labels before showing them.
- Workbook creation, export formatting, and other non-trivial data-output workflows must live in shared modules, not inside UserForms.
- Reusable caption builders, menu prompts, and status text assembly belong in shared modules once they stop being purely one-line control assignments.
- Hidden fallback literals such as unknown user names, unknown month markers, and attachment prefixes must be backed by localization keys instead of hardcoded strings in core modules.
- When renaming workbook ListObjects, keep a compatibility fallback for the old table name until the bootstrap script and smoke harness both recognize the new schema.
- COM VBA sync tooling must tolerate legacy exported module encodings during migration; do not assume every historical `.bas/.cls` file is already UTF-8-clean.
- Standard-module sync must strip exported attribute lines such as `Attribute VB_Name = "ModuleName"` before calling `CodeModule.AddFromString`; those lines are valid in exported files but invalid inside the VBE code pane.
- Treat `Attribute VB_Name = "..."` as export-only metadata. Never paste or inject it into the top of a module/class code pane manually or through automation; if a workflow targets VBE text insertion, strip all `Attribute VB_*` lines first.
- Class modules are not interchangeable with standard modules during VBA sync. For `.cls`, preserve or create a real class component in VBProject first; only the class body starting from `Option Explicit` belongs in the VBE code pane.
- If a class module must be created manually in VBE, insert `Class Module`, set its `(Name)` explicitly, paste only the class body, then export it back to `CreateLetter.xlsm.modules/` so the text source of truth stays aligned.
- Non-ASCII string literals may remain only for workbook compatibility fallbacks or localization content. Do not introduce them as new identifiers, enum names, constant names, procedure names, or persisted logic keys.
- A form may keep UI-only helpers such as control styling, listbox binding, focus management, and event routing. If code starts assembling reusable business text, parsing workbook schema, or updating persisted state, move it into a shared module.
- Before publishing or updating public docs, verify that file names, template names, and encoding guidance match the real repository state. Do not leave stale references to legacy Russian template names or Windows-1251 editor defaults after the UTF-8 migration.
- Keep local-only artifacts out of the public repository: restore points, temporary exports, runtime backup folders, and one-off manual recovery helpers should be ignored or removed before publication.
- The public GitHub repository is source-only: keep workbook binaries, template binaries, `.frx` resources, and local AI skill directories outside git while preserving local developer workflows through ignored files.
- GitHub CI for this project must validate repository consistency only. Do not design CI that depends on Excel COM or local runtime binaries being present in the public repository.
- Ribbon customization must be source-managed in `customUI/customUI.xml` and applied to the workbook through automation such as `scripts/apply_custom_ui.py`; do not rely on one-off manual RibbonX edits as the only source of truth.
- End-user UI surfaces must stay Russian-first: form captions, tooltips, summaries, export headers, status labels, and fallback `MsgBox` text should default to Russian even when internal identifiers remain English ASCII.
- Ribbon callback code may stay ASCII-safe internally, but the visible Ribbon tab/group/button labels, screentips, and user-facing dialog text must remain Russian and consistent with the workbook UI.
