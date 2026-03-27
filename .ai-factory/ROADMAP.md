# CreateLetter Migration Roadmap

> High-level staged roadmap for the English-internal, UTF-8, and localization migration.

## Stages

- [x] Stage 1: AI Factory rules, branch policy, restore-point workflow, and manual `VbaModuleManager` contract.
- [x] Stage 2: Restore-point script and rollback documentation for workbook + modules snapshots.
- [x] Stage 3: Source-management baseline documented with `CreateLetter.xlsm.modules/` as the text source of truth and UTF-8 as the target artifact baseline.
- [x] Stage 4: Localization foundation (`Localization` data + string lookup API) without changing workbook schema.
- [ ] Stage 5: Comment translation and user-facing string extraction to localization keys.
- [ ] Stage 6: Internal key migration for placeholders, service strings, and hidden literals.
- [ ] Stage 7: Incremental English ASCII migration for code identifiers by bounded domain slices.
- [ ] Stage 8: Workbook schema migration for internal sheet, table, and helper names.
- [ ] Stage 9: Post-migration architecture cleanup for thin UserForms, consolidated repositories/services, and Word integration services.
  Keep forms as orchestration shells only, move business logic into existing modules where practical, and add no more than 1-2 new modules unless a later stage proves more are necessary.

## Delivery Gates

- Every stage starts from a dedicated `pisces/<feature-name>` branch.
- Every stage starts with a restore point.
- No stage is merged until it passes the manual Excel smoke test and `$aif-verify`.
- If UTF-8/source-management becomes unstable, later migration stages stop until that baseline is fixed.

## AIF Refactor Subplan

The next bounded feature stages for code quality and maintainability are:

- [x] `pisces/feature-7-schema-enums-and-columns`
  Introduce shared enums/constants for `Addresses`, `Letters`, `Settings`, and array-based record parts so creator/history logic stops depending on raw numeric indexes.
- [x] `pisces/feature-8-array-based-excel-repositories`
  Move address/history/settings reads toward in-memory `Variant` arrays to reduce worksheet roundtrips and prepare faster search/filter paths.
- [x] `pisces/feature-9-targeted-error-handling-pass`
  Replace risky `On Error Resume Next` in workbook/Word/export flows with targeted handlers while leaving harmless UI-formatting fallbacks lightweight.
- [x] `pisces/feature-10-word-app-singleton`
  Reuse one `Word.Application` instance per session instead of repeatedly creating/attaching per letter generation call.
- [x] `pisces/feature-11-listobjects-migration-readiness`
  Introduce helpers/accessors for `tblAddresses` and `tblLetters`, then gradually move CRUD/search/history flows onto `ListObjects`.
- [x] `pisces/feature-12-listobjects-bootstrap`
  Bootstrap `tblAddresses` and `tblLetters` in the workbook itself through a repeatable COM script, while preserving data layout and keeping code fallback-safe.
- [x] `pisces/feature-13-structured-smoke-diagnostics`
  Teach the smoke harness to verify structured tables explicitly so workbook schema regressions are caught automatically.
