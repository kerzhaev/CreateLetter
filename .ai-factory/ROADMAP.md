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
