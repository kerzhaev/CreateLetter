# CreateLetter Migration Roadmap

> High-level staged roadmap for the English-internal, UTF-8, and localization migration.

## Stages

- [x] Stage 1: AI Factory rules, branch policy, restore-point workflow, and manual `VbaModuleManager` contract.
- [x] Stage 2: Restore-point script and rollback documentation for workbook + modules snapshots.
- [x] Stage 3: Source-management baseline documented with `CreateLetter.xlsm.modules/` as the text source of truth and UTF-8 as the target artifact baseline.
- [x] Stage 4: Localization foundation (`Localization` data + string lookup API) without changing workbook schema.
- [x] Stage 5: Comment translation and user-facing string extraction to localization keys.
- [x] Stage 6: Internal key migration for placeholders, service strings, and hidden literals.
- [x] Stage 7: Incremental English ASCII migration for code identifiers by bounded domain slices.
- [x] Stage 8: Workbook schema migration for internal sheet, table, and helper names.
- [x] Stage 9: Post-migration architecture cleanup for thin UserForms, consolidated repositories/services, and Word integration services.
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
- [x] `pisces/feature-14-localization-sheet-bootstrap`
  Materialize the workbook `Localization` sheet from `ModuleLocalization.bas` so localization can evolve in workbook data instead of only built-in VBA defaults.
- [x] `pisces/feature-15-localization-smoke-diagnostics`
  Extend the smoke harness with an explicit `Localization` worksheet gate so workbook-backed localization regressions are caught automatically.
- [x] `pisces/feature-16-localize-creator-ui-slice`
  Move `frmLetterCreator` captions, tips, dialogs, and creator/address validation messages onto workbook-backed localization keys while keeping English fallbacks safe.
- [x] `pisces/feature-17-localize-history-ui-slice`
  Move `frmLetterHistory` captions, tooltips, search/export/status dialogs, and history status text onto workbook-backed localization keys without relying on English-only status parsing.
- [x] `pisces/feature-18-localize-core-runtime-messages`
  Move remaining `ModuleMain` runtime messages and fallback Word letter text onto workbook-backed localization keys so localization no longer stops at the form layer.
- [x] `pisces/feature-19-internal-key-migration`
  Separate localized UI labels from persisted internal values for document types, letter types, and Word placeholder identifiers so business logic no longer depends on display text.
- [x] `pisces/feature-20-entrypoint-identifier-migration`
  Add English-safe public aliases for workbook bootstrap and letter-creator entry points while preserving legacy macro names as compatibility wrappers.
- [x] `pisces/feature-21-hide-internal-keys-from-user-outputs`
  Keep internal ASCII keys in workbook storage while converting history search/export surfaces back to localized display labels for operators.
- [x] `pisces/feature-22-history-output-refactor`
  Move history export and search-hint output workflows from `frmLetterHistory` into `ModuleMain` so the form remains a thin orchestration shell.
- [x] `pisces/feature-23-creator-summary-and-menu-refactor`
  Move creator progress/counter captions and document-action menu text assembly into `ModuleMain` so `frmLetterCreator` keeps only UI wiring.
- [x] `pisces/feature-24-hidden-literals-cleanup`
  Replace remaining hardcoded fallback literals in core helpers with localization-backed keys for unknown-user, unknown-month, and attachment-prefix paths.
- [x] `pisces/feature-25-letter-texts-table-migration`
  Rename the Settings sheet text table to `tblLetterTexts`, keep fallback compatibility for the legacy `Text` table name, and teach bootstrap/smoke tooling to validate the new schema.
- [x] `pisces/feature-26-sync-fallback-legacy-encoding`
  Make the COM sync helper tolerant to legacy exported VBA file encodings, normalize the corrupted cache module source, and verify that schema-migration changes compile cleanly after workbook sync.
- [x] `pisces/feature-27-legacy-modules-comment-cleanup`
  Normalize legacy helper modules to clean UTF-8 source, replace garbled comments with English headers/comments, and preserve runtime behavior through automated workbook sync and smoke validation.
- [x] `pisces/feature-28-normalize-exported-source-encoding`
  Re-normalize source-managed VBA exports after manual VbaModuleManager roundtrips so the repository returns to a stable UTF-8 baseline before further feature work.
- [x] `pisces/feature-29-localize-maintenance-runtime-messages`
  Move backup and audit helper user-facing runtime/admin messages onto workbook-backed localization keys and refresh the Localization sheet automatically.
- [x] `pisces/feature-30-maintenance-entrypoint-aliases`
  Add English-safe public aliases for legacy maintenance/admin entry points, then harden the VBA sync workflow against duplicate-name and VBE attribute insertion regressions.
- [x] `pisces/feature-31-bootstrap-maintenance-localization`
  Move workbook-bootstrap and legacy snapshot helper prompts/messages onto workbook-backed localization keys so maintenance flows follow the same localization contract as the main runtime.
- [x] `pisces/feature-32-final-identifier-and-stage-closure`
  Remove the last non-ASCII compatibility identifier names from core modules, move residual history caption builders into shared logic, and formally close the staged migration roadmap after a final thin-form audit.
- [x] `pisces/feature-33-github-publication-prep`
  Prepare the repository for GitHub publication by aligning public docs with real file names, switching workspace defaults to UTF-8, removing obsolete manual-recovery artifacts, and ignoring local-only runtime/export folders.
- [x] `pisces/feature-34-public-source-only-github`
  Convert the public repository to a source-only layout by removing local AI skill bundles and binary runtime assets from git, adding an MIT license, and documenting the local provisioning contract for workbook/template artifacts.
- [x] `pisces/feature-35-github-templates-and-ci`
  Add GitHub issue templates, a PR template, and a source-only repository consistency workflow so the public repo has basic contribution and publication guardrails.
- [x] `pisces/feature-36-repository-word-dto-refactor`
  Add `ModuleRepository`, `ModuleWordInterop`, and `clsLetterHistoryRecord`, move typed history loading away from pipe-delimited strings, route Word lifecycle through explicit acquire/release helpers, and keep `ModuleMain` as a compatibility façade with smoke assertions for the new contracts.

## Deferred Backlog

- [ ] `recipient-title-normalization`
  Add a controlled `RecipientTemplateKey` to `tblAddresses`, let operators choose an approved recipient-title template in `frmLetterCreator`, and build the final addressee line programmatically so users stop inventing inconsistent variants such as `Командиру`, `Начальнику`, or `Для ...` in free text.
