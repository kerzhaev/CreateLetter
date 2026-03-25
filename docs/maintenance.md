[← Previous Page](development-workflow.md) · [Back to README](../README.md)

# Maintenance

## Safe Update Workflow

1. Create a restore point with `.\scripts\create_restore_point.ps1 -Label <feature-name>`.
2. Import VBA modules manually through the modified `VbaModuleManager` workflow.
3. Apply scoped changes in the current `pisces/<feature-name>` branch only.
4. Export modules back to `CreateLetter.xlsm.modules/`.
5. Validate the smoke-test user flows in Excel.
6. Run the verification step before merging to `main`.

## Backup Strategy

- Keep periodic snapshots in `filesarchive/`.
- Before every feature-stage edit, create a dated restore point with both workbook and module snapshots.

## VBA Export/Import Tooling

- Use the modified manual `VbaModuleManager` flow as the primary sync mechanism.
- Keep `CreateLetter.xlsm.modules/` aligned with the workbook before and after each tested change.
- Use helper tooling only when it does not replace the manual sync contract defined for this project.

## Regression Checklist

- Form initialization works (`frmLetterCreator`).
- Required fields validation still blocks invalid submissions.
- Address/attachment search returns expected rows.
- Date and phone formatting output remains correct.
- No runtime macro errors on normal workflow.

## Documentation Hygiene

When structure/behavior changes:
- Update `AGENTS.md` for layout changes.
- Update `.ai-factory/DESCRIPTION.md` for scope/stack changes.
- Update `.ai-factory/ARCHITECTURE.md` for boundary rule changes.
- Update `.ai-factory/RULES.md` when workflow rules change.

## See Also

- [Configuration](configuration.md) - Runtime and data configuration
- [Development Workflow](development-workflow.md) - AI Factory pipeline and restore-point policy
- [Workflow](workflow.md) - Functional usage flow
- [Architecture](architecture.md) - Module-level design rules
