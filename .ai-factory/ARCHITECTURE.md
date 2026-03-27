# Architecture: Modular Monolith

## Overview
This project is best modeled as a modular monolith: one deployable workbook (`CreateLetter.xlsm`) with clear internal module boundaries. It fits the current scope, keeps operational complexity low, and preserves fast iteration inside Excel/VBA constraints.

The architecture separates UI orchestration, core business logic, and worksheet-backed storage conventions. This improves maintainability while staying aligned with how Excel VBA applications are built and shipped.

The migration baseline adds two operational constraints: source-managed VBA remains manual in the developer workflow, and runtime workbook behavior must not depend on workbook event hooks for module synchronization.

## Decision Rationale
- **Project type:** Excel VBA letter automation workbook
- **Tech stack:** VBA + Excel Object Model + UserForms
- **Key factor:** Single-runtime Excel solution with moderate domain complexity and strong need for predictable module boundaries

## Folder Structure
```text
.
├── CreateLetter.xlsm                    # Runtime workbook artifact
├── CreateLetter.xlsm.modules/           # Source-managed VBA exports
│   ├── frmLetterCreator.frm             # Main letter creation UI workflow
│   ├── frmLetterCreator.frx             # Form resources (binary)
│   ├── frmLetterHistory.frm             # Letter history UI
│   ├── frmLetterHistory.frx             # Form resources (binary)
│   ├── ModuleMain.bas                   # Core business logic and validation
│   ├── mdlInicialize.bas                # Workbook sheet/bootstrap initialization
│   ├── ModuleDates.bas                  # Date formatting/normalization helpers
│   ├── ModuleLocalization.bas           # Localization lookup foundation
│   ├── ModuleCache.bas                  # In-memory/state caching helpers
│   ├── ModuleBackup.bas                 # Backup and recovery helpers
│   ├── ModuleAuditLogger.bas            # Auditing/logging concerns
│   └── MdlBackup1.bas                   # Legacy backup module
├── filesarchive/                        # Archived workbook snapshots
├── scripts/                             # Safe developer helper scripts
│   └── create_restore_point.ps1         # Workbook + module rollback snapshot
├── .ai-factory/
│   ├── DESCRIPTION.md                   # Project spec and setup context
│   ├── ARCHITECTURE.md                  # This architecture guide
│   └── RULES.md                         # AI Factory execution rules
└── AGENTS.md                            # AI-facing project map
```

## Dependency Rules
- `frm*` UserForms may call `ModuleMain`, `ModuleDates`, `ModuleCache`.
- `frm*` UserForms may call `ModuleLocalization` for user-facing strings.
- Core logic modules must not depend on UserForm instance state.
- Initialization/bootstrap (`mdlInicialize`) may prepare worksheets but must not include business formatting logic from UI.
- Backup/audit modules are cross-cutting and should expose narrow procedures used by UI/core modules.
- Source-management tooling must stay outside workbook runtime flow; manual module sync must not become a hidden runtime dependency.

- ✅ Allowed: `frmLetterCreator -> ModuleMain -> Worksheet data`
- ✅ Allowed: `frmLetterCreator -> ModuleAuditLogger`
- ❌ Forbidden: `ModuleMain -> frmLetterCreator controls`
- ❌ Forbidden: cyclic calls across utility modules without a clear owning module

## Layer/Module Communication
- UI-to-core: UserForm events call explicit public procedures/functions.
- Core-to-data: Core modules read/write worksheets through validated access patterns.
- Cross-cutting concerns: Core/UI modules call backup and audit helpers through stable entry points.

## Key Principles
1. Keep domain rules in standard modules, not in form event handlers.
2. Treat worksheets as persistence boundaries with strict validation on read/write.
3. Keep side effects explicit: backup/audit actions should be easy to trace.
4. Treat `CreateLetter.xlsm.modules/` as the authoritative text source for VBA.
5. Execute migration in branch-sized feature stages with a restore point before repo-tracked edits.
6. Favor English ASCII for new internal identifiers and localization for user-facing Russian text.
7. Add localization in two steps: foundation first, user-facing string migration second.
8. Target thin UserForms: forms should orchestrate control state and delegate business logic to shared modules.
9. Prefer extending existing standard modules before creating new ones; during the form-refactor stage, cap new modules at 1-2 unless a concrete boundary requires more.
10. When moving worksheet persistence toward `ListObjects`, keep a compatibility fallback to raw ranges until the workbook schema itself is upgraded and verified.

## Code Examples

### UI Event Delegates to Core Logic
```vb
' In frmLetterCreator.frm
Private Sub btnValidate_Click()
    Dim errText As String
    errText = ValidateRequiredFields( _
        Me.txtAddressee.Value, _
        Me.txtCity.Value, _
        Me.txtRegion.Value, _
        Me.txtPostalCode.Value, _
        Me.cmbExecutor.Value)

    If Len(errText) > 0 Then
        MsgBox errText, vbExclamation
        Exit Sub
    End If
End Sub
```

### Core Logic Stays UI-Agnostic
```vb
' In ModuleMain.bas
Public Function NormalizeExecutorPhone(rawPhone As String) As String
    NormalizeExecutorPhone = FormatPhoneNumber(rawPhone)
End Function
```

## Anti-Patterns
- ❌ Embedding non-trivial business rules directly inside UserForm event handlers.
- ❌ Writing to worksheets without validation and error handling.
- ❌ Mixing backup, audit, UI updates, and domain rules in one procedure.
- ❌ Attaching module import/export to workbook open/save/close events in this project.
- ❌ Combining source-management migration, localization migration, and schema migration in one unverified change set.
