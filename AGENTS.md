# AGENTS.md

> Project map for AI agents. Keep this file up-to-date as the project evolves.

## Project Overview
CreateLetter is an Excel VBA workbook used to prepare letters from templates using structured address/settings data and guided user forms.
Source-managed VBA modules are stored in `CreateLetter.xlsm.modules/` and synchronized with the workbook artifact through a manual modified `VbaModuleManager` workflow.

## Tech Stack
- **Language:** VBA
- **Framework:** Excel UserForms + Excel Object Model
- **Database:** Excel worksheets (inside `CreateLetter.xlsm`)
- **ORM:** N/A

## Project Structure
```text
.
├── CreateLetter.xlsm                    # Main macro-enabled workbook
├── ШаблонПисьма.docx                    # Letter template
├── ШаблонПисьмаДСП.docx                 # Alternate letter template
├── CreateLetter.xlsm.modules/           # Exported VBA source modules/forms
│   ├── frmLetterCreator.frm/.frx        # Main wizard UI and resources
│   ├── frmLetterHistory.frm/.frx        # History UI and resources
│   ├── ModuleMain.bas                   # Core logic and validation
│   ├── mdlInicialize.bas                # Worksheet bootstrap/reset
│   ├── ModuleDates.bas                  # Date helpers
│   ├── ModuleLocalization.bas           # Localization lookup foundation
│   ├── ModuleCache.bas                  # Cache helpers
│   ├── ModuleBackup.bas                 # Backup helpers
│   ├── ModuleAuditLogger.bas            # Audit/logging helpers
│   └── MdlBackup1.bas                   # Legacy backup logic
├── filesarchive/                        # Archived workbook versions
│   └── restore-point-*/                 # Local rollback snapshots before feature work
├── scripts/
│   └── create_restore_point.ps1         # Creates workbook + modules restore points
│   └── run_excel_smoke_test.ps1         # Excel COM smoke-test helper
│   └── sync_vba_from_modules.py         # Excel COM VBA sync helper for modules/forms
├── .agents/skills/                      # External installed skills for agent use
│   ├── xlsx/
│   └── vbaexcel/
├── .ai-factory/                         # AI Factory context artifacts
│   ├── DESCRIPTION.md
│   ├── ARCHITECTURE.md
│   └── RULES.md
└── .codex/skills/                       # Built-in project-local AIF skills
```

## Key Entry Points
| File | Purpose |
|------|---------|
| CreateLetter.xlsm | Runtime workbook used by end users |
| CreateLetter.xlsm.modules/ModuleMain.bas | Core business rules and shared logic |
| CreateLetter.xlsm.modules/frmLetterCreator.frm | Main letter creation flow |
| CreateLetter.xlsm.modules/mdlInicialize.bas | Initializes/reset required worksheets |
| .ai-factory.json | AI Factory agent setup metadata |

## Documentation
| Document | Path | Description |
|----------|------|-------------|
| AGENTS map | AGENTS.md | This project map |
| Project description | .ai-factory/DESCRIPTION.md | Stack, scope, and setup context |
| Architecture guide | .ai-factory/ARCHITECTURE.md | Architecture pattern and dependency rules |
| Project rules | .ai-factory/RULES.md | AI Factory execution and migration rules |
| Development workflow | docs/development-workflow.md | Branching, restore points, manual module sync |

## AI Context Files
| File | Purpose |
|------|---------|
| AGENTS.md | This file - project structure map |
| .ai-factory/DESCRIPTION.md | Project specification and tech stack |
| .ai-factory/ARCHITECTURE.md | Architecture decisions and guidelines |
| .ai-factory/RULES.md | AI Factory delivery rules and migration constraints |
| .ai-factory/ROADMAP.md | High-level staged migration roadmap |
| .cursorrules | Existing repository agent/convention guidance |
