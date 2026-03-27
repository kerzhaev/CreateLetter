# CreateLetter

> Excel VBA tool for composing and tracking outgoing letters from reusable templates.

CreateLetter is a source-managed Excel VBA project for preparing letters from workbook-backed data and reusable templates.
The public repository focuses on the VBA source in `CreateLetter.xlsm.modules/` plus the project tooling/documentation. Runtime workbook and template binaries are intentionally kept out of the public tree.

## Quick Start

```bash
# 1) Provision local runtime assets (workbook + templates)
# 2) Sync VBA sources into the local workbook
python .\scripts\sync_vba_from_modules.py .\CreateLetter.xlsm .\CreateLetter.xlsm.modules

# 3) Open workbook in Excel and enable macros
```

## Requirements

- Microsoft Excel with macro support
- A local runtime workbook `CreateLetter.xlsm` (not stored in the public repository)
- Local template files:
  - `LetterTemplate.docx`
  - `LetterTemplateFOU.docx`

## Key Features

- **Guided form workflow** for letter creation (`frmLetterCreator`)
- **Address and executor lookup** from workbook sheets
- **Attachment composition** with formatted output text
- **Workbook initialization/reset** for required sheets
- **Backup and audit helpers** for safer operations

## Example

```text
1. Select recipient address from workbook data.
2. Fill letter metadata (date, number, executor).
3. Add one or more attachment items.
4. Validate and generate final letter content for template usage.
```

## Project Layout

```text
CreateLetter.xlsm.modules/    # Source-managed VBA modules/forms
scripts/                      # Local automation helpers
.ai-factory/                  # Project context and architecture docs
```

## Typical Use Cases

- Preparing official outgoing letters with consistent formatting
- Reusing saved recipient and executor reference data
- Tracking letter preparation history in workbook-managed records
- Maintaining VBA source in module files for safer updates

## Public Repository Scope

- Included:
  - VBA source modules and UserForm code
  - automation scripts
  - documentation and AI Factory context
- Excluded:
  - local workbook/runtime binaries
  - local Word templates
  - `.frx` form resource binaries
  - local AI skill directories and restore-point artifacts

## Maintenance At A Glance

- Keep workbook and exported modules synchronized.
- Back up before structural VBA changes.
- Validate main form flow after any logic update.
- Use documentation pages below for detailed operational guidance.

---

## Documentation

| Guide | Description |
|-------|-------------|
| [Getting Started](docs/getting-started.md) | Prerequisites, setup, first run |
| [Workflow](docs/workflow.md) | End-user flow and field behavior |
| [Architecture](docs/architecture.md) | Module boundaries and dependency rules |
| [Configuration](docs/configuration.md) | Worksheets, templates, and MCP/agent context |
| [Development Workflow](docs/development-workflow.md) | AI Factory pipeline, branching, restore points, and manual module sync |
| [Maintenance](docs/maintenance.md) | VBA export/import, backup, and safe updates |
| [GitHub Publishing](docs/github-publishing.md) | Public-repo hygiene and release checklist |

## AI Context

- [Project Description](.ai-factory/DESCRIPTION.md)
- [Architecture Guidelines](.ai-factory/ARCHITECTURE.md)
- [Agent Project Map](AGENTS.md)

## License

MIT. See [LICENSE](LICENSE).
