# CreateLetter

> Excel VBA tool for composing and tracking outgoing letters from reusable templates.

CreateLetter is a macro-enabled workbook (`CreateLetter.xlsm`) that guides users through letter preparation with structured address data, attachment lists, and template-driven output.
The codebase is source-managed in `CreateLetter.xlsm.modules/` for maintainable VBA workflows.

## Quick Start

```bash
# 1) Open workbook in Excel
start CreateLetter.xlsm

# 2) Enable macros when prompted
# 3) Run the main form from workbook UI/macro entrypoint
```

## Requirements

- Microsoft Excel with macro support
- Access to workbook file `CreateLetter.xlsm`
- Template files available in project root:
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
CreateLetter.xlsm             # Main workbook used in Excel
CreateLetter.xlsm.modules/    # Source-managed VBA modules/forms
filesarchive/                 # Workbook archive snapshots
.ai-factory/                  # Project context and architecture docs
```

## Typical Use Cases

- Preparing official outgoing letters with consistent formatting
- Reusing saved recipient and executor reference data
- Tracking letter preparation history in workbook-managed records
- Maintaining VBA source in module files for safer updates

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

Proprietary/internal unless specified otherwise by the project owner.
