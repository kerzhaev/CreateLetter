[← Previous Page](architecture.md) · [Back to README](../README.md) · [Next Page →](development-workflow.md)

# Configuration

## Workbook Data Configuration

Key sheets are initialized/managed by `mdlInicialize.bas`:

- Address sheet: recipient address fields
- Letters sheet: generated/outgoing letter records
- Settings sheet: static texts, executors, and helper values

## Template Configuration

- `ШаблонПисьма.docx`
- `ШаблонПисьмаДСП.docx`

Keep template names and placement stable unless code references are updated.

## Archive and Backup

- `filesarchive/` stores historical workbook copies.
- `filesarchive/restore-point-*/` stores per-feature rollback snapshots.
- Backup helpers in `ModuleBackup.bas` and `MdlBackup1.bas` support safer changes.

## Source Management

- `CreateLetter.xlsm.modules/` is the authoritative source for VBA text artifacts.
- The project uses a modified `VbaModuleManager` in manual developer mode.
- Automatic workbook lifecycle hooks are intentionally out of scope for module sync.

## Localization Baseline

- `ModuleLocalization.bas` provides the string lookup API for staged localization work.
- The localization worksheet is optional at this stage and is not required for runtime stability.
- Built-in fallback values keep Russian user-facing output stable until string extraction is performed.

## Agent and AI Context

- `.ai-factory/DESCRIPTION.md`: project specification
- `.ai-factory/ARCHITECTURE.md`: architecture constraints
- `AGENTS.md`: structural map
- `.mcp.json`: project MCP server configuration (`github`, `filesystem`)

## Environment Variables (for MCP)

- `GITHUB_TOKEN` (required for GitHub MCP)

## Operational Recommendations

- Keep workbook and `.modules` exports synchronized.
- Create a restore point before every repo-tracked change.
- Validate required sheet headers before major updates.
- Treat UTF-8-readable module exports as a release gate for migration work.
- Preserve template compatibility across releases.

## See Also

- [Getting Started](getting-started.md) - Setup prerequisites and first run
- [Architecture](architecture.md) - Structure and boundaries
- [Development Workflow](development-workflow.md) - Branching, restore points, and manual module sync
- [Maintenance](maintenance.md) - Update and recovery procedures
