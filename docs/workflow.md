[<- Previous Page](getting-started.md) · [Back to README](../README.md) · [Next Page ->](architecture.md)

# Workflow

## Primary User Flow

1. Launch `frmLetterCreator`.
2. Fill required recipient fields.
3. Select letter type and metadata.
4. Add attachments and optional values.
5. Validate fields and finalize output.

## Form and Validation Behavior

- Required field validation is centralized in `ModuleMain.ValidateRequiredFields`.
- Phone formatting and checks use:
  - `FormatPhoneNumber`
  - `IsPhoneNumberValid`
- Progressive input handling is managed in UserForm event handlers.

## Data Lookups

- Address lookup searches workbook address sheet (`SearchAddresses`).
- Attachment lookup uses workbook settings/attachments sheet (`SearchAttachments`).
- Executor options are loaded from settings (`GetExecutorsList`).

## Output Preparation

- Attachment lines are assembled into letter-ready text.
- Letter metadata and selected values are prepared for template insertion.
- History and audit-related actions are handled by dedicated modules/forms.

## Common Operator Checks

- Ensure required fields are non-empty before generation.
- Confirm selected recipient row is valid.
- Check date/number formatting prior to final output.

## See Also

- [Getting Started](getting-started.md) - Initial setup and run
- [Architecture](architecture.md) - Module responsibilities
- [Maintenance](maintenance.md) - Backup and source-sync operations
