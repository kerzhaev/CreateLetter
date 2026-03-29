# Project: CreateLetter

## Overview
CreateLetter is an Excel VBA automation project for preparing outgoing letters using predefined DOCX templates and workbook-managed reference data.
It provides a guided user form workflow for address selection, attachment composition, validation, and record keeping.

## Core Features
- Multi-step letter creation UI (`frmLetterCreator`) with required field validation.
- Address and executor lookup from workbook sheets.
- Attachment list composition and formatting for generated letter content.
- Workbook initialization/reset helpers for required sheets and starter data.
- Backup and audit-related VBA modules for operational safety and traceability.

## Tech Stack
- **Language:** VBA (Visual Basic for Applications)
- **Framework:** Excel UserForms + Workbook/Object Model
- **Database:** Excel worksheets (`Адреса`, `Письма`, `Настройки`)
- **ORM:** N/A
- **Integrations:** Local DOCX templates and file-system based archives
- **Developer Tooling:** AI Factory pipeline + manual modified `VbaModuleManager`

## Architecture Notes
Project uses a modular-monolith style inside a single Excel workbook:
- UI forms orchestrate user interaction.
- Core business logic and validation live in standard modules.
- Data persistence is worksheet-based.
- Utility modules handle cross-cutting concerns (cache, backup, audit, dates).

## Non-Functional Requirements
- Logging: Debug logging and audit-friendly operation paths.
- Error handling: Guard clauses, user-visible messages, and safe fallbacks.
- Security: Input validation, controlled file operations, and backup-before-mutation practices.
- Reliability: Keep workbook formulas/data stable and preserve user templates.

## Migration Baseline
- The workbook runtime must remain independent from source-management hooks.
- `CreateLetter.xlsm.modules/` is the source of truth for standard VBA modules, class modules, and forms.
- `CreateLetter.xlsm.document-modules/` is the source of truth for workbook and worksheet document modules.
- Module sync is manual and must not be attached to workbook open/save/close events.
- UTF-8-readable text artifacts are the target baseline for Git and AI-agent workflows.
- Internal identifiers are migrating toward English ASCII; user-facing Russian text will move to localization data in later stages.
- A localization foundation module now exists without requiring workbook schema changes; workbook-backed localization data remains optional at this stage.

## Architecture
See `.ai-factory/ARCHITECTURE.md` for detailed architecture guidelines.
Pattern: Modular Monolith
