# Excel VBA COM Playbook

## Purpose

This document formalizes the Excel COM automation workflow used in CreateLetter so the same pattern can be reused in future VBA workbook projects without rebuilding the process from scratch.

The goal is simple:

- keep VBA text artifacts reviewable in git;
- synchronize workbook code automatically where possible;
- verify workbook behavior through repeatable smoke tests;
- isolate known VBA/VBE edge cases so they do not get rediscovered on every new project.

## Recommended Baseline Layout

```text
.
├── Workbook.xlsm
├── Workbook.xlsm.modules/
│   ├── ModuleMain.bas
│   ├── UserFormA.frm
│   ├── UserFormA.frx
│   └── clsSomething.cls
├── scripts/
│   ├── sync_vba_from_modules.py
│   ├── run_excel_smoke_test.ps1
│   ├── create_restore_point.ps1
│   ├── ensure_workbook_tables.py
│   └── ensure_localization_sheet.py
├── customUI/
│   └── customUI.xml
└── filesarchive/
    └── restore-point-*/
```

## Core Automation Pattern

### 1. Treat exported text as the source of truth

- Keep `.bas`, `.frm`, `.cls` under version control.
- Treat the workbook binary as a runtime artifact, not the primary review surface.
- Export validated workbook changes back into the source folder if a manual VBE change was required.

### 2. Create restore points before tracked edits

Use a dedicated script that snapshots both:

- the workbook binary;
- the exported source folder.

This makes rollback cheap and removes fear from iterative refactors.

### 3. Sync into the workbook through COM before asking for manual help

The preferred flow is:

1. update source files;
2. run COM sync;
3. run smoke tests;
4. ask for manual Excel inspection only if automation hits a real blocker.

### 4. Bootstrap workbook structure by script

If the workbook depends on:

- worksheets,
- `ListObjects`,
- localization sheets,
- Ribbon XML,

materialize them through scripts, not through undocumented one-off manual edits.

### 5. Verify contracts, not only workbook opening

A useful smoke harness must check more than “Excel opened”.

It should verify:

- workbook opens;
- expected sheets exist;
- expected `ListObjects` exist;
- key VBA public entry points exist;
- key localization lookups work;
- schema/bootstrap artifacts are present;
- customUI/Ribbon package parts are embedded when required;
- chosen architecture contracts still hold after refactors.

## Minimal Script Set

### `sync_vba_from_modules.py`

Responsibilities:

- open workbook via COM;
- import/update standard modules, class modules, and forms;
- preserve true class modules as classes;
- strip export-only metadata before VBE text insertion;
- tolerate legacy encodings during migration.

Must handle:

- `Attribute VB_*` removal before `CodeModule.AddFromString`;
- `.cls` as real class components, not standard modules;
- legacy ANSI/CP1251 exported files when the project is not fully UTF-8-clean yet.

### `run_excel_smoke_test.ps1`

Responsibilities:

- open workbook in read-only mode when possible;
- inspect workbook sheets/tables/functions through Excel COM;
- verify package artifacts like Ribbon XML through a temporary workbook copy;
- return machine-readable PASS/FAIL-style results.

Recommended checks:

- workbook open;
- required worksheets;
- required tables;
- formatting helpers;
- validation helpers;
- localization module availability;
- architecture/refactor contracts;
- Ribbon customization presence.

### `create_restore_point.ps1`

Responsibilities:

- create timestamped rollback folders;
- copy workbook and source modules together;
- label the snapshot by feature or stage.

### Schema/bootstrap scripts

Examples:

- `ensure_workbook_tables.py`
- `ensure_localization_sheet.py`
- `apply_custom_ui.py`

These should be idempotent and safe to rerun.

## Smoke Test Design Rules

### Test contracts, not implementation trivia

Good checks:

- “table exists”;
- “public alias resolves”;
- “localization lookup returns non-fallback value”;
- “history flow no longer depends on pipe-delimited parsing”.

Avoid brittle checks that depend on incidental formatting or editor state.

### Inspect package parts from a temporary workbook copy

Excel often keeps the workbook file locked while COM holds it open.

For ZIP/package inspection such as:

- `customUI/customUI.xml`,
- `_rels/.rels`,
- content type overrides,

first copy the workbook to a temp path, then inspect the copy.

This avoids false negatives caused by file locks.

### Separate optional gates with explicit flags

Use flags like:

- `-RequireLocalizationModule`
- `-RequireStructuredTables`
- `-RequireLocalizationSheet`
- `-RequireRibbonCustomization`

This lets the smoke harness evolve safely as the workbook baseline grows.

## Known VBA/VBE Edge Cases

### `Attribute VB_*` lines

These are valid in exported source files but invalid when pasted into the VBE code pane through automation.

Rule:

- strip all `Attribute VB_*` lines before text injection;
- keep them only in exported source artifacts.

### Class modules are special

A `.cls` file cannot be treated like a standard module.

Rule:

- create or reuse a real class component in `VBProject`;
- inject only the class body beginning with `Option Explicit`.

If automation fails, manual fallback is:

1. `Insert -> Class Module`;
2. set `(Name)`;
3. paste class body only;
4. export back to the source folder.

### Duplicate procedure names across legacy modules

VBA will happily compile until ambiguous public names collide.

Rule:

- keep legacy compatibility wrappers thin and explicitly qualified;
- rename legacy duplicates instead of only changing the newest caller.

### Encoding drift

UserForms and old exports can still carry legacy codepages.

Rule:

- keep internal identifiers ASCII-safe;
- keep user-facing text in localization or workbook-backed data when possible;
- if a module repeatedly corrupts Cyrillic on import, use English fallback literals in the code and keep visible UI text in workbook/localization assets.

## Reusable Delivery Cycle

For each feature:

1. create a feature branch;
2. create a restore point;
3. edit source-managed modules;
4. sync to workbook through COM;
5. run smoke tests;
6. inspect workbook manually only if automation cannot fully verify the change;
7. export back only if a manual VBE change was required;
8. commit the source changes;
9. keep the smoke harness updated if the workbook baseline changed.

## What To Reuse On The Next Project

At minimum, copy the following ideas:

- source-managed `Workbook.xlsm.modules/`;
- COM sync helper;
- restore point script;
- smoke harness with explicit feature flags;
- schema/bootstrap scripts;
- rule that manual Excel checks are a fallback, not the first step.

At minimum, re-check these known blockers early:

- VBProject access permissions;
- class-module import path;
- form encoding/export behavior;
- workbook file locking during ZIP/package inspection;
- any customUI embedding step.

## Recommended AI Factory Rules For Excel VBA Projects

- Source-managed modules are the review surface.
- Workbook schema changes must be scripted and idempotent.
- Smoke tests must evolve whenever the workbook baseline evolves.
- COM sync should be attempted before asking for manual import.
- Manual VBE edits must be re-exported back to source immediately.
- Strip export-only metadata before VBE insertion.
- Package-level checks should read a temporary workbook copy.

## Final Principle

The main win is not “automation for its own sake”.

The real win is that workbook development becomes:

- reversible,
- reviewable,
- scriptable,
- and repeatable across projects.

That is the part worth copying to the next VBA repository.
