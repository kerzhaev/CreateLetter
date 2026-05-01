# Graph Report - CreateLetter  (2026-05-01)

## Corpus Check
- 21 files · ~33,140 words
- Verdict: corpus is large enough that graph structure adds value.

## Summary
- 92 nodes · 157 edges · 8 communities detected
- Extraction: 98% EXTRACTED · 2% INFERRED · 0% AMBIGUOUS · INFERRED: 3 edges (avg confidence: 0.8)
- Token cost: 0 input · 0 output

## Community Hubs (Navigation)
- [[_COMMUNITY_Community 0|Community 0]]
- [[_COMMUNITY_Community 1|Community 1]]
- [[_COMMUNITY_Community 2|Community 2]]
- [[_COMMUNITY_Community 3|Community 3]]
- [[_COMMUNITY_Community 4|Community 4]]
- [[_COMMUNITY_Community 5|Community 5]]
- [[_COMMUNITY_Community 6|Community 6]]
- [[_COMMUNITY_Community 7|Community 7]]

## God Nodes (most connected - your core abstractions)
1. `main()` - 13 edges
2. `sync_workbook()` - 9 edges
3. `main()` - 7 edges
4. `export_workbook()` - 7 edges
5. `sync_component()` - 7 edges
6. `export_component()` - 6 edges
7. `main()` - 5 edges
8. `remove_invalid_print_area_aliases()` - 5 edges
9. `qn()` - 4 edges
10. `ensure_root_relationship()` - 4 edges

## Surprising Connections (you probably didn't know these)
- `main()` --calls--> `remove_invalid_print_area_aliases()`  [INFERRED]
  scripts\import_envelope_template_layouts.py → scripts\repair_workbook_package.py
- `remove_invalid_print_area_aliases()` --calls--> `sync_workbook()`  [INFERRED]
  scripts\repair_workbook_package.py → starter-kit\excel-vba-com\scripts\sync_vba_from_modules.py
- `main()` --calls--> `remove_invalid_print_area_aliases()`  [INFERRED]
  scripts\ensure_localization_sheet.py → scripts\repair_workbook_package.py
- `export_workbook()` --calls--> `reset_excel_gen_cache()`  [EXTRACTED]
  starter-kit\excel-vba-com\scripts\export_vba_to_modules.py → scripts\export_vba_to_modules.py
- `sync_workbook()` --calls--> `reset_excel_gen_cache()`  [EXTRACTED]
  starter-kit\excel-vba-com\scripts\sync_vba_from_modules.py → scripts\sync_vba_from_modules.py

## Communities

### Community 0 - "Community 0"
Cohesion: 0.31
Nodes (13): build_parser(), ensure_address_group_column(), ensure_envelope_formats_seed(), ensure_layout_sheet(), ensure_print_sheet(), ensure_sheet_headers(), ensure_table(), ensure_table_columns() (+5 more)

### Community 1 - "Community 1"
Cohesion: 0.37
Nodes (12): build_parser(), derive_document_modules_dir(), extract_existing_userform_code(), get_component_by_name(), iter_source_files(), main(), read_source_text(), reset_excel_gen_cache() (+4 more)

### Community 2 - "Community 2"
Cohesion: 0.41
Nodes (11): build_document_module_source(), build_parser(), build_standard_module_source(), derive_document_modules_dir(), export_component(), export_workbook(), get_code_module_text(), main() (+3 more)

### Community 3 - "Community 3"
Cohesion: 0.31
Nodes (9): build_parser(), get_or_create_sheet(), main(), parse_translations(), reset_excel_gen_cache(), write_localization_sheet(), build_parser(), main() (+1 more)

### Community 4 - "Community 4"
Cohesion: 0.47
Nodes (7): build_parser(), ensure_content_type_entry(), ensure_root_relationship(), inject_custom_ui(), main(), next_relationship_id(), qn()

### Community 5 - "Community 5"
Cohesion: 0.28
Nodes (4): Add-TableRow(), Get-ShapeText(), Set-TableRowValues(), Test-EnvelopeSheet()

### Community 6 - "Community 6"
Cohesion: 0.52
Nodes (6): build_parser(), clear_dynamic_content(), clear_shapes(), import_template(), main(), reset_excel_gen_cache()

### Community 7 - "Community 7"
Cohesion: 0.29
Nodes (1): Add-Result()

## Knowledge Gaps
- **Thin community `Community 7`** (7 nodes): `Add-Result()`, `Get-TableColumnNames()`, `Get-WorksheetTableNames()`, `Test-DocumentModuleSourceCoverage()`, `Test-WorksheetVariants()`, `run_excel_smoke_test.ps1`, `run_excel_smoke_test.ps1`
  Too small to be a meaningful cluster - may be noise or needs more connections extracted.

## Suggested Questions
_Questions this graph is uniquely positioned to answer:_

- **Why does `remove_invalid_print_area_aliases()` connect `Community 3` to `Community 1`, `Community 6`?**
  _High betweenness centrality (0.080) - this node is a cross-community bridge._
- **Why does `sync_workbook()` connect `Community 1` to `Community 3`?**
  _High betweenness centrality (0.059) - this node is a cross-community bridge._