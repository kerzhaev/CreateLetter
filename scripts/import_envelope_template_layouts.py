#!/usr/bin/env python
"""
Import C4/C5/DL envelope print geometry from the reference workbook.

The script copies only worksheet layout primitives into the existing hidden
CreateLetter layout sheets: cell formats, merge geometry, row heights, column
widths, and page setup. It removes old formulas, buttons, list boxes, and other
Form Controls from the imported template surface.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pythoncom
import win32com.client.gencache
from win32com.client.dynamic import Dispatch

from repair_workbook_package import remove_invalid_print_area_aliases


TEMPLATE_SHEETS = {
    "c4": {
        "source": "C4",
        "target": "DispatchLayout_C4",
        "range": "A1:K25",
        "paper": 9,
        "margins": (1.0, 1.0, 1.0, 1.0),
    },
    "c5": {
        "source": "C5",
        "target": "DispatchLayout_C5",
        "range": "A1:K22",
        "paper": 28,
        "margins": (1.5, 1.5, 1.5, 1.2),
    },
    "dl": {
        "source": "DL",
        "target": "DispatchLayout_DL",
        "range": "A1:K18",
        "paper": 27,
        "margins": (0.5, 1.0, 0.5, 0.5),
    },
}


DYNAMIC_CELLS = {
    "c4": ("C2", "C3", "C4", "C5", "E7", "H18", "H19", "H20", "H21", "H22", "H24", "A23"),
    "c5": ("C2", "C3", "C4", "C5", "E7", "H15", "H16", "H17", "H18", "H19", "H21", "A20"),
    "dl": ("C2", "C3", "C4", "C5", "E7", "H11", "H12", "H13", "H14", "H15", "H17", "A16"),
}


def reset_excel_gen_cache() -> None:
    gen_path = Path(win32com.client.gencache.GetGeneratePath())
    for child in gen_path.glob("00020813-0000-0000-C000-000000000046*"):
        if child.is_dir():
            import shutil
            shutil.rmtree(child, ignore_errors=True)
        elif child.exists():
            child.unlink(missing_ok=True)


def clear_shapes(ws) -> None:
    for shape_index in range(ws.Shapes.Count, 0, -1):
        ws.Shapes.Item(shape_index).Delete()


def clear_dynamic_content(ws, format_key: str) -> None:
    for address in DYNAMIC_CELLS[format_key]:
        cell = ws.Range(address)
        if bool(cell.MergeCells):
            cell.MergeArea.ClearContents()
        else:
            cell.ClearContents()


def import_template(source_wb, target_wb, format_key: str) -> None:
    settings = TEMPLATE_SHEETS[format_key]
    source_ws = source_wb.Worksheets(settings["source"])
    target_ws = target_wb.Worksheets(settings["target"])
    source_range = source_ws.Range(settings["range"])
    target_range = target_ws.Range(settings["range"])

    target_ws.Visible = -1
    clear_shapes(target_ws)
    target_ws.Cells.UnMerge()
    target_ws.Cells.Clear()
    target_ws.ResetAllPageBreaks()

    source_range.Copy()
    target_ws.Range("A1").PasteSpecial(-4104)
    target_ws.Range("A1").PasteSpecial(-4122)
    target_wb.Application.CutCopyMode = False

    for col_index in range(1, source_range.Columns.Count + 1):
        target_ws.Columns(col_index).ColumnWidth = source_ws.Columns(col_index).ColumnWidth

    for row_index in range(1, source_range.Rows.Count + 1):
        target_ws.Rows(row_index).RowHeight = source_ws.Rows(row_index).RowHeight

    clear_shapes(target_ws)

    for row_index in range(1, source_range.Rows.Count + 1):
        for col_index in range(1, source_range.Columns.Count + 1):
            cell = target_ws.Cells(row_index, col_index)
            formula = cell.Formula
            if isinstance(formula, str) and formula.startswith("="):
                if bool(cell.MergeCells):
                    if cell.Address == cell.MergeArea.Cells(1, 1).Address:
                        cell.MergeArea.ClearContents()
                else:
                    cell.ClearContents()

    clear_dynamic_content(target_ws, format_key)

    last_cell = source_range.Cells(source_range.Rows.Count, source_range.Columns.Count)
    target_ws.PageSetup.PrintArea = target_ws.Range("A1", target_ws.Cells(last_cell.Row, last_cell.Column)).Address

    left_margin, right_margin, top_margin, bottom_margin = settings["margins"]
    page_setup = target_ws.PageSetup
    page_setup.Orientation = 2
    page_setup.PaperSize = settings["paper"]
    page_setup.Zoom = False
    page_setup.FitToPagesWide = 1
    page_setup.FitToPagesTall = 1
    page_setup.LeftMargin = target_wb.Application.CentimetersToPoints(left_margin)
    page_setup.RightMargin = target_wb.Application.CentimetersToPoints(right_margin)
    page_setup.TopMargin = target_wb.Application.CentimetersToPoints(top_margin)
    page_setup.BottomMargin = target_wb.Application.CentimetersToPoints(bottom_margin)
    page_setup.CenterHorizontally = True
    page_setup.CenterVertically = True

    target_ws.Visible = 2


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Import envelope template layouts from the reference workbook.")
    parser.add_argument("target_workbook", type=Path, help="Path to CreateLetter.xlsm")
    parser.add_argument("source_workbook", type=Path, help="Path to the reference workbook with C4/C5/DL sheets")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    target_path = args.target_workbook.resolve()
    source_path = args.source_workbook.resolve()

    if not target_path.exists():
        print(f"Target workbook not found: {target_path}", file=sys.stderr)
        return 1
    if not source_path.exists():
        print(f"Source workbook not found: {source_path}", file=sys.stderr)
        return 1

    remove_invalid_print_area_aliases(target_path)

    pythoncom.CoInitialize()
    reset_excel_gen_cache()
    excel = Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.AutomationSecurity = 3

    source_wb = None
    target_wb = None
    saved = False

    try:
        source_wb = excel.Workbooks.Open(str(source_path), UpdateLinks=0, ReadOnly=True, IgnoreReadOnlyRecommended=True)
        target_wb = excel.Workbooks.Open(str(target_path), UpdateLinks=0, ReadOnly=False, IgnoreReadOnlyRecommended=True)

        for format_key in ("c4", "c5", "dl"):
            import_template(source_wb, target_wb, format_key)
            print(f"imported={format_key}")

        target_wb.Save()
        saved = True
        return 0
    except Exception as exc:
        print(f"ENVELOPE TEMPLATE IMPORT ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        if source_wb is not None:
            source_wb.Close(SaveChanges=False)
        if target_wb is not None:
            target_wb.Close(SaveChanges=saved)
        excel.Quit()
        pythoncom.CoUninitialize()
        remove_invalid_print_area_aliases(target_path)


if __name__ == "__main__":
    raise SystemExit(main())
