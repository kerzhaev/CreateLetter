#!/usr/bin/env python
"""
Ensure structured tables exist in the CreateLetter workbook.

This script creates workbook ListObjects for the Addresses and Letters sheets
without changing the existing data layout. It is safe to run multiple times.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pythoncom
import win32com.client


TABLE_SPECS = (
    ("Addresses", "tblAddresses"),
    ("Letters", "tblLetters"),
)

XL_SRC_RANGE = 1
XL_YES = 1


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create missing ListObjects in the workbook.")
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument("--visible", action="store_true", help="Show Excel during the operation")
    return parser


def ensure_table(ws, table_name: str) -> str:
    for index in range(1, ws.ListObjects.Count + 1):
        if ws.ListObjects(index).Name == table_name:
            return "existing"

    used_range = ws.UsedRange
    row_count = used_range.Rows.Count
    col_count = used_range.Columns.Count
    if row_count < 2 or col_count < 1:
        return "skipped-empty"

    first_row = used_range.Row
    first_col = used_range.Column
    last_row = first_row + row_count - 1
    last_col = first_col + col_count - 1

    source_range = ws.Range(ws.Cells(first_row, first_col), ws.Cells(last_row, last_col))
    list_object = ws.ListObjects.Add(XL_SRC_RANGE, source_range, None, XL_YES)
    list_object.Name = table_name
    return "created"


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = args.visible
    excel.DisplayAlerts = False
    workbook = None

    try:
        workbook = excel.Workbooks.Open(str(workbook_path))
        for sheet_name, table_name in TABLE_SPECS:
            ws = workbook.Worksheets(sheet_name)
            status = ensure_table(ws, table_name)
            print(f"{sheet_name}:{table_name}:{status}")

        workbook.Save()
        return 0
    except Exception as exc:  # noqa: BLE001 - developer tooling script
        print(f"TABLE BOOTSTRAP ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
