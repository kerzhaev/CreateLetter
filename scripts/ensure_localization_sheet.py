#!/usr/bin/env python
"""
Create or refresh the workbook-backed Localization sheet from ModuleLocalization.bas.

The script parses built-in AddTranslation entries and writes them into the
Localization worksheet as:
    key | ru | en
"""

from __future__ import annotations

import argparse
import re
import sys
from collections import defaultdict
from pathlib import Path

import pythoncom
import win32com.client


LOCALIZATION_SHEET_NAME = "Localization"
HEADER_ROW = ("key", "ru", "en")
TRANSLATION_PATTERN = re.compile(
    r'AddTranslation\s+"(?P<lang>[^"]+)",\s+"(?P<key>[^"]+)",\s+"(?P<value>[^"]*)"',
    re.IGNORECASE,
)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create or refresh workbook Localization sheet.")
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument(
        "--module-path",
        type=Path,
        default=Path("CreateLetter.xlsm.modules/ModuleLocalization.bas"),
        help="Path to ModuleLocalization.bas",
    )
    return parser


def parse_translations(module_path: Path) -> dict[str, dict[str, str]]:
    if not module_path.exists():
        raise FileNotFoundError(f"Localization module not found: {module_path}")

    translations: dict[str, dict[str, str]] = defaultdict(dict)
    content = module_path.read_text(encoding="utf-8")
    for match in TRANSLATION_PATTERN.finditer(content):
        lang = match.group("lang").strip().lower()
        key = match.group("key").strip().lower()
        value = match.group("value")
        translations[key][lang] = value
    return dict(translations)


def get_or_create_sheet(workbook):
    for index in range(1, workbook.Worksheets.Count + 1):
        ws = workbook.Worksheets(index)
        if ws.Name == LOCALIZATION_SHEET_NAME:
            return ws, False

    ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    ws.Name = LOCALIZATION_SHEET_NAME
    return ws, True


def write_localization_sheet(ws, translations: dict[str, dict[str, str]]) -> None:
    ws.Cells.Clear()

    for col_index, header in enumerate(HEADER_ROW, start=1):
        ws.Cells(1, col_index).Value = header

    sorted_keys = sorted(translations.keys())
    for row_index, key in enumerate(sorted_keys, start=2):
        ws.Cells(row_index, 1).Value = key
        ws.Cells(row_index, 2).Value = translations[key].get("ru", "")
        ws.Cells(row_index, 3).Value = translations[key].get("en", "")

    ws.Columns("A:C").AutoFit()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    module_path = args.module_path.resolve()

    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    try:
        translations = parse_translations(module_path)
    except Exception as exc:  # noqa: BLE001 - developer tooling script
        print(f"LOCALIZATION PARSE ERROR: {exc}", file=sys.stderr)
        return 1

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    workbook = None

    try:
        workbook = excel.Workbooks.Open(str(workbook_path))
        ws, created = get_or_create_sheet(workbook)
        write_localization_sheet(ws, translations)
        workbook.Save()
        print(f"sheet={'created' if created else 'updated'} rows={len(translations)}")
        return 0
    except Exception as exc:  # noqa: BLE001 - developer tooling script
        print(f"LOCALIZATION SHEET ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
