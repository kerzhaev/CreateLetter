#!/usr/bin/env python
"""
Repair package-level workbook issues that block Excel COM automation.

Excel reserves `_xlnm.Print_Area` for worksheet print areas. Some workbook
saves can leave an invalid `Print_Area` alias beside the reserved name, and
Excel then opens a modal "name conflict" dialog during unattended COM runs.
"""

from __future__ import annotations

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path


INVALID_PRINT_AREA_PATTERN = re.compile(
    r'<definedName\s+name="Print_Area"[^>]*>.*?</definedName>',
    re.IGNORECASE,
)


def remove_invalid_print_area_aliases(workbook_path: Path) -> int:
    workbook_path = workbook_path.resolve()
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")

    with zipfile.ZipFile(workbook_path, "r") as source_zip:
        entries = {name: source_zip.read(name) for name in source_zip.namelist()}

    workbook_xml_name = "xl/workbook.xml"
    if workbook_xml_name not in entries:
        return 0

    workbook_xml = entries[workbook_xml_name].decode("utf-8")
    repaired_xml, removed_count = INVALID_PRINT_AREA_PATTERN.subn("", workbook_xml)
    if removed_count == 0:
        return 0

    entries[workbook_xml_name] = repaired_xml.encode("utf-8")

    temp_path = workbook_path.with_name(f"{workbook_path.stem}.repair.tmp{workbook_path.suffix}")

    try:
        with zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as target_zip:
            for entry_name, entry_data in entries.items():
                target_zip.writestr(entry_name, entry_data)
        shutil.move(temp_path, workbook_path)
    finally:
        temp_path.unlink(missing_ok=True)

    return removed_count


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Repair invalid workbook package names that block Excel COM automation.")
    parser.add_argument("workbook", type=Path, help="Path to the workbook package")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        removed_count = remove_invalid_print_area_aliases(args.workbook)
    except Exception as exc:
        print(f"WORKBOOK REPAIR ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"invalid_print_area_aliases_removed={removed_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
