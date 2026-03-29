#!/usr/bin/env python
"""
Inject source-managed Ribbon XML into an XLSM workbook package.

This is a generic helper for workbooks that store Ribbon markup in:

    customUI/customUI.xml
"""

from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path


CUSTOM_UI_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
CUSTOM_UI_CONTENT_TYPE = "application/vnd.ms-office.customUI+xml"
CUSTOM_UI_PACKAGE_PATH = "customUI/customUI.xml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Apply Ribbon customUI XML to an XLSM workbook.")
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument(
        "--xml",
        type=Path,
        default=Path("customUI/customUI.xml"),
        help="Path to customUI XML (default: customUI/customUI.xml)",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    workbook = args.workbook.resolve()
    xml_path = args.xml.resolve()

    if not workbook.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook}")
    if not xml_path.exists():
        raise FileNotFoundError(f"customUI XML not found: {xml_path}")

    with tempfile.TemporaryDirectory() as temp_dir:
        temp_workbook = Path(temp_dir) / workbook.name
        shutil.copy2(workbook, temp_workbook)

        with zipfile.ZipFile(temp_workbook, "a") as archive:
            archive.writestr(CUSTOM_UI_PACKAGE_PATH, xml_path.read_text(encoding="utf-8"))

            content_types_path = "[Content_Types].xml"
            content_types = archive.read(content_types_path).decode("utf-8")
            if CUSTOM_UI_PACKAGE_PATH not in content_types:
                marker = "</Types>"
                override = (
                    f'<Override PartName="/{CUSTOM_UI_PACKAGE_PATH}" '
                    f'ContentType="{CUSTOM_UI_CONTENT_TYPE}"/>'
                )
                content_types = content_types.replace(marker, override + marker)
                archive.writestr(content_types_path, content_types)

            rels_path = "_rels/.rels"
            rels_xml = archive.read(rels_path).decode("utf-8")
            if CUSTOM_UI_RELATIONSHIP_TYPE not in rels_xml:
                marker = "</Relationships>"
                rel = (
                    '<Relationship Id="rIdCustomUI" '
                    f'Type="{CUSTOM_UI_RELATIONSHIP_TYPE}" '
                    'Target="customUI/customUI.xml"/>'
                )
                rels_xml = rels_xml.replace(marker, rel + marker)
                archive.writestr(rels_path, rels_xml)

        shutil.copy2(temp_workbook, workbook)

    print("custom_ui=applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
