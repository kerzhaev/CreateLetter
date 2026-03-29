#!/usr/bin/env python
"""
Inject or refresh Ribbon customUI markup inside an XLSM workbook.

This keeps Ribbon XML source-managed in the repository instead of relying on
manual RibbonX Editor steps.
"""

from __future__ import annotations

import argparse
import shutil
import tempfile
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELATIONSHIPS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CUSTOM_UI_REL_TYPE = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
CUSTOM_UI_PART_NAME = "/customUI/customUI.xml"
CUSTOM_UI_TARGET = "customUI/customUI.xml"
CUSTOM_UI_CONTENT_TYPE = "application/xml"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Inject customUI.xml into an XLSM workbook.")
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument(
        "--xml",
        type=Path,
        default=Path("customUI/customUI.xml"),
        help="Path to the Ribbon XML source file",
    )
    return parser


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def ensure_content_type_entry(content_types_xml: bytes) -> bytes:
    ET.register_namespace("", CONTENT_TYPES_NS)
    root = ET.fromstring(content_types_xml)

    for override in root.findall(qn(CONTENT_TYPES_NS, "Override")):
        if override.attrib.get("PartName") == CUSTOM_UI_PART_NAME:
            override.set("ContentType", CUSTOM_UI_CONTENT_TYPE)
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    ET.SubElement(
        root,
        qn(CONTENT_TYPES_NS, "Override"),
        {"PartName": CUSTOM_UI_PART_NAME, "ContentType": CUSTOM_UI_CONTENT_TYPE},
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def next_relationship_id(root: ET.Element) -> str:
    max_id = 0
    for rel in root.findall(qn(RELATIONSHIPS_NS, "Relationship")):
        rel_id = rel.attrib.get("Id", "")
        if rel_id.startswith("rId"):
            try:
                max_id = max(max_id, int(rel_id[3:]))
            except ValueError:
                continue
    return f"rId{max_id + 1}"


def ensure_root_relationship(rels_xml: bytes) -> bytes:
    ET.register_namespace("", RELATIONSHIPS_NS)
    root = ET.fromstring(rels_xml)

    for rel in root.findall(qn(RELATIONSHIPS_NS, "Relationship")):
        if rel.attrib.get("Type") == CUSTOM_UI_REL_TYPE or rel.attrib.get("Target") == CUSTOM_UI_TARGET:
            rel.set("Type", CUSTOM_UI_REL_TYPE)
            rel.set("Target", CUSTOM_UI_TARGET)
            return ET.tostring(root, encoding="utf-8", xml_declaration=True)

    ET.SubElement(
        root,
        qn(RELATIONSHIPS_NS, "Relationship"),
        {"Id": next_relationship_id(root), "Type": CUSTOM_UI_REL_TYPE, "Target": CUSTOM_UI_TARGET},
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def inject_custom_ui(workbook_path: Path, ribbon_xml_path: Path) -> None:
    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not ribbon_xml_path.exists():
        raise FileNotFoundError(f"Ribbon XML not found: {ribbon_xml_path}")

    ribbon_bytes = ribbon_xml_path.read_bytes()

    with zipfile.ZipFile(workbook_path, "r") as source_zip:
        payload = {entry.filename: source_zip.read(entry.filename) for entry in source_zip.infolist()}

    payload["[Content_Types].xml"] = ensure_content_type_entry(payload["[Content_Types].xml"])
    payload["_rels/.rels"] = ensure_root_relationship(payload["_rels/.rels"])
    payload[CUSTOM_UI_TARGET] = ribbon_bytes

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as temp_file:
        temp_path = Path(temp_file.name)

    try:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as target_zip:
            for filename, content in payload.items():
                target_zip.writestr(filename, content)
        shutil.copyfile(temp_path, workbook_path)
    finally:
        if temp_path.exists():
            temp_path.unlink()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        inject_custom_ui(args.workbook.resolve(), args.xml.resolve())
    except Exception as exc:  # noqa: BLE001 - developer tooling script
        print(f"CUSTOM UI ERROR: {exc}")
        return 1

    print("custom_ui=applied")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
