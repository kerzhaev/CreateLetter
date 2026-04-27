#!/usr/bin/env python
"""
Export workbook VBA source into source-managed directories.

This helper complements `sync_vba_from_modules.py` by exporting standard
modules, class modules, and user forms into the modules directory, plus
workbook/sheet document modules into a separate document-modules directory.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pythoncom
import win32com.client.gencache
from win32com.client.dynamic import Dispatch


COMPONENT_TYPE_EXTENSIONS = {
    1: ".bas",   # Standard module
    2: ".cls",   # Class module
    3: ".frm",   # UserForm
    100: ".cls", # Document module
}
SOURCE_TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "cp866", "mbcs")


def reset_excel_gen_cache() -> None:
    gen_path = Path(win32com.client.gencache.GetGeneratePath())
    for child in gen_path.glob("00020813-0000-0000-C000-000000000046*"):
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
        elif child.exists():
            child.unlink(missing_ok=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Export VBA modules/forms/document modules from an XLSM workbook."
    )
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument("modules_dir", type=Path, help="Directory to receive exported standard modules, class modules, and forms")
    parser.add_argument(
        "document_modules_dir",
        nargs="?",
        type=Path,
        help="Directory to receive exported workbook/sheet document modules (.cls). Defaults to <workbook>.document-modules next to the workbook.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel while exporting (default: hidden)",
    )
    return parser


def derive_document_modules_dir(workbook: Path, document_modules_dir: Path | None) -> Path:
    if document_modules_dir is not None:
        return document_modules_dir
    return workbook.parent / f"{workbook.name}.document-modules"


def validate_paths(workbook: Path, modules_dir: Path, document_modules_dir: Path) -> None:
    if not workbook.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook}")
    modules_dir.mkdir(parents=True, exist_ok=True)
    document_modules_dir.mkdir(parents=True, exist_ok=True)


def normalize_exported_text_file(source_file: Path) -> None:
    raw_bytes = source_file.read_bytes()

    for encoding in SOURCE_TEXT_ENCODINGS:
        try:
            decoded_text = raw_bytes.decode(encoding)
            source_file.write_text(decoded_text, encoding="utf-8")
            return
        except UnicodeDecodeError:
            continue

    source_file.write_bytes(raw_bytes)


def get_code_module_text(component) -> str:
    line_count = component.CodeModule.CountOfLines
    if line_count <= 0:
        return ""
    code_text = component.CodeModule.Lines(1, line_count)
    return code_text.replace("\r\r\n", "\r\n").replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


def build_standard_module_source(component) -> str:
    lines = [f'Attribute VB_Name = "{component.Name}"']
    code_text = get_code_module_text(component)
    if code_text:
        lines.append(code_text)
    return "\r\n".join(lines)


def build_document_module_source(component) -> str:
    code_text = get_code_module_text(component)
    lines = [
        "VERSION 1.0 CLASS",
        "BEGIN",
        "  MultiUse = -1  'True",
        "END",
        f'Attribute VB_Name = "{component.Name}"',
        "Attribute VB_GlobalNameSpace = False",
        "Attribute VB_Creatable = False",
        "Attribute VB_PredeclaredId = True",
        "Attribute VB_Exposed = True",
    ]
    if code_text:
        lines.append(code_text)
    return "\r\n".join(lines)


def export_component(component, modules_dir: Path, document_modules_dir: Path) -> str | None:
    extension = COMPONENT_TYPE_EXTENSIONS.get(component.Type)
    if extension is None:
        return None

    destination = modules_dir / f"{component.Name}{extension}"

    if component.Type == 1:
        destination.write_text(build_standard_module_source(component), encoding="utf-8")
        return destination.name

    if component.Type == 100:
        destination = document_modules_dir / f"{component.Name}{extension}"
        destination.write_text(build_document_module_source(component), encoding="utf-8")
        return destination.name

    component.Export(str(destination.resolve()))
    normalize_exported_text_file(destination)
    return destination.name


def export_workbook(
    workbook_path: Path,
    modules_dir: Path,
    document_modules_dir: Path | None = None,
    visible: bool = False,
) -> int:
    resolved_document_modules_dir = derive_document_modules_dir(workbook_path, document_modules_dir)
    validate_paths(workbook_path, modules_dir, resolved_document_modules_dir)

    pythoncom.CoInitialize()
    reset_excel_gen_cache()
    excel = Dispatch("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    workbook = None
    exported_count = 0

    try:
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()), False, True)
        project = workbook.VBProject

        for index in range(1, project.VBComponents.Count + 1):
            component = project.VBComponents(index)
            exported_name = export_component(component, modules_dir, resolved_document_modules_dir)
            if exported_name is None:
                continue
            exported_count += 1
            print(f"EXPORTED {component.Name} -> {exported_name}")

        return exported_count
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        exported_count = export_workbook(
            args.workbook,
            args.modules_dir,
            document_modules_dir=args.document_modules_dir,
            visible=args.visible,
        )
    except Exception as exc:  # noqa: BLE001 - explicit developer tooling script
        print(f"EXPORT ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Export completed. Components exported: {exported_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
