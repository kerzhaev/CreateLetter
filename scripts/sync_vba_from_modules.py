#!/usr/bin/env python
"""
Synchronize exported VBA source files back into the workbook through Excel COM.

This script treats `CreateLetter.xlsm.modules/` as the authoritative text source
for standard modules, class modules, and user forms, plus
`CreateLetter.xlsm.document-modules/` for workbook/sheet document modules.
It updates matching components in the target workbook and is intended for local
developer automation only. It does not hook into workbook runtime events.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pythoncom
import win32com.client.gencache
from win32com.client.dynamic import Dispatch


SUPPORTED_EXTENSIONS = (".bas", ".cls", ".frm")
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
        description="Sync VBA modules/forms/document modules from exported source files into an XLSM workbook."
    )
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument("modules_dir", type=Path, help="Directory with exported standard modules, class modules, and forms")
    parser.add_argument(
        "document_modules_dir",
        nargs="?",
        type=Path,
        help="Directory with exported workbook/sheet document modules (.cls). Defaults to <workbook>.document-modules next to the workbook.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel while syncing (default: hidden)",
    )
    return parser


def derive_document_modules_dir(workbook: Path, document_modules_dir: Path | None) -> Path:
    if document_modules_dir is not None:
        return document_modules_dir
    return workbook.parent / f"{workbook.name}.document-modules"


def validate_paths(workbook: Path, modules_dir: Path, document_modules_dir: Path) -> None:
    if not workbook.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook}")
    if not modules_dir.exists():
        raise FileNotFoundError(f"Modules directory not found: {modules_dir}")
    if not document_modules_dir.exists():
        raise FileNotFoundError(f"Document-modules directory not found: {document_modules_dir}")


def iter_source_files(modules_dir: Path, document_modules_dir: Path) -> list[tuple[Path, bool]]:
    source_files: list[tuple[Path, bool]] = []
    for extension in SUPPORTED_EXTENSIONS:
        source_files.extend((path, False) for path in sorted(modules_dir.glob(f"*{extension}")))
    source_files.extend((path, True) for path in sorted(document_modules_dir.glob("*.cls")))
    if not source_files:
        raise FileNotFoundError(
            f"No VBA source files found in {modules_dir} or {document_modules_dir}"
        )
    return source_files


def get_component_by_name(project, component_name: str):
    for index in range(1, project.VBComponents.Count + 1):
        component = project.VBComponents(index)
        if component.Name == component_name:
            return component
    return None


def read_source_text(source_file: Path) -> str:
    last_error: UnicodeDecodeError | None = None

    for encoding in SOURCE_TEXT_ENCODINGS:
        try:
            return source_file.read_text(encoding=encoding)
        except UnicodeDecodeError as exc:
            last_error = exc

    if last_error is not None:
        raise last_error

    return source_file.read_text()


def sanitize_module_source(source_text: str) -> str:
    sanitized_lines: list[str] = []
    previous_line_continues = False

    for line in source_text.splitlines():
        # VB attributes belong to exported files, but AddFromString cannot
        # accept them inside the code pane of an existing component.
        stripped = line.strip()
        if previous_line_continues and not stripped:
            continue
        if stripped.startswith("Attribute "):
            continue
        if stripped == "VERSION 1.0 CLASS":
            continue
        if stripped == "BEGIN":
            continue
        if stripped == "END":
            continue
        if stripped.startswith("MultiUse ="):
            continue
        sanitized_lines.append(line)
        previous_line_continues = stripped.endswith("_")

    return "\r\n".join(sanitized_lines)


def extract_existing_userform_code(source_text: str) -> str:
    source_lines = source_text.splitlines()
    last_attribute_index = -1

    for index, line in enumerate(source_lines):
        if line.strip().startswith("Attribute VB_"):
            last_attribute_index = index

    if last_attribute_index >= 0:
        return sanitize_module_source("\n".join(source_lines[last_attribute_index + 1:]))

    for index, line in enumerate(source_lines):
        if line.strip().lower() == "option explicit":
            return sanitize_module_source("\n".join(source_lines[index:]))

    return sanitize_module_source(source_text)


def sync_component(project, source_file: Path, is_document_module: bool) -> str:
    component_name = source_file.stem
    existing_component = get_component_by_name(project, component_name)
    suffix = source_file.suffix.lower()

    if existing_component is not None:
        component_type = existing_component.Type
        if suffix == ".bas":
            code_module = existing_component.CodeModule
            if code_module.CountOfLines > 0:
                code_module.DeleteLines(1, code_module.CountOfLines)
            code_module.AddFromString(sanitize_module_source(read_source_text(source_file)))
            return existing_component.Name

        if suffix == ".cls" and component_type in (2, 100):
            code_module = existing_component.CodeModule
            if code_module.CountOfLines > 0:
                code_module.DeleteLines(1, code_module.CountOfLines)
            code_module.AddFromString(sanitize_module_source(read_source_text(source_file)))
            return existing_component.Name

        if suffix == ".frm" and component_type == 3:
            code_module = existing_component.CodeModule
            if code_module.CountOfLines > 0:
                code_module.DeleteLines(1, code_module.CountOfLines)
            code_module.AddFromString(extract_existing_userform_code(read_source_text(source_file)))
            return existing_component.Name

        if is_document_module:
            raise RuntimeError(
                f"Document module source '{source_file.name}' matched non-document component type {component_type}"
            )

        project.VBComponents.Remove(existing_component)

    if is_document_module:
        raise RuntimeError(
            f"Document module '{component_name}' does not exist in the workbook and cannot be created as a normal class module"
        )

    if suffix == ".cls":
        class_component = project.VBComponents.Add(2)
        class_component.Name = component_name
        code_module = class_component.CodeModule
        if code_module.CountOfLines > 0:
            code_module.DeleteLines(1, code_module.CountOfLines)
        code_module.AddFromString(sanitize_module_source(read_source_text(source_file)))
        return class_component.Name

    imported_component = project.VBComponents.Import(str(source_file.resolve()))
    return imported_component.Name


def sync_workbook(
    workbook_path: Path,
    modules_dir: Path,
    document_modules_dir: Path | None = None,
    visible: bool = False,
) -> int:
    resolved_document_modules_dir = derive_document_modules_dir(workbook_path, document_modules_dir)
    validate_paths(workbook_path, modules_dir, resolved_document_modules_dir)
    source_files = iter_source_files(modules_dir, resolved_document_modules_dir)

    pythoncom.CoInitialize()
    reset_excel_gen_cache()
    excel = Dispatch("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    workbook = None
    synced_count = 0

    try:
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()))
        project = workbook.VBProject

        for source_file, is_document_module in source_files:
            synced_name = sync_component(project, source_file, is_document_module)
            synced_count += 1
            print(f"SYNCED {source_file.name} -> {synced_name}")

        workbook.Save()
        return synced_count
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        synced_count = sync_workbook(
            args.workbook,
            args.modules_dir,
            document_modules_dir=args.document_modules_dir,
            visible=args.visible,
        )
    except Exception as exc:  # noqa: BLE001 - explicit developer tooling script
        print(f"SYNC ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Synchronization completed. Components synced: {synced_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
