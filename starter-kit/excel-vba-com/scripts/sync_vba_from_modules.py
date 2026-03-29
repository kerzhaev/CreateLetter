#!/usr/bin/env python
"""
Generic Excel COM sync helper for source-managed VBA projects.

Adjust the workbook path and modules directory to match the target repository.
This script updates standard modules, class modules, and forms from exported
text files into the workbook VBProject.
"""

from __future__ import annotations

import argparse
from pathlib import Path

import pythoncom
import win32com.client


SUPPORTED_EXTENSIONS = (".bas", ".cls", ".frm")
SOURCE_TEXT_ENCODINGS = ("utf-8-sig", "utf-8", "cp1251", "cp866", "mbcs")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Sync VBA modules/forms from exported source files into an XLSM workbook."
    )
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument("modules_dir", type=Path, help="Directory with exported VBA source files")
    parser.add_argument("--visible", action="store_true", help="Show Excel while syncing")
    return parser


def read_source_text(source_file: Path) -> str:
    for encoding in SOURCE_TEXT_ENCODINGS:
        try:
            return source_file.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    return source_file.read_text()


def sanitize_module_source(source_text: str) -> str:
    lines: list[str] = []
    for line in source_text.splitlines():
        stripped = line.strip()
        if stripped.startswith("Attribute VB_"):
            continue
        if stripped.startswith("VERSION 1.0 CLASS"):
            continue
        if stripped == "BEGIN" or stripped == "End":
            continue
        if stripped.startswith("MultiUse ="):
            continue
        lines.append(line)
    return "\n".join(lines).lstrip("\ufeff").strip() + "\n"


def get_component_by_name(project, component_name: str):
    for index in range(1, project.VBComponents.Count + 1):
        component = project.VBComponents(index)
        if component.Name == component_name:
            return component
    return None


def ensure_component(project, source_file: Path):
    component_name = source_file.stem
    extension = source_file.suffix.lower()
    existing = get_component_by_name(project, component_name)

    if existing is not None:
        return existing

    if extension == ".cls":
        component = project.VBComponents.Add(2)
        component.Name = component_name
        return component

    if extension == ".frm":
        raise RuntimeError(
            f"Missing form component '{component_name}'. Create/import the form once first so sync can update it safely."
        )

    component = project.VBComponents.Add(1)
    component.Name = component_name
    return component


def sync_component(project, source_file: Path) -> None:
    component = ensure_component(project, source_file)
    if source_file.suffix.lower() == ".frm":
        # Generic starter behavior: forms should already exist in the workbook once.
        # After that, update the code-behind only.
        source_text = sanitize_module_source(read_source_text(source_file))
    else:
        source_text = sanitize_module_source(read_source_text(source_file))

    code_module = component.CodeModule
    if code_module.CountOfLines > 0:
        code_module.DeleteLines(1, code_module.CountOfLines)
    code_module.AddFromString(source_text)
    print(f"SYNCED {source_file.name} -> {component.Name}")


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    modules_dir = args.modules_dir.resolve()

    if not workbook_path.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook_path}")
    if not modules_dir.exists():
        raise FileNotFoundError(f"Modules directory not found: {modules_dir}")

    source_files: list[Path] = []
    for extension in SUPPORTED_EXTENSIONS:
        source_files.extend(sorted(modules_dir.glob(f"*{extension}")))

    pythoncom.CoInitialize()
    excel = None
    workbook = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = bool(args.visible)
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(str(workbook_path))
        project = workbook.VBProject

        for source_file in source_files:
            sync_component(project, source_file)

        workbook.Save()
        print(f"Synchronization completed. Components synced: {len(source_files)}")
        return 0
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
