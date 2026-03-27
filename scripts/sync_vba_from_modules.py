#!/usr/bin/env python
"""
Synchronize exported VBA source files back into the workbook through Excel COM.

This script treats `CreateLetter.xlsm.modules/` as the authoritative text source
and updates matching standard modules, class modules, and user forms in the
target workbook. It is intended for local developer automation only and does not
hook into workbook runtime events.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pythoncom
import win32com.client


SUPPORTED_EXTENSIONS = (".bas", ".cls", ".frm")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Sync VBA modules/forms from exported source files into an XLSM workbook."
    )
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument("modules_dir", type=Path, help="Directory with exported VBA source files")
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show Excel while syncing (default: hidden)",
    )
    return parser


def validate_paths(workbook: Path, modules_dir: Path) -> None:
    if not workbook.exists():
        raise FileNotFoundError(f"Workbook not found: {workbook}")
    if not modules_dir.exists():
        raise FileNotFoundError(f"Modules directory not found: {modules_dir}")


def iter_source_files(modules_dir: Path) -> list[Path]:
    source_files: list[Path] = []
    for extension in SUPPORTED_EXTENSIONS:
        source_files.extend(sorted(modules_dir.glob(f"*{extension}")))
    if not source_files:
        raise FileNotFoundError(f"No VBA source files found in {modules_dir}")
    return source_files


def get_component_by_name(project, component_name: str):
    for index in range(1, project.VBComponents.Count + 1):
        component = project.VBComponents(index)
        if component.Name == component_name:
            return component
    return None


def sync_component(project, source_file: Path) -> str:
    component_name = source_file.stem
    existing_component = get_component_by_name(project, component_name)

    if existing_component is not None:
        component_type = existing_component.Type
        if component_type == 100:
            raise RuntimeError(
                f"Refusing to replace document component '{component_name}'. "
                "Document modules must be handled separately."
            )
        project.VBComponents.Remove(existing_component)

    imported_component = project.VBComponents.Import(str(source_file.resolve()))
    return imported_component.Name


def sync_workbook(workbook_path: Path, modules_dir: Path, visible: bool = False) -> int:
    validate_paths(workbook_path, modules_dir)
    source_files = iter_source_files(modules_dir)

    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = visible
    excel.DisplayAlerts = False

    workbook = None
    synced_count = 0

    try:
        workbook = excel.Workbooks.Open(str(workbook_path.resolve()))
        project = workbook.VBProject

        for source_file in source_files:
            synced_name = sync_component(project, source_file)
            synced_count += 1
            print(f"SYNCED {source_file.name} -> {synced_name}")

        workbook.Save()
        return synced_count
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=True)
        excel.Quit()
        pythoncom.CoUninitialize()


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        synced_count = sync_workbook(args.workbook, args.modules_dir, visible=args.visible)
    except Exception as exc:  # noqa: BLE001 - explicit developer tooling script
        print(f"SYNC ERROR: {exc}", file=sys.stderr)
        return 1

    print(f"Synchronization completed. Components synced: {synced_count}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
