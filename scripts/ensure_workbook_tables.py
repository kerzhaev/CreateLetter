#!/usr/bin/env python
"""
Ensure structured tables and workbook schema helpers exist in the CreateLetter workbook.

This script creates workbook ListObjects for the core CreateLetter sheets and the
mail-dispatch foundation sheets without changing existing business data layout.
It is safe to run multiple times.
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import pythoncom
import win32com.client.gencache
from win32com.client.dynamic import Dispatch


TABLE_SPECS = (
    ("Addresses", "tblAddresses", ("Наименование адресата", "Улица, дом, квартира", "Населенный пункт", "Район", "Область/край/республика", "Почтовый индекс", "Телефон", "AddressGroup")),
    ("Letters", "tblLetters", ("Наименование адресата", "Исходящий номер", "Дата исходящего", "Наименование приложения", "Сумма документа", "Отметка о возврате", "Исполнитель", "Тип отправки", "Упаковано", "Пакет", "Номер реестра", "Дата реестра")),
    ("Settings", "tblLetterTexts", None),
    ("EnvelopeFormats", "tblEnvelopeFormats", ("FormatKey", "DisplayName", "IsActive", "SortOrder")),
    ("Senders", "tblSenders", ("SenderName", "AddressLine1", "AddressLine2", "AddressLine3", "PostalCode", "Phone", "IsDefault")),
    ("DispatchItems", "tblDispatchItems", ("DispatchId", "LetterNumber", "LetterDate", "LetterRowNumber", "Addressee", "AddressLine", "PostalCode", "SenderName", "EnvelopeFormatKey", "MailType", "Mass", "DeclaredValue", "Comment", "Phone", "BatchId", "Status", "CreatedAt", "RegistryNumber", "RegistryDate")),
    ("DispatchRegistry", "tblDispatchRegistry", ("RegistryNumber", "RegistryDate", "BatchId", "Addressee", "AddressLine", "EnvelopeFormatKey", "MailType", "Mass", "DeclaredValue", "Payment", "Comment", "Phone", "IndexFrom", "SenderName", "OutgoingNumbers", "CreatedAt", "PostalCode")),
)

LAYOUT_SHEET_SPECS = (
    ("DispatchLayout_C4", ("BatchId", "RegistryNumber", "RegistryDate", "Addressee", "AddressLine", "PostalCode", "SenderName", "IndexFrom", "OutgoingNumbers", "EnvelopeFormatKey", "MailType", "Mass", "DeclaredValue", "Comment", "PreparedAt")),
    ("DispatchLayout_C5", ("BatchId", "RegistryNumber", "RegistryDate", "Addressee", "AddressLine", "PostalCode", "SenderName", "IndexFrom", "OutgoingNumbers", "EnvelopeFormatKey", "MailType", "Mass", "DeclaredValue", "Comment", "PreparedAt")),
    ("DispatchLayout_DL", ("BatchId", "RegistryNumber", "RegistryDate", "Addressee", "AddressLine", "PostalCode", "SenderName", "IndexFrom", "OutgoingNumbers", "EnvelopeFormatKey", "MailType", "Mass", "DeclaredValue", "Comment", "PreparedAt")),
)

PRINT_SHEET_NAMES = (
    "PostalRegistryPrint",
)

ADDRESS_GROUP_COLUMN_NAME = "AddressGroup"
ENVELOPE_FORMAT_DEFAULT_ROWS = (
    ("c4", "C4", True, 10),
    ("c5", "C5", True, 20),
    ("dl", "DL", True, 30),
)

XL_SRC_RANGE = 1
XL_YES = 1


def safe_set_com_property(com_object, property_name: str, value) -> None:
    try:
        setattr(com_object, property_name, value)
    except Exception:
        pass


def reset_excel_gen_cache() -> None:
    gen_path = Path(win32com.client.gencache.GetGeneratePath())
    for child in gen_path.glob("00020813-0000-0000-C000-000000000046*"):
        if child.is_dir():
            shutil.rmtree(child, ignore_errors=True)
        elif child.exists():
            child.unlink(missing_ok=True)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Create missing ListObjects in the workbook.")
    parser.add_argument("workbook", type=Path, help="Path to the target .xlsm workbook")
    parser.add_argument("--visible", action="store_true", help="Show Excel during the operation")
    return parser


def ensure_table(ws, table_name: str, headers: tuple[str, ...] | None = None) -> str:
    for index in range(1, ws.ListObjects.Count + 1):
        if ws.ListObjects(index).Name == table_name:
            return "existing"

    if table_name == "tblLetterTexts":
        for index in range(1, ws.ListObjects.Count + 1):
            if ws.ListObjects(index).Name in {"Text", "Текст"}:
                ws.ListObjects(index).Name = table_name
                return "renamed"

    if headers is not None:
        first_row = 1
        first_col = 1
        last_col = len(headers)
        last_row = max(2, ws.Cells(ws.Rows.Count, first_col).End(-4162).Row)  # xlUp
        source_range = ws.Range(ws.Cells(first_row, first_col), ws.Cells(last_row, last_col))
    else:
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


def get_or_create_sheet(workbook, sheet_name: str):
    for index in range(1, workbook.Worksheets.Count + 1):
        ws = workbook.Worksheets(index)
        if ws.Name == sheet_name:
            return ws, False

    ws = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
    ws.Name = sheet_name
    return ws, True


def ensure_sheet_headers(ws, headers: tuple[str, ...]) -> str:
    if not headers:
        return "skipped"

    created = False
    for col_index, header in enumerate(headers, start=1):
        if ws.Cells(1, col_index).Value != header:
            ws.Cells(1, col_index).Value = header
            created = True

    if ws.UsedRange.Rows.Count < 2:
        for col_index in range(1, len(headers) + 1):
            if ws.Cells(2, col_index).Value is None:
                ws.Cells(2, col_index).Value = ""
        created = True

    return "updated" if created else "existing"


def ensure_envelope_formats_seed(ws) -> str:
    existing_keys: set[str] = set()
    last_row = ws.Cells(ws.Rows.Count, 1).End(-4162).Row  # xlUp

    for row_index in range(2, last_row + 1):
        key_value = ws.Cells(row_index, 1).Value
        if key_value is not None and str(key_value).strip():
            existing_keys.add(str(key_value).strip().lower())

    next_row = max(2, last_row + 1)
    created = 0
    for format_key, display_name, is_active, sort_order in ENVELOPE_FORMAT_DEFAULT_ROWS:
        if format_key in existing_keys:
            continue
        ws.Cells(next_row, 1).Value = format_key
        ws.Cells(next_row, 2).Value = display_name
        ws.Cells(next_row, 3).Value = is_active
        ws.Cells(next_row, 4).Value = sort_order
        next_row += 1
        created += 1

    return "created" if created > 0 else "existing"


def ensure_address_group_column(ws) -> str:
    target_table = None
    for index in range(1, ws.ListObjects.Count + 1):
        if ws.ListObjects(index).Name == "tblAddresses":
            target_table = ws.ListObjects(index)
            break

    if target_table is None:
        return "skipped-no-table"

    for index in range(1, target_table.ListColumns.Count + 1):
        if target_table.ListColumns(index).Name == ADDRESS_GROUP_COLUMN_NAME:
            return "existing"

    target_table.ListColumns.Add()
    target_table.ListColumns(target_table.ListColumns.Count).Name = ADDRESS_GROUP_COLUMN_NAME
    return "created"


def ensure_table_columns(ws, table_name: str, headers: tuple[str, ...] | None) -> str:
    if not headers:
        return "skipped"

    target_table = None
    for index in range(1, ws.ListObjects.Count + 1):
        if ws.ListObjects(index).Name == table_name:
            target_table = ws.ListObjects(index)
            break

    if target_table is None:
        return "skipped-no-table"

    updated = False
    for col_index, header in enumerate(headers, start=1):
        if col_index <= target_table.ListColumns.Count:
            if target_table.ListColumns(col_index).Name != header:
                target_table.ListColumns(col_index).Name = header
                updated = True
        else:
            target_table.ListColumns.Add()
            target_table.ListColumns(target_table.ListColumns.Count).Name = header
            updated = True

    return "updated" if updated else "existing"


def migrate_dispatch_items_legacy_layout(ws) -> str:
    target_table = None
    for index in range(1, ws.ListObjects.Count + 1):
        if ws.ListObjects(index).Name == "tblDispatchItems":
            target_table = ws.ListObjects(index)
            break

    if target_table is None:
        return "skipped-no-table"

    if target_table.DataBodyRange is None:
        return "skipped-empty"

    row_count = target_table.DataBodyRange.Rows.Count
    col_count = target_table.DataBodyRange.Columns.Count
    if row_count < 1 or col_count < 16:
        return "skipped-insufficient-data"

    raw_values = target_table.DataBodyRange.Value
    if row_count == 1 and not isinstance(raw_values, tuple):
        raw_values = (raw_values,)

    def normalize_text(value) -> str:
        if value is None:
            return ""
        return str(value).strip()

    def looks_like_legacy_row(row_values) -> bool:
        envelope_candidate = normalize_text(row_values[8]).lower()
        batch_candidate = normalize_text(row_values[14]).lower()
        letter_row_candidate = normalize_text(row_values[3])
        valid_envelopes = {"", "c4", "c5", "dl"}
        valid_statuses = {"draft", "queued", "registered"}

        return (
            envelope_candidate not in valid_envelopes
            and batch_candidate in valid_statuses
            and letter_row_candidate != ""
            and not letter_row_candidate.isdigit()
        )

    if not any(looks_like_legacy_row(row_values) for row_values in raw_values):
        return "existing"

    migrated_rows: list[list[str]] = []
    for row_values in raw_values:
        migrated_rows.append([
            row_values[0],   # DispatchId
            row_values[1],   # LetterNumber
            row_values[2],   # LetterDate
            "",              # LetterRowNumber (was absent in legacy layout)
            row_values[3],   # Addressee
            row_values[4],   # AddressLine
            row_values[5],   # PostalCode
            row_values[6],   # SenderName
            row_values[7],   # EnvelopeFormatKey
            row_values[8],   # MailType
            row_values[9],   # Mass
            row_values[10],  # DeclaredValue
            row_values[11],  # Comment
            row_values[12],  # Phone
            row_values[13],  # BatchId
            row_values[14],  # Status
            row_values[15],  # CreatedAt
            "",              # RegistryNumber
            "",              # RegistryDate
        ])

    target_table.DataBodyRange.ClearContents()
    write_range = ws.Range(ws.Cells(2, 1), ws.Cells(1 + len(migrated_rows), 19))
    write_range.Value = migrated_rows
    return "migrated"


def ensure_layout_sheet(workbook, sheet_name: str, headers: tuple[str, ...]) -> str:
    ws, sheet_created = get_or_create_sheet(workbook, sheet_name)
    header_status = ensure_sheet_headers(ws, headers)
    ws.Visible = 2  # xlSheetVeryHidden
    if sheet_created:
        return "created"
    if header_status == "updated":
        return "updated"
    return "existing"


def ensure_print_sheet(workbook, sheet_name: str) -> str:
    _, sheet_created = get_or_create_sheet(workbook, sheet_name)
    return "created" if sheet_created else "existing"


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    workbook_path = args.workbook.resolve()
    if not workbook_path.exists():
        print(f"Workbook not found: {workbook_path}", file=sys.stderr)
        return 1

    pythoncom.CoInitialize()
    reset_excel_gen_cache()
    try:
        excel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    except Exception:
        excel = Dispatch("Excel.Application")

    safe_set_com_property(excel, "Visible", args.visible)
    safe_set_com_property(excel, "DisplayAlerts", False)
    workbook = None

    try:
        workbook = excel.Workbooks.Open(str(workbook_path))
        for sheet_name, table_name, headers in TABLE_SPECS:
            ws, sheet_created = get_or_create_sheet(workbook, sheet_name)
            if headers is not None:
                header_status = ensure_sheet_headers(ws, headers)
                print(f"{sheet_name}:headers:{header_status}")
            status = ensure_table(ws, table_name, headers=headers)
            print(f"{sheet_name}:{table_name}:{status}")
            if headers is not None:
                table_column_status = ensure_table_columns(ws, table_name, headers)
                print(f"{sheet_name}:{table_name}:columns:{table_column_status}")
                if table_name == "tblDispatchItems":
                    dispatch_migration_status = migrate_dispatch_items_legacy_layout(ws)
                    print(f"{sheet_name}:{table_name}:legacy-migration:{dispatch_migration_status}")

            if table_name == "tblEnvelopeFormats":
                seed_status = ensure_envelope_formats_seed(ws)
                print(f"{sheet_name}:seed:{seed_status}")

        for sheet_name, headers in LAYOUT_SHEET_SPECS:
            layout_status = ensure_layout_sheet(workbook, sheet_name, headers)
            print(f"{sheet_name}:layout:{layout_status}")

        for sheet_name in PRINT_SHEET_NAMES:
            print_status = ensure_print_sheet(workbook, sheet_name)
            print(f"{sheet_name}:print:{print_status}")

        address_group_status = ensure_address_group_column(workbook.Worksheets("Addresses"))
        print(f"Addresses:{ADDRESS_GROUP_COLUMN_NAME}:{address_group_status}")

        workbook.Save()
        return 0
    except Exception as exc:  # noqa: BLE001 - developer tooling script
        print(f"TABLE BOOTSTRAP ERROR: {exc}", file=sys.stderr)
        return 1
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=True)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    raise SystemExit(main())
