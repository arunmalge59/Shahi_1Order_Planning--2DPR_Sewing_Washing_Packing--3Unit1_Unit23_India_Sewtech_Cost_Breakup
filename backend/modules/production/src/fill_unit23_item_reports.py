#!/usr/bin/env python3
import datetime
import json
import os
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.styles import Alignment


def fail(message):
    print(json.dumps({"error": message}), file=sys.stderr)
    sys.exit(1)


def normalize(value):
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).replace("\xa0", " ").strip()).upper()


def normalize_header(value):
    normalized = normalize(value)
    return normalized.rstrip(".")


def parse_number(value):
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("$", "").replace("\xa0", "")
        if not text:
            return None
        try:
            return float(text)
        except ValueError:
            return None
    return None


def parse_destination_date(value):
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, str):
        text = value.strip()
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y",
                    "%d-%b-%Y", "%d-%b-%y", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(text, fmt).date()
            except ValueError:
                pass
    return None


def load_excel_workbook(filepath):
    if filepath.lower().endswith(".xls"):
        try:
            with open(filepath, "rb") as source_file:
                signature = source_file.read(8)
            if signature != b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
                fail("Report '{}' is not a valid .xls BIFF file.".format(os.path.basename(filepath)))
        except OSError as exc:
            fail("Could not open report '{}': {}".format(os.path.basename(filepath), exc))
    try:
        return openpyxl.load_workbook(filepath, data_only=False)
    except Exception as exc:
        if not filepath.lower().endswith(".xls"):
            raise exc
        try:
            import xlrd
        except ImportError:
            fail("Could not read .xls file: xlrd module is not installed.")
        try:
            legacy = xlrd.open_workbook(filepath)
        except Exception as xl_exc:
            fail("Could not open .xls report: {}".format(xl_exc))

        workbook = openpyxl.Workbook()
        workbook.remove(workbook.active)
        for sheet_name in legacy.sheet_names():
            sheet = legacy.sheet_by_name(sheet_name)
            converted = workbook.create_sheet(title=sheet_name)
            for row_index in range(sheet.nrows):
                for col_index in range(sheet.ncols):
                    value = sheet.cell_value(row_index, col_index)
                    if sheet.cell_type(row_index, col_index) == xlrd.XL_CELL_DATE:
                        try:
                            value = datetime.datetime(*xlrd.xldate_as_tuple(value, legacy.datemode))
                        except Exception:
                            pass
                    elif sheet.cell_type(row_index, col_index) == xlrd.XL_CELL_EMPTY:
                        value = None
                    converted.cell(row=row_index + 1, column=col_index + 1, value=value)
        return workbook


def main():
    if len(sys.argv) != 6:
        fail("Usage: fill_unit23_item_reports.py <workbook> <report> <output> <washing|packing> <YYYY-MM-DD>")

    workbook_path, report_path, output_path, process, date_text = sys.argv[1:6]
    if process not in ("washing", "packing"):
        fail("Process must be washing or packing.")
    try:
        target_date = datetime.datetime.strptime(date_text, "%Y-%m-%d").date()
    except ValueError:
        fail("Invalid report date: {}. Use YYYY-MM-DD.".format(date_text))

    try:
        workbook = openpyxl.load_workbook(workbook_path, data_only=False)
    except Exception as exc:
        fail("Could not open destination workbook: {}".format(exc))

    try:
        report_workbook = load_excel_workbook(report_path)
    except Exception as exc:
        fail("Could not open report '{}': {}".format(os.path.basename(report_path), exc))

    source_sheet = report_workbook.worksheets[0]
    source_rows = list(source_sheet.iter_rows(values_only=True))
    if not source_rows:
        fail("Report '{}' has an empty worksheet.".format(os.path.basename(report_path)))

    normalized_headers = [normalize(value) for value in source_rows[0]]
    required = ("ITEM", "REPORTED QTY")
    missing = [header for header in required if header not in normalized_headers]
    if missing:
        fail("Report '{}' is missing required header(s): {}.".format(
            os.path.basename(report_path), ", ".join(missing)))

    item_index = normalized_headers.index("ITEM")
    quantity_index = normalized_headers.index("REPORTED QTY")
    buyer_index = normalized_headers.index("BUYER") if "BUYER" in normalized_headers else None

    totals_pattern = re.compile(r"^(TOTAL|SUBTOTAL|GRAND TOTAL|SUM)(?:\s|$)", re.I)
    aggregates = defaultdict(float)
    buyers = {}
    for row_number, row in enumerate(source_rows[1:], start=2):
        if all(value is None or (isinstance(value, str) and not value.strip()) for value in row):
            continue
        item_value = row[item_index] if item_index < len(row) else None
        item_text = str(item_value).strip() if item_value is not None else ""
        row_text = " ".join("" if value is None else str(value) for value in row)
        if totals_pattern.search(row_text.strip()):
            continue

        quantity_value = row[quantity_index] if quantity_index < len(row) else None
        if item_text:
            normalized_item = normalize(item_text)
            quantity = parse_number(quantity_value)
            if quantity is None:
                fail("Report '{}' row {} has a blank or invalid Reported Qty.".format(
                    os.path.basename(report_path), row_number))
            aggregates[normalized_item] += quantity
            if buyer_index is not None and buyer_index < len(row) and row[buyer_index] is not None:
                buyers.setdefault(normalized_item, str(row[buyer_index]).strip())
        elif quantity_value is not None and str(quantity_value).strip():
            fail("Report '{}' row {} has Reported Qty but no Item.".format(
                os.path.basename(report_path), row_number))

    item_sheets = defaultdict(set)
    valid_sheets = []
    for sheet in workbook.worksheets:
        if sheet.cell(row=3, column=1).value is None or sheet.cell(row=3, column=3).value is None:
            continue
        if normalize_header(sheet.cell(row=3, column=1).value) != "STYLE" or normalize_header(sheet.cell(row=3, column=2).value) != "ITEM NO":
            continue
        if normalize(sheet.cell(row=3, column=3).value) != "DATE":
            continue
        valid_sheets.append(sheet)
        for row_number in range(4, sheet.max_row + 1):
            item_value = sheet.cell(row=row_number, column=2).value
            if item_value is not None and str(item_value).strip():
                item_sheets[normalize(item_value)].add(sheet.title)

    unmatched = []
    ambiguous_items = []
    missing_dates = []
    duplicate_destinations = []
    written = 0
    sheets_written = set()
    target_column = 11 if process == "washing" else 16

    for normalized_item, quantity in aggregates.items():
        candidate_sheets = item_sheets.get(normalized_item, set())
        source_name = os.path.basename(report_path)
        if not candidate_sheets:
            unmatched.append({
                "item": normalized_item,
                "quantity": quantity,
                "buyer": buyers.get(normalized_item, ""),
                "file": source_name,
                "process": process,
                "date": date_text
            })
            continue
        if len(candidate_sheets) > 1:
            ambiguous_items.append({
                "item": normalized_item,
                "quantity": quantity,
                "buyer": buyers.get(normalized_item, ""),
                "sheets": sorted(candidate_sheets),
                "file": source_name,
                "process": process,
                "date": date_text
            })
            continue

        sheet_name = next(iter(candidate_sheets))
        sheet = workbook[sheet_name]
        matching_date_rows = []
        for row_number in range(4, sheet.max_row + 1):
            parsed = parse_destination_date(sheet.cell(row=row_number, column=3).value)
            if parsed == target_date:
                matching_date_rows.append(row_number)
        if not matching_date_rows:
            missing_dates.append({
                "item": normalized_item,
                "style": sheet_name,
                "file": source_name,
                "process": process,
                "date": date_text
            })
            continue

        valid_destination_rows = [
            row_number for row_number in matching_date_rows
            if normalize(sheet.cell(row=row_number, column=2).value) == normalized_item
        ]
        if len(valid_destination_rows) != 1:
            if len(valid_destination_rows) > 1:
                duplicate_destinations.append({
                    "item": normalized_item,
                    "style": sheet_name,
                    "rows": valid_destination_rows,
                    "file": source_name,
                    "process": process,
                    "date": date_text
                })
            continue

        target_row = valid_destination_rows[0]
        cell = sheet.cell(row=target_row, column=target_column)
        cell.value = int(quantity) if quantity.is_integer() else quantity
        cell.alignment = Alignment(horizontal="center", vertical="center")
        written += 1
        sheets_written.add(sheet_name)

    try:
        workbook.save(output_path)
    except Exception as exc:
        fail("Could not save updated workbook: {}".format(exc))

    print(json.dumps({
        "success": True,
        "process": process,
        "targetDate": date_text,
        "cellsWritten": written,
        "sheetsUpdated": len(sheets_written),
        "styleSheetsUpdated": sorted(sheets_written),
        "unmatchedItems": unmatched,
        "ambiguousItems": ambiguous_items + duplicate_destinations,
        "missingDestinationDates": missing_dates,
        "duplicateDestinationRows": duplicate_destinations
    }))


if __name__ == "__main__":
    main()
