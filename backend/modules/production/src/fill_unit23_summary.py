#!/usr/bin/env python3
import datetime
import json
import re
import sys
from collections import defaultdict

import openpyxl
from openpyxl.utils import get_column_letter


def fail(message):
    print(json.dumps({"error": message}), file=sys.stderr)
    sys.exit(1)


def norm(value):
    if value is None:
        return ""
    text = str(value).replace("\xa0", " ").replace("\r", " ").replace("\n", " ").lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def text(value):
    return "" if value is None else str(value).strip()


def parse_date(value):
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            return datetime.datetime(*openpyxl.utils.datetime.from_excel(value).timetuple()[:6]).date()
        except Exception:
            return None
    if isinstance(value, str):
        raw = value.strip()
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y", "%d-%b-%y", "%Y/%m/%d"):
            try:
                return datetime.datetime.strptime(raw, fmt).date()
            except ValueError:
                pass
    return None


def number_value(value, cached, row, column):
    if isinstance(value, str):
        raw = value.strip().replace(",", "")
        if raw == "":
            return None
        try:
            return float(raw)
        except ValueError:
            raise ValueError("Invalid quantity '{}' at {}{}".format(value, column, row))
    if isinstance(value, bool):
        raise ValueError("Invalid quantity at {}{}".format(column, row))
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str) and value.startswith("="):
        if isinstance(cached, (int, float)) and not isinstance(cached, bool):
            return float(cached)
        raise ValueError("Formula quantity at {}{} has no cached result".format(column, row))
    if value is None:
        return None
    raise ValueError("Invalid quantity at {}{}".format(column, row))


def unpack_merged(ws):
    resolved = {}
    for row in ws.iter_rows():
        for cell in row:
            resolved[cell.coordinate] = cell.value
    for merged in ws.merged_cells.ranges:
        value = ws.cell(merged.min_row, merged.min_col).value
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                resolved["{}{}".format(get_column_letter(col), row)] = value
    return resolved


def find_source_headers(ws, values):
    headers = []
    for row in range(1, ws.max_row + 1):
        for col in range(1, ws.max_column + 1):
            if norm(values["{}{}".format(get_column_letter(col), row)]) in ("style no", "item no", "delivery date"):
                headers.append(row)
    for row in range(1, ws.max_row + 1):
        names = [norm(values["{}{}".format(get_column_letter(c), row)]) for c in range(1, ws.max_column + 1)]
        if ("style no" in names and "delivery date" in names and
                any("sewing" in n for n in names) and any("till date" in n for n in names)):
            headers.append(row)
    return sorted(set(headers))


def source_header_columns(ws, values, header_row):
    child = {col: norm(values.get("{}{}".format(get_column_letter(col), header_row)))
             for col in range(1, ws.max_column + 1)}
    parent = {}
    for col in range(1, ws.max_column + 1):
        found = ""
        for candidate in range(header_row - 3, header_row + 1):
            if candidate < 1:
                continue
            found = norm(values.get("{}{}".format(get_column_letter(col), candidate)))
            if found and found != "till date":
                break
        parent[col] = found

    aliases = {
        "style": ("style no",),
        "item": ("item no",),
        "po": ("po",),
        "po_qty": ("po qty",),
        "country": ("country",),
        "delivery": ("delivery date",),
        "product": ("product type",),
    }
    columns = {}
    for key, names in aliases.items():
        found = [col for col, value in child.items() if value in names]
        if found:
            columns[key] = found[0]

    process_parents = (
        ("sewing", "sewing out"),
        ("washing", "washing out"),
        ("packing", "finishing out"),
    )
    for key, process_name in process_parents:
        found = [col for col in range(1, ws.max_column + 1)
                 if process_name in parent[col] and child[col] == "till date"]
        if found:
            columns[key] = found[0]

    required = ("style", "item", "po", "delivery", "sewing", "washing", "packing")
    missing = [key for key in required if key not in columns]
    if missing:
        raise ValueError("Cumulative report header row {} is missing required source columns: {}."
                         .format(header_row, ", ".join(missing)))
    return columns


def is_header_row(ws, values, row):
    child = [norm(values.get("{}{}".format(get_column_letter(col), row)))
             for col in range(1, ws.max_column + 1)]
    return "style no" in child and "item no" in child and "delivery date" in child and \
        any(name == "till date" for name in child)


def is_title_row(ws, values, row):
    populated = [text(values.get("{}{}".format(get_column_letter(col), row)))
                 for col in range(1, ws.max_column + 1)
                 if text(values.get("{}{}".format(get_column_letter(col), row)))]
    if len(populated) != 1:
        return False
    value = populated[0]
    return bool(re.match(r"^[^-]+-\s*[A-Z0-9]+\s*\(", value.strip(), re.IGNORECASE))


def is_section_boundary(values, row, ws):
    populated = [norm(values.get("{}{}".format(get_column_letter(col), row)))
                 for col in range(1, ws.max_column + 1)]
    joined = " ".join(value for value in populated if value)
    return bool(re.search(r"(^|\s)(grand total|sub total|subtotal|total)(\s|$)", joined))


def is_note_row(values, row, ws):
    joined = " ".join(norm(values.get("{}{}".format(get_column_letter(col), row)))
                      for col in range(1, ws.max_column + 1)).strip()
    return bool(joined and
                (("red colour" in joined and "rejected pcs" in joined) or
                 ("orange colour" in joined and "sample pcs" in joined)))


def parse_source(path):
    workbook = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    ws = workbook.worksheets[0]
    cached_ws = cached.worksheets[0]
    values = unpack_merged(ws)
    cached_values = unpack_merged(cached_ws)
    aggregated = defaultdict(lambda: [0.0, 0.0, 0.0])
    seen_detail = False
    columns = None
    context = {"style": "", "item": "", "po": "", "po_qty": "", "country": "",
               "delivery": None, "product": ""}

    def reset_context():
        return {"style": "", "item": "", "po": "", "po_qty": "", "country": "",
                "delivery": None, "product": ""}

    for row in range(1, ws.max_row + 1):
        populated = "".join(text(values.get("{}{}".format(get_column_letter(col), row)))
                            for col in range(1, ws.max_column + 1)).strip()
        if not populated:
            continue
        if is_note_row(values, row, ws) or is_title_row(ws, values, row):
            context = reset_context()
            columns = None
            continue
        if is_section_boundary(values, row, ws):
            context = reset_context()
            continue
        if is_header_row(ws, values, row):
            columns = source_header_columns(ws, values, row)
            continue
        if columns is None:
            continue

        def value(key):
            col = columns.get(key)
            return None if col is None else values.get("{}{}".format(get_column_letter(col), row))

        explicit_style = text(value("style"))
        explicit_item = text(value("item"))
        explicit_po = text(value("po"))
        explicit_po_qty = value("po_qty")
        explicit_country = text(value("country"))
        explicit_delivery = value("delivery")
        explicit_product = text(value("product"))

        if explicit_style:
            context["style"] = explicit_style
        if explicit_item:
            context["item"] = explicit_item
        if explicit_po_qty not in (None, ""):
            context["po_qty"] = explicit_po_qty
        if explicit_country:
            context["country"] = explicit_country
        if explicit_product:
            context["product"] = explicit_product
        if explicit_po:
            if explicit_po != context["po"]:
                context["delivery"] = None
            context["po"] = explicit_po

        if explicit_delivery not in (None, ""):
            parsed_delivery = parse_date(explicit_delivery)
            if parsed_delivery is None:
                raise ValueError("Invalid delivery date at row {}.".format(row))
            context["delivery"] = parsed_delivery

        quantity_values = [value(key) for key in ("sewing", "washing", "packing")]
        has_identifier = explicit_style or explicit_item or explicit_po
        has_quantity = any(raw not in (None, "") for raw in quantity_values)
        if not has_identifier and not has_quantity:
            continue
        if not context["style"] or not context["item"] or not context["po"] or context["delivery"] is None:
            raise ValueError("Detail row {} has unresolved style, item, PO, or delivery date context.".format(row))

        quantities = []
        for key in ("sewing", "washing", "packing"):
            col = columns[key]
            coord = "{}{}".format(get_column_letter(col), row)
            raw = value(key)
            if isinstance(raw, str) and raw.startswith("="):
                raw_cached = cached_values.get(coord)
                if raw_cached is None:
                    raise ValueError("Formula quantity at {} has no cached result.".format(coord))
                raw = raw_cached
            qty = number_value(raw, None, row, get_column_letter(col))
            if qty is None:
                raise ValueError("{} at row {} has a blank required quantity.".format(key.upper(), row))
            quantities.append(qty)

        aggregate_key = (norm(context["style"]), norm(context["item"]), context["delivery"].isoformat())
        for index, qty in enumerate(quantities):
            aggregated[aggregate_key][index] += qty
        seen_detail = True
    if not seen_detail:
        fail("Cumulative report contains no valid detail rows.")
    return aggregated


def find_summary_header(ws):
    aliases = {
        "xfty": ("x fty", "xfty", "x fty date"),
        "qty": ("qty", "quantity"),
        "sewing": ("sewing completed", "sewing completd"),
        "sewing_balance": ("bal to sew",),
        "wash": ("wash completed", "wash completd"),
        "wash_balance": ("bal to wash",),
        "pack": ("pack completed", "pack completd"),
        "pack_balance": ("bal to pack",),
    }
    for row in range(1, ws.max_row + 1):
        found = {}
        for col in range(1, ws.max_column + 1):
            header = norm(ws.cell(row=row, column=col).value)
            for key, names in aliases.items():
                if header in names:
                    found[key] = col
        if {"xfty", "qty", "sewing", "sewing_balance", "wash", "wash_balance", "pack", "pack_balance"} <= set(found):
            return row, found
    return None, {}


def date_key(value):
    parsed = parse_date(value)
    return parsed.isoformat() if parsed else ""


def write_summary(workbook, aggregates):
    updated_rows = 0
    completion_cells = 0
    updated_sheets = set()
    unmatched = []
    ambiguous = []
    for ws in workbook.worksheets:
        header_row, cols = find_summary_header(ws)
        if header_row is None:
            continue

        daily_pairs = set()
        for row in range(4, ws.max_row + 1):
            style = text(ws.cell(row=row, column=1).value)
            item = text(ws.cell(row=row, column=2).value)
            if style and item:
                daily_pairs.add((norm(style), norm(item)))

        summary_style_col = next((c for c in range(1, ws.max_column + 1)
                                  if norm(ws.cell(header_row, c).value) == "style"), None)
        for row in range(header_row + 1, ws.max_row + 1):
            xfty = ws.cell(row=row, column=cols["xfty"]).value
            delivery = date_key(xfty)
            if not delivery:
                continue
            summary_style = norm(ws.cell(row=row, column=summary_style_col).value) if summary_style_col else ""
            candidates = set()
            if summary_style:
                candidates.update(pair for pair in daily_pairs if pair[0] == summary_style)
            else:
                candidates.update(daily_pairs)
            if len(daily_pairs) == 1 and not summary_style:
                candidates = daily_pairs
            if len(candidates) != 1:
                ambiguous.append({"sheet": ws.title, "style": summary_style, "xfty": delivery,
                                  "reason": "Destination Style/Item context is absent or ambiguous."})
                continue
            style, item = next(iter(candidates))
            source_key = (style, item, delivery)
            actual = aggregates.get(source_key)
            if actual is None:
                unmatched.append({"sheet": ws.title, "style": style, "item": item, "xfty": delivery})
                continue

            qty_cell = "{}{}".format(get_column_letter(cols["qty"]), row)
            process_columns = (
                ("sewing", cols["sewing"], cols["sewing_balance"]),
                ("wash", cols["wash"], cols["wash_balance"]),
                ("pack", cols["pack"], cols["pack_balance"]),
            )
            for index, (name, completed_col, balance_col) in enumerate(process_columns):
                ws.cell(row=row, column=completed_col, value=clean_number(actual[index]))
                completion_cells += 1
                completed_cell = "{}{}".format(get_column_letter(completed_col), row)
                ws.cell(row=row, column=balance_col, value="={}-{}{}".format(
                    completed_cell, get_column_letter(cols["qty"]), row))
            updated_rows += 1
            updated_sheets.add(ws.title)
    return {
        "summaryRowsUpdated": updated_rows,
        "completionCellsWritten": completion_cells,
        "styleSheetsUpdated": sorted(updated_sheets),
        "unmatchedGroups": unmatched,
        "ambiguousGroups": ambiguous,
    }


def clean_number(value):
    return int(value) if float(value).is_integer() else value


def main():
    if len(sys.argv) != 4:
        fail("Usage: fill_unit23_summary.py <workbook> <summary report> <output>")
    workbook_path, report_path, output_path = sys.argv[1:4]
    try:
        aggregates = parse_source(report_path)
    except (ValueError, OSError) as exc:
        fail(str(exc))
    try:
        workbook = openpyxl.load_workbook(workbook_path, data_only=False)
    except Exception as exc:
        fail("Could not open destination workbook: {}".format(exc))
    summary = write_summary(workbook, aggregates)
    try:
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.calculation.calcMode = "auto"
        workbook.save(output_path)
    except Exception as exc:
        fail("Could not save updated workbook: {}".format(exc))
    print(json.dumps(dict({"success": True, "process": "summary"}, **summary)))


if __name__ == "__main__":
    main()
