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
        "colour": ("colour", "color"),
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


def parse_source_allocations(path):
    workbook = openpyxl.load_workbook(path, data_only=False)
    cached = openpyxl.load_workbook(path, data_only=True)
    ws = workbook.worksheets[0]
    cached_ws = cached.worksheets[0]
    values = unpack_merged(ws)
    cached_values = unpack_merged(cached_ws)
    allocations = []
    allocation_by_signature = {}
    ambiguous_signatures = set()
    seen_detail = False
    columns = None
    current = None
    context = {"style": "", "item": "", "po": "", "po_qty": "", "country": "",
               "colour": "", "delivery": None, "product": ""}

    def reset_context():
        return {"style": "", "item": "", "po": "", "po_qty": "", "country": "",
                "colour": "", "delivery": None, "product": ""}

    for row in range(1, ws.max_row + 1):
        populated = "".join(text(values.get("{}{}".format(get_column_letter(col), row)))
                            for col in range(1, ws.max_column + 1)).strip()
        if not populated:
            continue
        if is_note_row(values, row, ws) or is_title_row(ws, values, row):
            context = reset_context()
            current = None
            columns = None
            continue
        if is_section_boundary(values, row, ws):
            context = reset_context()
            current = None
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
        explicit_colour = text(value("colour"))
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
        if explicit_colour:
            context["colour"] = explicit_colour
        if explicit_product:
            context["product"] = explicit_product
        if explicit_po:
            context["po"] = explicit_po

        parsed_delivery = context["delivery"]
        if explicit_delivery not in (None, ""):
            parsed_delivery = parse_date(explicit_delivery)
            if parsed_delivery is None:
                raise ValueError("Invalid delivery date at row {}.".format(row))
            context["delivery"] = parsed_delivery

        quantity_values = [value(key) for key in ("sewing", "washing", "packing")]
        has_identifier = explicit_style or explicit_item or explicit_po or explicit_country or explicit_colour
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

        signature = (
            norm(context["style"]), norm(context["item"]), text(context["po"]),
            context["delivery"].isoformat(), norm(context["country"]), norm(context["colour"])
        )
        if current is None or current["signature"] != signature:
            previous = allocation_by_signature.get(signature)
            if previous is not None and current is not None and current is not previous:
                ambiguous_signatures.add(signature)
            if previous is None or previous is not current:
                current = {
                    "signature": signature,
                    "style": context["style"],
                    "item": context["item"],
                    "po": text(context["po"]),
                    "delivery": context["delivery"].isoformat(),
                    "country": context["country"],
                    "colour": context["colour"],
                    "po_qty": context["po_qty"],
                    "totals": [0.0, 0.0, 0.0],
                    "rows": [],
                }
                allocations.append(current)
                allocation_by_signature[signature] = current
        if current["po_qty"] in (None, ""):
            current["po_qty"] = context["po_qty"]
        current["rows"].append(row)
        for index, qty in enumerate(quantities):
            current["totals"][index] += qty
        seen_detail = True

    if not seen_detail:
        fail("Cumulative report contains no valid detail rows.")

    allocations_by_source_signature = defaultdict(list)
    ambiguous = []
    aggregate = defaultdict(lambda: [0.0, 0.0, 0.0])
    for index, allocation in enumerate(allocations, start=1):
        allocation["id"] = "A{}".format(index)
        allocation["order_qty"] = None
        if allocation["po_qty"] not in (None, ""):
            allocation["order_qty"] = float(number_value(allocation["po_qty"], None, allocation["rows"][0],
                                                         "PO QTY"))
        allocation["totals"] = [clean_number(value) for value in allocation["totals"]]
        allocation["ambiguous"] = allocation["signature"] in ambiguous_signatures
        allocations_by_source_signature[allocation["signature"]].append(allocation)
        key = allocation["signature"][:3] + (allocation["delivery"],)
        for total_index, total in enumerate(allocation["totals"]):
            aggregate[key][total_index] += total
        if allocation["ambiguous"]:
            ambiguous.append({
                "allocationId": allocation["id"],
                "style": allocation["style"],
                "item": allocation["item"],
                "po": allocation["po"],
                "delivery": allocation["delivery"],
                "country": allocation["country"],
                "colour": allocation["colour"],
                "reason": "The identical Style/Item/PO/Delivery/Country/Colour allocation signature is repeated in non-contiguous source sections; split cannot be proven."
            })
    return allocations, ambiguous, aggregate


def parse_source(path):
    allocations, _ambiguous, aggregate = parse_source_allocations(path)
    aggregated = defaultdict(lambda: [0.0, 0.0, 0.0])
    for key, totals in aggregate.items():
        legacy_key = (key[0], key[1], key[3])
        aggregated[legacy_key] = totals
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


def summary_destinations(workbook):
    destinations = []
    for ws in workbook.worksheets:
        header_row, cols = find_summary_header(ws)
        if header_row is None:
            continue
        style_col = next((c for c in range(1, ws.max_column + 1)
                          if norm(ws.cell(header_row, c).value) == "style"), None)
        item_col = next((c for c in range(1, ws.max_column + 1)
                         if norm(ws.cell(header_row, c).value) in ("item", "item no")), None)
        for row in range(header_row + 1, ws.max_row + 1):
            qty_raw = ws.cell(row=row, column=cols["qty"]).value
            if qty_raw in (None, "") and not text(ws.cell(row=row, column=cols["xfty"]).value):
                continue
            if not parse_date(ws.cell(row=row, column=cols["xfty"]).value):
                continue
            destinations.append({
                "sheet": ws.title,
                "row": row,
                "style": text(ws.cell(row=row, column=style_col).value) if style_col else "",
                "item": text(ws.cell(row=row, column=item_col).value) if item_col else "",
                "xfty": date_key(ws.cell(row=row, column=cols["xfty"]).value),
                "qty": qty_raw,
            })
    return destinations


def allocation_review(workbook, report_path):
    allocations, ambiguous, _aggregate = parse_source_allocations(report_path)
    destinations = summary_destinations(workbook)
    candidate_allocations = []
    for allocation in allocations:
        candidate_allocations.append({
            "id": allocation["id"],
            "style": allocation["style"],
            "item": allocation["item"],
            "po": allocation["po"],
            "deliveryDate": allocation["delivery"],
            "country": allocation["country"],
            "colour": allocation["colour"],
            "orderQty": allocation["order_qty"],
            "productionTotals": allocation["totals"],
            "ambiguous": allocation["ambiguous"],
        })

    rows = []
    unresolved = []
    for destination in destinations:
        candidates = [
            allocation for allocation in candidate_allocations
            if norm(allocation["style"]) == norm(destination["style"])
            and norm(allocation["item"]) == norm(destination["item"])
        ]
        suggested = []
        if candidates and destination["qty"] not in (None, ""):
            try:
                destination_qty = float(number_value(destination["qty"], None, destination["row"], "Qty"))
            except ValueError:
                destination_qty = None
            if destination_qty is not None:
                suggested = [allocation["id"] for allocation in candidates
                             if allocation["orderQty"] is not None
                             and abs(float(allocation["orderQty"]) - destination_qty) < 0.0001]
        rows.append(dict(destination, candidateAllocationIds=[allocation["id"] for allocation in candidates],
                         suggestedAllocationIds=suggested, confirmedAllocationIds=[]))
        if not candidates:
            unresolved.append({
                "sheet": destination["sheet"], "row": destination["row"],
                "reason": "No source allocation matches the destination Style and Item."
            })
        elif len(candidates) > 1 and not suggested:
            unresolved.append({
                "sheet": destination["sheet"], "row": destination["row"],
                "reason": "Multiple source allocations require explicit shipment selection."
            })
    return {
        "review": True,
        "destinations": rows,
        "candidateAllocations": candidate_allocations,
        "ambiguousAllocations": ambiguous,
        "unresolvedMappings": unresolved,
        "destinationIdentifiers": [
            {"sheet": row["sheet"], "row": row["row"], "style": row["style"],
             "item": row["item"], "xfty": row["xfty"], "qty": text(row["qty"])}
            for row in rows
        ],
    }


def write_summary(workbook, allocations, mapping):
    updated_rows = 0
    completion_cells = 0
    updated_sheets = set()
    unmatched = []
    ambiguous = []
    quantity_conflicts = []
    allocations_by_id = {allocation["id"]: allocation for allocation in allocations}
    mappings_by_destination = {
        "{}\u0000{}".format(item["sheet"], item["row"]): item.get("allocationIds", [])
        for item in mapping
    }
    used_allocations = defaultdict(list)
    allocation_mappings = defaultdict(list)
    for ws in workbook.worksheets:
        header_row, cols = find_summary_header(ws)
        if header_row is None:
            continue

        summary_style_col = next((c for c in range(1, ws.max_column + 1)
                                  if norm(ws.cell(header_row, c).value) == "style"), None)
        for row in range(header_row + 1, ws.max_row + 1):
            xfty = ws.cell(row=row, column=cols["xfty"]).value
            delivery = date_key(xfty)
            if not delivery:
                continue
            summary_style = norm(ws.cell(row=row, column=summary_style_col).value) if summary_style_col else ""
            summary_item_col = next((c for c in range(1, ws.max_column + 1)
                                     if norm(ws.cell(header_row, c).value) in ("item", "item no")), None)
            summary_item = norm(ws.cell(row=row, column=summary_item_col).value) if summary_item_col else ""
            selected_ids = mappings_by_destination.get("{}\u0000{}".format(ws.title, row), [])
            if not selected_ids:
                unmatched.append({"sheet": ws.title, "row": row, "style": summary_style,
                                  "item": summary_item, "xfty": delivery,
                                  "reason": "No confirmed source allocation mapping."})
                continue
            selected_allocations = [allocations_by_id.get(allocation_id) for allocation_id in selected_ids]
            if any(allocation is None for allocation in selected_allocations):
                unmatched.append({"sheet": ws.title, "row": row, "reason": "Mapping contains an unknown allocation."})
                continue
            for allocation in selected_allocations:
                used_allocations[allocation["id"]].append((ws.title, row))
                allocation_mappings[allocation["id"]].append((ws.title, row))
            repeated = [allocation_id for allocation_id, destinations in allocation_mappings.items()
                        if allocation_id in selected_ids and len(destinations) > 1]
            mismatched = [allocation for allocation in selected_allocations
                          if norm(allocation["style"]) != summary_style
                          or norm(allocation["item"]) != summary_item]
            if repeated or mismatched:
                ambiguous.append({
                    "sheet": ws.title, "row": row, "allocationIds": selected_ids,
                    "reason": "A source allocation is duplicated or does not match the destination Style/Item."
                })
                continue
            actual = [sum(float(a["totals"][index]) for a in selected_allocations) for index in range(3)]

            qty_cell = "{}{}".format(get_column_letter(cols["qty"]), row)
            try:
                destination_qty = float(number_value(ws.cell(row=row, column=cols["qty"]).value, None, row, "Qty"))
            except ValueError:
                destination_qty = None
            mapped_order_qty = sum(float(a["order_qty"]) for a in selected_allocations if a["order_qty"] is not None)
            if destination_qty is None or abs(mapped_order_qty - destination_qty) > 0.0001:
                quantity_conflicts.append({
                    "sheet": ws.title, "row": row, "style": summary_style, "item": summary_item,
                    "destinationQty": destination_qty, "mappedOrderQty": mapped_order_qty,
                    "allocationIds": selected_ids,
                })
            process_columns = (
                ("sewing", cols["sewing"], cols["sewing_balance"]),
                ("wash", cols["wash"], cols["wash_balance"]),
                ("pack", cols["pack"], cols["pack_balance"]),
            )
            for index, (name, completed_col, balance_col) in enumerate(process_columns):
                ws.cell(row=row, column=completed_col, value=clean_number(actual[index]))
                completion_cells += 1
                completed_cell = "{}{}".format(get_column_letter(completed_col), row)
                balance_cell = ws.cell(row=row, column=balance_col)
                expected = "={}{}-{}{}".format(completed_cell, "", get_column_letter(cols["qty"]), row)
                expected = "={}-{}{}".format(completed_cell, get_column_letter(cols["qty"]), row)
                if balance_cell.data_type == "f" and balance_cell.value == expected:
                    continue
                if balance_cell.value is None or isinstance(balance_cell.value, (int, float)) or \
                        balance_cell.data_type == "f":
                    balance_cell.value = expected
            updated_rows += 1
            updated_sheets.add(ws.title)

    for allocation_id, destinations in used_allocations.items():
        if len(destinations) > 1:
            ambiguous.append({
                "allocationId": allocation_id,
                "destinations": destinations,
                "reason": "A source allocation was mapped to more than one destination row."
            })
    return {
        "summaryRowsUpdated": updated_rows,
        "completionCellsWritten": completion_cells,
        "styleSheetsUpdated": sorted(updated_sheets),
        "unmatchedGroups": unmatched,
        "ambiguousGroups": ambiguous,
        "quantityConflicts": quantity_conflicts,
    }


def clean_number(value):
    return int(value) if float(value).is_integer() else value


def main():
    if len(sys.argv) < 4:
        fail("Usage: fill_unit23_summary.py <workbook> <summary report> <output> [--review | --mapping <json>]")
    workbook_path, report_path, output_path = sys.argv[1:4]
    try:
        workbook = openpyxl.load_workbook(workbook_path, data_only=False)
    except Exception as exc:
        fail("Could not open destination workbook: {}".format(exc))

    try:
        if len(sys.argv) >= 5 and sys.argv[4] == "--review":
            result = allocation_review(workbook, report_path)
            print(json.dumps(dict({"success": True, "process": "summaryReview"}, **result)))
            return
        allocations, ambiguous, _aggregate = parse_source_allocations(report_path)
    except (ValueError, OSError) as exc:
        fail(str(exc))

    mapping = []
    if len(sys.argv) >= 6 and sys.argv[4] == "--mapping":
        try:
            with open(sys.argv[5], "r", encoding="utf-8") as handle:
                mapping_payload = json.load(handle)
            mapping = mapping_payload.get("mappings", [])
        except (OSError, ValueError, AttributeError) as exc:
            fail("Could not read confirmed summary mapping: {}".format(exc))
    else:
        fail("Confirmed PO-to-summary mappings are required.")

    summary = write_summary(workbook, allocations, mapping)
    summary["unresolvedMappings"] = summary["unmatchedGroups"]
    summary["ambiguousAllocations"] = ambiguous
    summary["success"] = bool(allocations)
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
