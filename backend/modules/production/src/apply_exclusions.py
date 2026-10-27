#!/usr/bin/env python3
"""
apply_exclusions.py -- remove selected dates from a generated Plan vs Actuals workbook.

Usage:
    python3 apply_exclusions.py <generated_workbook.xlsx> <output.xlsx> <excluded_dates_json>

The excluded_dates_json argument is a JSON array of ISO dates, for example:
    ["2026-09-05", "2026-09-12"]
"""
import sys
import json
import datetime

import openpyxl


def fail(message):
    print(json.dumps({"error": message}), file=sys.stderr)
    sys.exit(1)


def parse_sheet_date(value):
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, str):
        try:
            return datetime.date.fromisoformat(value[:10])
        except ValueError:
            return None
    return None


def main():
    if len(sys.argv) != 4:
        fail("Usage: apply_exclusions.py <generated_workbook.xlsx> <output.xlsx> <excluded_dates_json>")

    workbook_path, out_path, raw_dates = sys.argv[1], sys.argv[2], sys.argv[3]

    try:
        excluded = {datetime.date.fromisoformat(d) for d in json.loads(raw_dates)}
    except Exception as e:
        fail(f"Invalid excluded dates: {e}")

    try:
        wb = openpyxl.load_workbook(workbook_path)
    except Exception as e:
        fail(f"Could not open the generated workbook: {e}")

    rows_removed = 0
    for ws in wb.worksheets:
        rows_to_delete = []
        for row in range(4, ws.max_row + 1):
            row_date = parse_sheet_date(ws.cell(row=row, column=3).value)
            if row_date in excluded:
                rows_to_delete.append(row)

        for row in reversed(rows_to_delete):
            ws.delete_rows(row, 1)
            rows_removed += 1

        last_data_row = ws.max_row
        while last_data_row >= 4 and ws.cell(row=last_data_row, column=3).value is None:
            last_data_row -= 1

        working_rows = [r for r in range(4, last_data_row + 1) if ws.cell(row=r, column=3).value is not None]
        sew_plans = [ws.cell(row=r, column=4).value for r in working_rows]

        for i, row in enumerate(working_rows):
            wash_plan = sew_plans[i - 1] if i > 0 else None
            pack_plan = sew_plans[i - 2] if i > 1 else None

            ws.cell(row=row, column=9, value=wash_plan)
            ws.cell(row=row, column=14, value=pack_plan)

            ws.cell(row=row, column=5, value=f'=IF(ISNUMBER(F{row}),SUM($D$4:D{row}),"")')
            ws.cell(row=row, column=7, value=f'=IF(ISNUMBER(F{row}),SUM($F$4:F{row}),"")')
            ws.cell(row=row, column=8, value=f'=IF(ISNUMBER(G{row}),N(E{row})-G{row},"")')
            ws.cell(row=row, column=10, value=f'=IF(ISNUMBER(K{row}),SUM($I$4:I{row}),"")')
            ws.cell(row=row, column=12, value=f'=IF(ISNUMBER(K{row}),SUM($K$4:K{row}),"")')
            ws.cell(row=row, column=13, value=f'=IF(ISNUMBER(L{row}),N(J{row})-L{row},"")')
            ws.cell(row=row, column=15, value=f'=IF(ISNUMBER(P{row}),SUM($N$4:N{row}),"")')
            ws.cell(row=row, column=17, value=f'=IF(ISNUMBER(P{row}),SUM($P$4:P{row}),"")')
            ws.cell(row=row, column=18, value=f'=IF(ISNUMBER(Q{row}),N(O{row})-Q{row},"")')

        # NOTE: Order Summary columns (T to AB) are intentionally NOT touched here.
        # That section (Style, X-fty, Qty, Sewing completed, Bal to sew, Wash
        # completed, Bal to Wash, Pack completed, Bal to Pack) is maintained
        # manually by the user in the uploaded Existing Workbook and must be
        # returned exactly as uploaded.

        # Update conditional formatting rules range to H4:H{last_data_row}, M4:M{last_data_row}, R4:R{last_data_row}
        ws.conditional_formatting = openpyxl.formatting.formatting.ConditionalFormattingList()
        FONT = 'Arial'
        from openpyxl.styles import Font, PatternFill
        from openpyxl.formatting.rule import FormulaRule

        deviation_positive_fill = PatternFill('solid', fgColor='C6EFCE')
        deviation_positive_font = Font(name=FONT, size=10, color='006100')
        deviation_negative_fill = PatternFill('solid', fgColor='FFC7CE')
        deviation_negative_font = Font(name=FONT, size=10, color='9C0006')

        for column in ('H', 'M', 'R'):
            deviation_range = f'{column}4:{column}{last_data_row}'
            ws.conditional_formatting.add(
                deviation_range,
                FormulaRule(
                    formula=[f'{column}4>0'],
                    fill=deviation_positive_fill,
                    font=deviation_positive_font,
                ),
            )
            ws.conditional_formatting.add(
                deviation_range,
                FormulaRule(
                    formula=[f'{column}4<0'],
                    fill=deviation_negative_fill,
                    font=deviation_negative_font,
                ),
            )

    try:
        wb.save(out_path)
    except Exception as e:
        fail(f"Could not save output file: {e}")

    print(json.dumps({
        "success": True,
        "excludedDates": sorted(d.isoformat() for d in excluded),
        "rowsRemoved": rows_removed,
    }))


if __name__ == "__main__":
    main()