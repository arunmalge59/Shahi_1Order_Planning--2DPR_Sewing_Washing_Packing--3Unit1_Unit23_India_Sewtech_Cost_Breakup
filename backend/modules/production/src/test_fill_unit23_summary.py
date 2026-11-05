import importlib.util
import os
import tempfile
import unittest

import openpyxl
from openpyxl.utils import get_column_letter


MODULE_PATH = os.path.join(os.path.dirname(__file__), "fill_unit23_summary.py")
SPEC = importlib.util.spec_from_file_location("fill_unit23_summary", MODULE_PATH)
SUMMARY = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(SUMMARY)


def header(ws, row, finishing_col=26):
    headers = {
        "A": "STYLE NO",
        "B": "ITEM NO",
        "F": "PO",
        "G": "PO QTY",
        "H": "COUNTRY",
        "I": "DELIVERY DATE",
        "J": "PRODUCT TYPE",
        "R": "SEWING OUT",
        "V": "WASHING OUT",
    }
    parent_row = row - 1
    for col, value in headers.items():
        ws["{}{}".format(col, parent_row)] = value
    ws["{}{}".format(get_column_letter(finishing_col), parent_row)] = "FINISHING OUT"
    for col, value in headers.items():
        ws["{}{}".format(col, row)] = value if col in ("A", "B", "F", "G", "H", "I", "J") else "TILL DATE"
    ws["{}{}".format(get_column_letter(finishing_col), row)] = "TILL DATE"


def detail(ws, row, style="", item="", po="", delivery="", sewing="", washing="", finishing="", finishing_col=26):
    values = {
        "A": style,
        "B": item,
        "F": po,
        "I": delivery,
        "R": sewing,
        "V": washing,
    }
    values[get_column_letter(finishing_col)] = finishing
    for col, value in values.items():
        ws["{}{}".format(col, row)] = value


class Unit23SummaryParserTests(unittest.TestCase):
    def make_report(self, mutate=None):
        temporary = tempfile.TemporaryDirectory()
        path = os.path.join(temporary.name, "source.xlsx")
        workbook = openpyxl.Workbook()
        ws = workbook.active

        ws["A1"] = "BS30 - F2625M501B (ADIDAS)"
        header(ws, 3, 26)
        detail(ws, 4, "F2625M501B", "BS30", "0902223859", "2026-05-07", 975, 968, 889)
        detail(ws, 5, finishing="", finishing_col=26)
        ws["Z5"] = 439
        ws["R5"] = 468
        ws["V5"] = 459

        ws["A7"] = "** Red Colour (Rejected Pcs) and Orange Colour (Sample Pcs)"
        ws["A8"] = "GRAND TOTAL"
        ws["R8"] = 99999

        ws["A10"] = "BS31 - F2625M502B (ADIDAS)"
        header(ws, 12, 28)
        detail(ws, 13, "F2625M502B", "BS31", "0902223860", "2026-05-08", 10, 20, 30, 28)
        if mutate:
            mutate(ws)
        workbook.save(path)
        return temporary, path

    def test_multiple_sections_and_consecutive_continuations(self):
        temporary, path = self.make_report()
        self.addCleanup(temporary.cleanup)
        aggregates = SUMMARY.parse_source(path)
        key1 = ("f2625m501b", "bs30", "2026-05-07")
        key2 = ("f2625m502b", "bs31", "2026-05-08")
        self.assertEqual(aggregates[key1], [1443.0, 1427.0, 1328.0])
        self.assertEqual(aggregates[key2], [10.0, 20.0, 30.0])

    def test_shifted_finishing_column_is_rebuilt(self):
        temporary = tempfile.TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        path = os.path.join(temporary.name, "source.xlsx")
        workbook = openpyxl.Workbook()
        ws = workbook.active
        ws["A1"] = "BS40 - STYLE40 (ADIDAS)"
        header(ws, 3, 28)
        detail(ws, 4, "STYLE40", "BS40", "SAMPLE-1", "2026-06-01", 3, 4, 5, 28)
        workbook.save(path)
        aggregates = SUMMARY.parse_source(path)
        self.assertEqual(aggregates[("style40", "bs40", "2026-06-01")], [3.0, 4.0, 5.0])

    def test_invalid_detail_quantity_is_rejected(self):
        def mutate(ws):
            ws["AB14"] = "NOT A NUMBER"
            ws["R14"] = 1
            ws["V14"] = 1

        temporary, path = self.make_report(mutate)
        self.addCleanup(temporary.cleanup)
        # Fix the mutation to use the shifted finishing column in section two.
        with self.assertRaises(ValueError):
            SUMMARY.parse_source(path)


if __name__ == "__main__":
    unittest.main()
