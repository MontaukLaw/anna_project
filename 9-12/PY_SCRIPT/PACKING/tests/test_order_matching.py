from pathlib import Path
import unittest

from models.table import TableData
from services.order_matching_service import match_order_pdfs


class OrderMatchingTests(unittest.TestCase):
    def test_purchase_order_priority_duplicates_and_missing(self):
        table = TableData(["Customer PO No.", "Purchase Order No."], [
            ["1007875885", "POHK-26-14888-001"],
            ["1007875885", "POHK-26-14888-001"],
            ["000123", "POHK-26-99999-001"],
            ["000456", ""],
        ])
        files = [Path("COUTURE_POHK-26-14888-001_WM-US.pdf"),
                 Path("sub/POHK-26-14888-001_copy.PDF"),
                 Path("1007875885.pdf"), Path("POHK-26-14888-0010.pdf")]
        result = match_order_pdfs(table, files)
        self.assertEqual(result.column, "Purchase Order No.")
        self.assertEqual(result.files_by_order["POHK-26-14888-001"], files[:2])
        self.assertEqual(result.files_by_order["POHK-26-99999-001"], [])
        self.assertEqual(result.blank_rows, 1)
        self.assertEqual(len(result.files_by_order), 2)

    def test_normalizes_wrapping_case_and_hyphens(self):
        order = "POHK–26-14888-\n001"
        result = match_order_pdfs(TableData(["订单号"], [[order]]), [Path("couture_pohk-26-14888-001.pdf")])
        self.assertEqual(len(result.files_by_order[order]), 1)

    def test_preserves_leading_zero_and_empty_directory(self):
        table = TableData(["订单号"], [["000123"]])
        self.assertEqual(match_order_pdfs(table, [Path("order_123.pdf")]).files_by_order, {"000123": []})
        self.assertEqual(match_order_pdfs(table, []).files_by_order, {"000123": []})

    def test_unknown_column_is_reported(self):
        with self.assertRaises(ValueError):
            match_order_pdfs(TableData(["Item No."], [["123"]]), [])


if __name__ == "__main__":
    unittest.main()
