import unittest
from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch

from openpyxl import load_workbook

from models.table import TableData
from services.recognition_job import run_recognition
from services.table_layout import TextBox, reconstruct_table


class PipelineTests(unittest.TestCase):
    def test_multiline_cells_and_leading_zero(self):
        boxes = [TextBox(0, 10, 30, 20, "PO"), TextBox(100, 0, 130, 8, "Order"),
                 TextBox(100, 10, 130, 20, "No."), TextBox(200, 10, 230, 20, "Qty"),
                 TextBox(0, 50, 40, 60, "000123"), TextBox(100, 35, 150, 45, "ABC-"),
                 TextBox(100, 50, 130, 60, "001"), TextBox(210, 50, 230, 60, "20")]
        table = reconstruct_table(boxes)
        self.assertEqual(table.columns, ["PO", "Order No.", "Qty"])
        self.assertEqual(table.rows, [["000123", "ABC-001", "20"]])

    def test_blank_image_has_no_table(self):
        with self.assertRaises(ValueError):
            reconstruct_table([])

    def test_automatic_export_preserves_text_and_uses_unique_paths(self):
        table = TableData(["PO", "Qty"], [["000123", "20"], ["=1+1", "2"]])
        service = Mock()
        service.recognize.return_value = table
        with TemporaryDirectory() as temporary:
            saved = []
            for _ in range(2):
                events = Queue()
                run_recognition(service, Path("source.png"), Path(temporary), events)
                messages = list(events.queue)
                self.assertEqual([name for name, _ in messages], ["table", "saved", "done"])
                saved.append(messages[1][1])
            self.assertNotEqual(*saved)
            book = load_workbook(saved[0])
            try:
                self.assertEqual(book.active["A2"].value, "000123")
                self.assertEqual(book.active["A3"].data_type, "s")
            finally:
                book.close()

    def test_save_failure_keeps_recognized_table(self):
        service = Mock()
        service.recognize.return_value = TableData(["A", "B"], [["1", "2"]])
        events = Queue()
        with patch("services.recognition_job.write_table", side_effect=PermissionError("locked")):
            run_recognition(service, Path("source.png"), Path("output"), events)
        self.assertEqual([name for name, _ in events.queue], ["table", "save_error", "done"])

    def test_recognition_failure_does_not_export(self):
        service = Mock()
        service.recognize.side_effect = ValueError("no text")
        events = Queue()
        with patch("services.recognition_job.write_table") as writer:
            run_recognition(service, Path("source.png"), Path("output"), events)
            writer.assert_not_called()
        self.assertEqual([name for name, _ in events.queue], ["error", "done"])


if __name__ == "__main__":
    unittest.main()
