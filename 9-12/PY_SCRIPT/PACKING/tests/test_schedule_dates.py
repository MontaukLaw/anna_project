from datetime import date, datetime
import unittest
from openpyxl.utils.datetime import to_excel, MAC_EPOCH
from services.date_service import parse_schedule_date
from models.packing import PackingDataError


class ScheduleDateTests(unittest.TestCase):
    def test_excel_dates_and_text(self):
        expected=date(2026,9,17)
        for value in [expected,datetime(2026,9,17),'2026-09-17','2026/9/17','2026.9.17','2026年9月17日',to_excel(expected)]:
            self.assertEqual(parse_schedule_date(value),expected)
        self.assertEqual(parse_schedule_date(to_excel(expected,MAC_EPOCH),MAC_EPOCH),expected)

    def test_missing_and_ambiguous_dates_fail(self):
        for value in [None,'待定','9/17','2026/2/30']:
            with self.assertRaises(PackingDataError):
                parse_schedule_date(value)
