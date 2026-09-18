from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import Mock, patch
import unittest

from models.packing import PackingDataError
from services.order_pdf_service import read_order_items
from services.packing_generation_service import generate_packing_lists


class ZeroQuantityTests(unittest.TestCase):
    order = 'POHK-26-12345-001'

    def read(self, lines):
        text = f'PURCHASE ORDER NO.: {self.order}\nCUSTOMER NAME: Customer\n' + lines
        with patch('services.order_pdf_service.PdfReader') as reader:
            reader.return_value.pages = [Mock(extract_text=lambda: text)]
            return read_order_items(Path('order.pdf'), self.order)

    def test_zero_item_skipped_before_missing_packing_fields(self):
        items, customer = self.read('28483-000-1B-001-OPS\nBall 0 PCS\n'
            '33465-000-6B-012-HDQ\nDuck 120 PCS\n12 PCS / BOX Packing:\n123 Customer Order No.:\n')
        self.assertEqual(customer, 'Customer')
        self.assertEqual([item.item_no for item in items], ['33465-000-6B-012-HDQ'])
        self.assertEqual(items[0].quantity, 120)

    def test_nonzero_invalid_case_pack_still_rejected(self):
        with self.assertRaises(PackingDataError):
            self.read('33465-000-6B-012-HDQ\nDuck 120 PCS\n0 PCS / BOX Packing:\n123 Customer Order No.:\n')

    def test_all_zero_order_skips_without_creating_workbook(self):
        items, customer = self.read('28483-000-1B-001-OPS\nBall 0.00 PCS\n')
        self.assertEqual(items, [])
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            with patch('services.packing_generation_service.read_order_items', return_value=(items, customer)), \
                    patch('services.packing_generation_service.PackingLookup') as lookup, \
                    patch('services.packing_generation_service.build_packing_excel') as build:
                result = generate_packing_lists({self.order: [root / 'order.pdf']}, None, None, None,
                                                root, None, lambda *args: None)
            self.assertEqual(result, ([], [], [self.order]))
            lookup.assert_not_called()
            build.assert_not_called()
            self.assertEqual(list(root.iterdir()), [])
