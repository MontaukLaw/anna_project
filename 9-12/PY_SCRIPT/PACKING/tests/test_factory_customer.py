from pathlib import Path
import unittest
from unittest.mock import Mock, patch

from models.packing import PackingDataError
from services.order_pdf_service import factory_customer_name, read_order_items


class FactoryCustomerTests(unittest.TestCase):
    def test_ultimate_consignee_layout_when_customer_name_is_absent(self):
        from test_consignee import page_with_text
        for lines in (
            [(290, 650, 'Ultimate Consignee'), (290, 635, 'Customer B'), (290, 620, '123 Street')],
            [(290, 650, 'Ultimate Consignee: Customer B'), (290, 635, '123 Street')],
            [(40, 650, 'Different Buyer'), (430, 650, 'Customer B'),
             (430, 635, '123 Street'), (290, 650, 'Ultimate Consignee:')],
        ):
            page = page_with_text(lines)
            with self.subTest(lines=lines):
                self.assertEqual(factory_customer_name(page.extract_text(), 'order.pdf', [page]), 'Customer B')

    def test_customer_name_takes_priority_and_missing_both_names_identifies_file(self):
        from test_consignee import page_with_text
        page = page_with_text([(290, 650, 'Ultimate Consignee: Different Customer')])
        self.assertEqual(factory_customer_name('CUSTOMER NAME\nCustomer A\n' + page.extract_text(),
                                               'order.pdf', [page]), 'Customer A')
        empty = page_with_text([(40, 650, 'Other fields')])
        with self.assertRaisesRegex(PackingDataError, 'order.pdf.*CUSTOMER NAME.*Ultimate Consignee'):
            factory_customer_name(empty.extract_text(), 'order.pdf', [empty])

    def test_first_line_after_heading_and_inline_extraction(self):
        for header in ('CUSTOMER NAME\nCustomer A', 'CUSTOMER NAME :\nCustomer A',
                       'CUSTOMER NAME：\n\nCustomer A', 'customer name: Customer A'):
            text = header + '\n123 Street\nUltimate Consignee: Different Company\n'
            with self.subTest(header=header):
                self.assertEqual(factory_customer_name(text, 'order.pdf'), 'Customer A')

    def test_repeated_headers_missing_names_and_conflicts(self):
        same = 'CUSTOMER NAME\nCustomer A\n'
        self.assertEqual(factory_customer_name(same + same, 'order.pdf'), 'Customer A')
        for text in ('Ultimate Consignee: Wal-Mart Store Inc.', 'CUSTOMER NAME\n',
                     'CUSTOMER NAME\nCUSTOMER ORDER NO: 123',
                     same + 'CUSTOMER NAME\nCustomer B'):
            with self.subTest(text=text), self.assertRaisesRegex(PackingDataError, 'order.pdf.*CUSTOMER NAME'):
                factory_customer_name(text, 'order.pdf')

    def test_factory_parser_uses_customer_name_without_ultimate_consignee(self):
        text = ('PURCHASE ORDER NO.: POHK-26-12345-001\n'
                'CUSTOMER NAME\nCustomer From Order\n123 Street\n'
                '28483-000-1B-001-OPS\nToy 8 PCS\n4 PCS / BOX Packing:\n'
                '000123 Customer Order No.:\n')
        reader = Mock(pages=[Mock(extract_text=lambda: text)])
        with patch('services.order_pdf_service.PdfReader', return_value=reader), \
             patch('services.consignee_service.first_consignee_line', side_effect=AssertionError('Must not read consignee')):
            items, customer = read_order_items(Path('order.pdf'), 'POHK-26-12345-001', customer_from_name=True)
        self.assertEqual(customer, 'Customer From Order')
        self.assertEqual(len(items), 1)
        self.assertEqual(items[0].quantity, 8)
