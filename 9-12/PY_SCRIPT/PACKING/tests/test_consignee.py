import unittest

from pypdf import PdfWriter
from pypdf.generic import DictionaryObject, NameObject, DecodedStreamObject

from models.packing import PackingDataError
from services.consignee_service import first_consignee_line


def page_with_text(lines):
    writer = PdfWriter()
    page = writer.add_blank_page(width=612, height=792)
    font = DictionaryObject({NameObject('/Type'): NameObject('/Font'),
                             NameObject('/Subtype'): NameObject('/Type1'),
                             NameObject('/BaseFont'): NameObject('/Helvetica')})
    page[NameObject('/Resources')] = DictionaryObject({NameObject('/Font'): DictionaryObject({
        NameObject('/F1'): font})})
    stream = DecodedStreamObject()
    stream.set_data('\n'.join(f'BT /F1 10 Tf 1 0 0 1 {x} {y} Tm ({text}) Tj ET'
                              for x, y, text in lines).encode('ascii'))
    page[NameObject('/Contents')] = stream
    return page


class ConsigneeTests(unittest.TestCase):
    def test_same_line_company_before_label_in_stream_excludes_address_and_buyer(self):
        page = page_with_text([(40, 650, 'Different Buyer'),
                               (430, 650, 'Example Retail Inc.'),
                               (430, 635, '123 Street US'),
                               (290, 650, 'Ultimate Consignee:')])
        self.assertEqual(first_consignee_line([page]), 'Example Retail Inc.')

    def test_inline_and_below_heading(self):
        inline = page_with_text([(290, 650, 'Ultimate Consignee: Another Customer'),
                                 (290, 635, '123 Street')])
        below = page_with_text([(290, 650, 'Ultimate Consignee'),
                                (290, 635, 'Another Customer'),
                                (290, 620, '123 Street')])
        self.assertEqual(first_consignee_line([inline, below]), 'Another Customer')

    def test_missing_and_conflicting_names_fail_instead_of_guessing(self):
        with self.assertRaises(PackingDataError):
            first_consignee_line([page_with_text([(20, 600, 'Customer Name: Other')])])
        with self.assertRaises(PackingDataError):
            first_consignee_line([page_with_text([(20, 600, 'Ultimate Consignee: Customer A')]),
                                  page_with_text([(20, 600, 'Ultimate Consignee: Customer B')])])
