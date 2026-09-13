import unittest
from models.table import TableData
from services.pdf_remarks_service import extract_remarks, append_remarks, filter_packaging_remarks


class RemarksTests(unittest.TestCase):
    def test_extract_remarks_excludes_signatures_and_preserves_lines(self):
        result=extract_remarks(['Header\nREMARKS:\nNeed RFID.\nShip loose carton.\nCouture Toy LimitedJust Play (HK) Ltd\nAccepted By'])
        self.assertEqual(result,'Need RFID.\nShip loose carton.')

    def test_continuation_and_missing_section(self):
        self.assertEqual(extract_remarks(['REMARKS:\nFirst rule.',
            'PURCHASE ORDER\nITEM NO. DESCRIPTION QTY. UNIT AMOUNT\nPRICE (USD)\nSecond rule.\nAccepted By']),
            'First rule.\nSecond rule.')
        self.assertEqual(extract_remarks(['No remarks here']), '')

    def test_append_matches_order_and_replaces_old_column(self):
        table=TableData(['Purchase Order No.','REMARKS'],[['POHK-26-12345-001','old'],['POHK-26-54321-001','old']])
        result=append_remarks(table,'Purchase Order No.',{'pohk-26-12345-001':'New remarks'})
        self.assertEqual(result.columns,['Purchase Order No.','REMARKS'])
        self.assertEqual(result.rows[0][-1],'New remarks')
        self.assertNotEqual(result.rows[1][-1],'old')
        self.assertEqual(table.rows[0][-1],'old')

    def test_packaging_filter_only_requested_instructions(self):
        text='Need RFID.\n1,800 pcs = 180 cnts =\n1 slip sheet.\nSHIP loose cartons.\nOther notes.'
        self.assertEqual(filter_packaging_remarks(text),'1,800 pcs = 180 cnts = 1 slip sheet\n散箱出货/不打托')
        self.assertEqual(filter_packaging_remarks('720 PCS = 90 CTNS = 1 SLIP'),'720 PCS = 90 CTNS = 1 SLIP')
        self.assertEqual(filter_packaging_remarks('Do not ship loose cartons. Need RFID.'),'')

    def test_empty_remarks_hidden_and_unmatched_rows_blank(self):
        table=TableData(['Purchase Order No.','REMARKS'],[['POHK-26-12345-001','old'],['POHK-26-54321-001','old']])
        result=append_remarks(table,'Purchase Order No.',{})
        self.assertEqual(result.columns,['Purchase Order No.'])
        result=append_remarks(table,'Purchase Order No.',{'pohk-26-12345-001':'散箱出货/不打托'})
        self.assertEqual(result.rows[1][-1],'')
