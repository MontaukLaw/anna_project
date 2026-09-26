from copy import deepcopy
from datetime import date
from decimal import Decimal as D
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

import xlrd

from config import ASN_TEMPLATE
from models.customs_packing import CustomsPackingData, CustomsPackingRow
from services.asn_service import AuditReport, Document, Item, audit_directory
from services.customs_packing_service import read_customs_packing
from services.asn_generation_service import (build_asn_plan, generation_options, product_key,
                                             write_asn_workbook, output_filename)


def fixture():
    item = Item('填充玩具/测试商品', D(8), D('7.5'), D(60), 'test', code='9503008390',
                cartons=D(2), country='美国', net_weight=D('5.4'), gross_weight=D('6.88'))
    sale = deepcopy(item)
    sale.specification = '3|0|测试规格'
    invoice = deepcopy(item)
    invoice.unit = '套'
    docs = {kind: Document(kind, Path('test.xls'), 'TEST001', signed_date='2026-05-10', items=[entry])
            for kind, entry in [('装箱单', item), ('香港合同', sale), ('形式发票', invoice)]}
    report = AuditReport(documents={'TEST001': docs}, checks=1)
    values = {'customer': 'Just Play LLC', 'item': '00123', 'chinese_name': '测试商品',
              'customer_order': '000123-PO', 'so': 'SO001', 'quantity': D(4), 'cartons': D(1),
              'per_carton': D(4), 'price': D('7.5'), 'length': D(40), 'width': D(30),
              'height': D(20), 'volume': D('.024'), 'net_weight': D('2.7'), 'gross_weight': D('3.44')}
    rows = [CustomsPackingRow('Sheet1', r, [], deepcopy(values), {}) for r in (4, 5)]
    rows[1].values['so'] = 'SO002'
    return report, CustomsPackingData(Path('PL.xlsx'), rows)


class AsnGenerationTests(unittest.TestCase):
    def setUp(self):
        self.report, self.packing = fixture()

    def plan(self):
        return build_asn_plan(self.report, self.packing, ASN_TEMPLATE, generation_options(self.report, self.packing))

    def test_confirmed_rules_and_original_row_order(self):
        plan = self.plan()
        self.assertEqual(plan['scope'], 'asn_only')
        self.assertEqual(plan['headers']['C5'], date.today().isoformat())
        self.assertEqual(plan['headers']['G13'], '2026-05-10')
        self.assertEqual(len(plan['rows']), 2)
        self.assertEqual([r[1] for r in plan['rows']], ['SO001', 'SO002'])
        row = plan['rows'][0]
        self.assertEqual([row[i] for i in (12, 27, 32)], ['套'] * 3)
        self.assertEqual(row[4], '000123-PO')
        self.assertEqual(row[5], '00123')
        self.assertEqual(row[33:35], [2.7, '千克'])
        self.assertEqual(row[43:46], ['中国', '达州市', '51169'])
        self.assertIsNone(row[20])
        self.assertEqual(plan['formulas'][1], {'cell': 'AD27', 'value': '=ROUND(AA27*AC27,2)'})

    def test_invoice_unit_required_and_unique(self):
        invoice = self.report.documents['TEST001']['形式发票']
        invoice.items[0].unit = ''
        with self.assertRaisesRegex(ValueError, '单位缺失'):
            self.plan()
        invoice.items[0].unit = '套'
        other = deepcopy(invoice.items[0])
        other.unit = '个'
        invoice.items.append(other)
        with self.assertRaisesRegex(ValueError, '单位缺失或不唯一'):
            self.plan()

    def test_customer_changes_only_consignee(self):
        original = self.plan()
        customer = 'Wal-Mart Store Inc.601 N. Walton Bentonville BENTONVILLE\nUS'
        for row in self.packing.rows:
            row.values['customer'] = customer
        updated = self.plan()
        self.assertEqual(updated['headers']['C8'], customer)
        original['headers']['C8'] = customer
        self.assertEqual(updated, original)

    def test_customer_missing_or_conflicting_is_not_guessed(self):
        for value in (None, '', '   '):
            with self.subTest(value=value):
                self.packing.rows[1].values['customer'] = value
                with self.assertRaisesRegex(ValueError, '第 5 行 Customer 为空'):
                    self.plan()
        self.packing.rows[1].values['customer'] = 'Another Customer'
        with self.assertRaisesRegex(ValueError, '多个不同的 Customer'):
            self.plan()

    def test_multiple_products_use_each_invoice_unit(self):
        for kind, doc in self.report.documents['TEST001'].items():
            first = doc.items[0]
            for field in ('quantity', 'cartons', 'amount', 'net_weight', 'gross_weight'):
                setattr(first, field, getattr(first, field) / 2)
            second = deepcopy(first)
            second.name = '填充玩具/第二商品'
            second.unit = '件'
            doc.items.append(second)
        self.packing.rows[1].values.update(item='00234', chinese_name='第二商品')
        rows = self.plan()['rows']
        self.assertEqual([row[12] for row in rows], ['套', '件'])
        self.assertEqual([row[0] for row in rows], [1, 2])

    def test_ambiguous_or_missing_product_needs_selection(self):
        self.packing.rows[0].values['chinese_name'] = '未知商品'
        with self.assertRaisesRegex(ValueError, '请确认货号'):
            self.plan()
        self.packing.rows[0].values['chinese_name'] = '测试商品'
        doc = self.report.documents['TEST001']['装箱单']
        doc.items.append(deepcopy(doc.items[0]))
        self.assertEqual(generation_options(self.report, self.packing)['product_mapping'], {})

    def test_mismatches_block_generation(self):
        self.packing.rows[0].values['quantity'] += 1
        with self.assertRaisesRegex(ValueError, '数量与箱数'):
            self.plan()
        self.packing.rows[0].values['quantity'] -= 1
        self.packing.rows[0].values['gross_weight'] += 1
        with self.assertRaisesRegex(ValueError, 'gross_weight 汇总不一致'):
            self.plan()
        self.report.failures = 1
        with self.assertRaisesRegex(ValueError, '核对尚未通过'):
            self.plan()

    def test_unsafe_contract_filename_rejected(self):
        with self.assertRaises(ValueError):
            output_filename('../invalid')

    @unittest.skipUnless(os.name == 'nt', 'Requires Windows and Microsoft Excel')
    def test_native_xls_preserves_styles_formulas_ids_and_removes_example_rows(self):
        original = ASN_TEMPLATE.read_bytes()
        customer = 'Wal-Mart Store Inc.601 N. Walton Bentonville BENTONVILLE\nUS'
        for row in self.packing.rows:
            row.values['customer'] = customer
        with TemporaryDirectory() as folder:
            path = write_asn_workbook(self.plan(), folder)
            self.assertEqual(path.name, 'ZLC盐田综合保税区仓库 -入仓预申报ASN表格-TEST001.xls')
            self.assertEqual(path.read_bytes()[:8], bytes.fromhex('d0cf11e0a1b11ae1'))
            book = xlrd.open_workbook(str(path), formatting_info=True)
            template = xlrd.open_workbook(str(ASN_TEMPLATE), formatting_info=True)
            try:
                self.assertEqual(book.sheet_names(), ['报关资料与ASN'])
                sheet = book.sheet_by_index(0)
                self.assertEqual(sheet.cell_value(7, 2), customer)
                self.assertEqual(sheet.cell_value(25, 5), '00123')
                self.assertEqual(sheet.cell_value(25, 29), 30)
                self.assertEqual(sheet.cell_value(26, 29), 30)
                self.assertEqual(sheet.cell_value(25, 20), '')
                self.assertFalse(any(v != '' for r in range(27, sheet.nrows) for v in sheet.row_values(r)))
                source = template.sheet_by_name('报关资料与ASN')
                self.assertEqual(sorted(sheet.merged_cells), sorted(source.merged_cells))
                style = book.xf_list[sheet.cell_xf_index(25, 1)]
                old_style = template.xf_list[source.cell_xf_index(25, 1)]
                self.assertEqual(style.border.bottom_line_style, old_style.border.bottom_line_style)
                self.assertEqual(sheet.colinfo_map[1].width, source.colinfo_map[1].width)
            finally:
                book.release_resources()
                template.release_resources()
            with self.assertRaises(FileExistsError):
                write_asn_workbook(self.plan(), folder)
        self.assertEqual(ASN_TEMPLATE.read_bytes(), original)

    def test_local_sample_twelve_rows_totals_and_units(self):
        directory = Path(__file__).resolve().parents[4] / 'asn'
        sources = list(directory.glob('PL*.xlsx'))
        if not sources:
            self.skipTest('Local sample unavailable')
        report = audit_directory(directory)
        packing = read_customs_packing(sources[0])
        plan = build_asn_plan(report, packing, ASN_TEMPLATE, generation_options(report, packing))
        self.assertEqual(len(plan['rows']), 12)
        self.assertEqual(sum(r[9] for r in plan['rows']), 309)
        self.assertEqual(sum(r[11] for r in plan['rows']), 1236)
        self.assertTrue(all(r[12] == r[27] == r[32] == '套' for r in plan['rows']))

    @unittest.skipUnless(os.name == 'nt', 'Requires Windows and Microsoft Excel')
    def test_native_single_row_and_extension_beyond_sample(self):
        for count in (1, 15):
            with self.subTest(count=count), TemporaryDirectory() as folder:
                self.report, self.packing = fixture()
                self.packing.rows = [deepcopy(self.packing.rows[0]) for _ in range(count)]
                for i, row in enumerate(self.packing.rows):
                    row.values['so'] = f'SO-{i + 1:03}'
                for doc in self.report.documents['TEST001'].values():
                    for field in ('quantity', 'cartons', 'amount', 'net_weight', 'gross_weight'):
                        setattr(doc.items[0], field, getattr(doc.items[0], field) * count / 2)
                path = write_asn_workbook(self.plan(), folder)
                book = xlrd.open_workbook(str(path))
                try:
                    sheet = book.sheet_by_index(0)
                    actual = [sheet.cell_value(r, 1) for r in range(25, sheet.nrows) if sheet.cell_value(r, 1)]
                    self.assertEqual(actual, [f'SO-{i + 1:03}' for i in range(count)])
                    self.assertEqual(sheet.cell_value(24 + count, 29), 30)
                finally:
                    book.release_resources()


if __name__ == '__main__':
    unittest.main()
