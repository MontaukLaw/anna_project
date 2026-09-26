from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil
import unittest

from openpyxl import Workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from services.asn_service import Sheet, audit_directory, number


SAMPLES = Path(__file__).resolve().parents[4] / 'asn'
CONTRACT = 'SCKTY260914'
KINDS = ('装箱单', '香港合同', '形式发票')


class AsnTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Small synthetic fixtures keep error-path tests runnable without local business files.
        fixtures = {
            '装箱单': {'C4': '协议合同号：', 'E4': CONTRACT, 'D7': '货名', 'E7': '商品编码',
                      'I7': '每箱数量', 'N7': '总箱数', 'Q7': '消费国', 'K8': '单\n位', 'L8': '箱数',
                      'N8': '总\n数量', 'O8': '总\n净重', 'P8': '总\n毛重', 'I9': '单价', 'J9': '总价',
                      'D10': '填充玩具/投影米奇', 'E10': '9503008390', 'I10': 7.5, 'J10': 9270,
                      'K10': '纸箱', 'L10': 309, 'N10': 1236, 'O10': 834.3, 'P10': 1062.96, 'Q10': '美国',
                      'C31': '合计', 'L31': 309, 'N31': 1236, 'N32': 'TOTAL:USD', 'P32': 9270},
            '香港合同': {'D2': '合同号：\nNO.', 'E2': CONTRACT, 'D3': '日期：\nDATE', 'E3': datetime(2026, 5, 10),
                        'A9': '1. 品名及规格\nCommodity &Specification', 'C9': '2. 数量\nQuantity',
                        'D9': '3. 单价(USD)\nUnit Price', 'E9': '4. 金额\nAmount', 'A10': '填充玩具/投影米奇',
                        'B10': '3|0|动物玩偶|带有动力装置|Disney牌', 'C10': 1236, 'D10': 7.5, 'E10': 9270,
                        'D12': 'Total Amount', 'E12': 9270, 'A13': '5. 总值： USD', 'B13': 9270, 'A14': 'Total Value'},
            '形式发票': {'F6': '合同号码:', 'G6': CONTRACT, 'F8': '币别:', 'G8': '美元', 'B10': '标记号码',
                        'C10': '货物名称', 'D10': '数 量', 'E10': '单位', 'F10': '单价(USD)', 'G10': '总价(USD)',
                        'B11': 1, 'C11': '填充玩具/投影米奇', 'D11': 1236, 'E11': '套', 'F11': 7.5, 'G11': 9270,
                        'B29': '合计(USD)', 'G29': 9270}}
        cls.source_sheets = {}
        for kind, cells in fixtures.items():
            rows = [[None] * 18 for _ in range(34)]
            for address, value in cells.items():
                column, row = coordinate_from_string(address)
                rows[row - 1][column_index_from_string(column) - 1] = value
            cls.source_sheets[kind] = [Sheet('Sheet1', rows)]

    def setUp(self):
        self.temp = TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.directory = Path(self.temp.name)
        for kind in KINDS:
            self.write(kind)

    def write(self, kind, changes=None):
        book = Workbook()
        book.remove(book.active)
        for source in self.source_sheets[kind]:
            sheet = book.create_sheet(source.name)
            for row in source.rows:
                sheet.append(row)
        for address, value in (changes or {}).items():
            book.worksheets[0][address] = value
        path = self.directory / f'{kind}{CONTRACT}.xlsx'
        book.save(path)
        book.close()
        return path

    def errors(self, report):
        return '\n'.join(text for text, level in report.entries if level == 'ERROR')

    def test_sample_values_and_multilevel_headers(self):
        directories = [self.directory]
        if (SAMPLES / f'装箱单{CONTRACT}.xls').exists():
            directories.append(SAMPLES)
        for directory in directories:
            with self.subTest(directory=directory):
                result = audit_directory(directory)
                self.assertTrue(result.passed, self.errors(result))
                docs = result.documents[CONTRACT]
                packing = docs['装箱单']
                self.assertEqual(packing.total, Decimal('9270'))
                self.assertEqual(packing.quantity, 1236)
                self.assertEqual(packing.items[0].cartons, 309)
                self.assertEqual(packing.items[0].per_carton, 4)
                self.assertEqual(packing.items[0].code, '9503008390')
                self.assertEqual(packing.items[0].country, '美国')
                self.assertEqual(packing.items[0].price, Decimal('7.5'))
                self.assertEqual(docs['香港合同'].signed_date, '2026-05-10')
                self.assertIn('Disney', docs['香港合同'].items[0].specification)

    def test_contract_number_checked_before_detail_and_does_not_stop_other_files(self):
        for kind, cell in (('装箱单', 'E4'), ('香港合同', 'E2'), ('形式发票', 'G6')):
            with self.subTest(kind=kind):
                self.write(kind, {cell: 'WRONG123'})
                result = audit_directory(self.directory)
                errors = self.errors(result)
                self.assertFalse(result.passed)
                self.assertIn('WRONG123', errors)
                self.assertIn(CONTRACT, errors)
                self.assertIn(cell, errors)
                self.assertEqual(len(result.documents[CONTRACT]), 3)
                entries = [text for text, _ in result.entries]
                identity = next(i for i, text in enumerate(entries) if f'{kind}{CONTRACT}.xlsx 合同号核对' in text)
                detail = next(i for i, text in enumerate(entries) if text.startswith(f'{kind}{CONTRACT}.xlsx 第 '))
                self.assertLess(identity, detail)
                self.write(kind)

    def test_missing_and_duplicate_files_never_pass(self):
        invoice = self.directory / f'形式发票{CONTRACT}.xlsx'
        invoice.unlink()
        report = audit_directory(self.directory)
        self.assertIn('缺少 形式发票', self.errors(report))
        self.write('形式发票')
        shutil.copy(self.directory / f'装箱单{CONTRACT}.xlsx', self.directory / f'装箱单{CONTRACT}.xls')
        report = audit_directory(self.directory)
        self.assertFalse(report.passed)
        self.assertIn('重复', self.errors(report))
        self.assertNotIn('装箱单', report.documents[CONTRACT])

    def test_mismatches_missing_values_and_formula_cache(self):
        cases = [
            ('香港合同', {'E10': 9000, 'E12': 9000, 'B13': 9000}, '三单总金额'),
            ('形式发票', {'C11': '另一种玩具'}, '三单品名'),
            ('形式发票', {'D11': 1235}, '三单总数量'),
            ('形式发票', {'F11': 7.6}, '数量、单价、金额'),
            ('形式发票', {'G11': '=D11*F11'}, '明细完整性'),
            ('形式发票', {'G29': None}, '合计金额'),
            ('形式发票', {'G8': '人民币'}, '币种'),
            ('香港合同', {'E3': 'not a date'}, '合同签订日期'),
            ('装箱单', {'E10': None}, '装箱资料'),
            ('装箱单', {'E10': '#N/A'}, '装箱资料'),
            ('装箱单', {'Q10': '#VALUE!'}, '装箱资料'),
            ('装箱单', {'L31': 308}, '总箱数合计'),
            ('装箱单', {'N31': 1200}, '总数量合计'),
            ('装箱单', {'L10': 308}, '每箱数量 × 箱数'),
        ]
        for kind, changes, expected in cases:
            with self.subTest(kind=kind, changes=changes):
                self.write(kind, changes)
                result = audit_directory(self.directory)
                self.assertFalse(result.passed)
                self.assertIn(expected, self.errors(result))
                self.write(kind)

    def test_leading_zeros_and_unchanged_sources(self):
        path = self.write('装箱单', {'E10': '0012345678'})
        before = path.read_bytes()
        report = audit_directory(self.directory)
        self.assertEqual(report.documents[CONTRACT]['装箱单'].items[0].code, '0012345678')
        self.assertEqual(path.read_bytes(), before)

    def test_corrupt_workbook_reports_filename_and_continues(self):
        (self.directory / f'装箱单{CONTRACT}.xlsx').write_bytes(b'broken')
        result = audit_directory(self.directory)
        self.assertFalse(result.passed)
        self.assertIn(f'装箱单{CONTRACT}.xlsx 读取', self.errors(result))
        self.assertIn('形式发票', result.documents[CONTRACT])

    def test_split_rows_are_grouped_and_row_level_mismatch_cannot_hide_in_totals(self):
        self.write('形式发票', {'D11': 600, 'G11': 4500, 'B12': 2, 'C12': '填充玩具/投影米奇',
                              'D12': 636, 'F12': 7.5, 'G12': 4770})
        report = audit_directory(self.directory)
        self.assertTrue(report.passed, self.errors(report))
        self.write('形式发票', {'D11': 618, 'F11': 7, 'G11': 4326, 'B12': 2, 'C12': '填充玩具/投影米奇',
                              'D12': 618, 'F12': 8, 'G12': 4944})
        report = audit_directory(self.directory)
        self.assertFalse(report.passed)
        self.assertIn('数量、单价、金额', self.errors(report))
        self.assertNotIn('三单总金额', self.errors(report))

    def test_multiple_contracts_and_unrelated_files(self):
        for kind, cell in (('装箱单', 'E4'), ('香港合同', 'E2'), ('形式发票', 'G6')):
            path = self.write(kind, {cell: 'ABC123'})
            path.rename(self.directory / f'{kind}ABC123.xlsx')
            self.write(kind)
        (self.directory / '~$形式发票ABC123.xlsx').write_bytes(b'lock')
        (self.directory / 'unrelated.xlsx').write_bytes(b'not an input')
        result = audit_directory(self.directory)
        self.assertTrue(result.passed, self.errors(result))
        self.assertEqual(set(result.documents), {CONTRACT, 'ABC123'})


class AsnEdgeTests(unittest.TestCase):
    def test_empty_or_missing_directory_and_numeric_parsing(self):
        with TemporaryDirectory() as temporary:
            self.assertFalse(audit_directory(Path(temporary)).passed)
            self.assertFalse(audit_directory(Path(temporary) / 'missing').passed)
        self.assertEqual(number(0), 0)
        self.assertEqual(number('USD 9,270.00'), 9270)
        self.assertIsNone(number(''))
        self.assertIsNone(number('NaN'))
