from decimal import Decimal as D
from http.client import IncompleteRead
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch
from urllib.error import URLError

from config import ASN_TEMPLATE, ZHONGTONG_ASN_TEMPLATE
from services.asn_generation_service import build_asn_plan, generation_options, write_asn_workbook
from services.asn_quantity_service import quantities_with_units, legal_quantity
from services.hs_units_service import (LegalUnitError, LegalUnitCancelled, LegalUnits, parse_legal_units,
                                      query_legal_units, load_legal_units, manual_legal_units)
from test_asn_generation import fixture


HTML = '''<table class="result"><tr><td>商品编码</td><td>商品名称</td><td>计量单位</td></tr>
<tr><td><font>9503 0021.00</font></td><td>动物玩偶</td><td>个/千克</td></tr>
<tr><td>9503002101</td><td>其他</td><td>套/千克</td></tr></table>'''


class UnitLookupTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        cached_path = patch('services.hs_units_service.HS_UNIT_CACHE_FILE', Path(temporary.name) / 'cache.json')
        cached_path.start()
        self.addCleanup(cached_path.stop)

    def test_retry_only_queries_failed_code_again(self):
        units = LegalUnits('个', '千克')
        choose = Mock(return_value='retry')
        with patch('services.hs_units_service.query_legal_units',
                   side_effect=[units, LegalUnitError('offline'), LegalUnitError('timeout'), units]) as query:
            result = load_legal_units(['9503002100', '9503008900', '9503002100'], on_failure=choose)
        self.assertEqual(list(result), ['9503002100', '9503008900'])
        self.assertEqual([call.args[0] for call in query.call_args_list],
                         ['9503002100', '9503008900', '9503008900', '9503008900'])
        self.assertEqual(choose.call_count, 2)

    def test_manual_choice_is_logged_and_not_reused_on_next_generation(self):
        log = Mock()
        choose = Mock(return_value=manual_legal_units('个', '千克'))
        with patch('services.hs_units_service.query_legal_units', side_effect=LegalUnitError('offline')) as query:
            result = load_legal_units(['9503002100', '9503002100'], log, choose)
            self.assertTrue(result['9503002100'].manual)
            self.assertEqual(query.call_count, 1)
            with self.assertRaises(LegalUnitError):
                load_legal_units(['9503002100'])
            self.assertEqual(query.call_count, 2)
        self.assertTrue(any('用户手动确认' in call.args[0] for call in log.call_args_list))

    def test_cancel_does_not_query_remaining_codes(self):
        with patch('services.hs_units_service.query_legal_units', side_effect=LegalUnitError('offline')) as query:
            with self.assertRaises(LegalUnitCancelled):
                load_legal_units(['9503002100', '9503008900'], on_failure=lambda *_: None)
            self.assertEqual(query.call_count, 1)

    def test_manual_selection_requires_valid_units(self):
        for first, second in [('', '千克'), ('个', '请选择法2单位'), ('个', '个')]:
            with self.subTest(first=first, second=second), self.assertRaises(LegalUnitError):
                manual_legal_units(first, second)
        self.assertEqual(manual_legal_units('个', ''), LegalUnits('个', manual=True))
        choose = Mock(return_value=LegalUnits('个'))
        with self.assertRaises(LegalUnitError):
            load_legal_units(['invalid'], on_failure=choose)
        choose.assert_not_called()

    def test_exact_code_header_and_single_or_two_units(self):
        self.assertEqual(parse_legal_units(HTML, '9503002100'), LegalUnits('个', '千克'))
        self.assertEqual(parse_legal_units(HTML.replace('个/千克', '个'), '9503002100'), LegalUnits('个'))
        with self.assertRaises(LegalUnitError):
            parse_legal_units(HTML, '950300210')
        for html in ('<input value="9503002100">计量单位 个/千克',
                     HTML + HTML.replace('个/千克', '套/千克'), HTML.replace('个/千克', '登录后查看')):
            with self.subTest(html=html), self.assertRaises(LegalUnitError):
                parse_legal_units(html, '9503002100')

    def test_request_uses_code_timeout_and_failure_is_actionable(self):
        response = Mock()
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        response.read.return_value = HTML.encode('utf-8')
        response.headers.get_content_charset.return_value = 'utf-8'
        with patch('services.hs_units_service.urlopen', return_value=response) as opened:
            units = query_legal_units('9503002100')
            self.assertEqual(units.first, '个')
            self.assertEqual(opened.call_args.args[0].full_url,
                             'https://www.hsbianma.com/search?keywords=9503002100')
            self.assertEqual(opened.call_args.kwargs['timeout'], 15)
        for error in (URLError('offline'), TimeoutError('timeout'), IncompleteRead(b'partial')):
            with patch('services.hs_units_service.urlopen', side_effect=error), self.assertRaisesRegex(LegalUnitError, '9503002100.*网络'):
                query_legal_units('9503002100')
        with self.assertRaisesRegex(LegalUnitError, '10 位'):
            query_legal_units('9503002100000')

    def test_only_one_query_per_code_and_reuse_on_next_generation(self):
        with patch('services.hs_units_service.query_legal_units', return_value=LegalUnits('个', '千克')) as query:
            load_legal_units(['9503002100', '9503002100', '9503008900'])
            self.assertEqual(query.call_count, 2)
            query.side_effect = LegalUnitError('offline')
            self.assertEqual(load_legal_units(['9503002100'])['9503002100'].first, '个')
            self.assertEqual(query.call_count, 2)
            with self.assertRaises(LegalUnitError):
                load_legal_units(['9503002101'])


class LegalQuantityTests(unittest.TestCase):
    def test_multi_unit_text_is_strict(self):
        self.assertEqual(quantities_with_units('1,920套\u00a0 23,040个'), {'套': D(1920), '个': D(23040)})
        self.assertEqual(quantities_with_units('1920 套\n23040 个'), {'套': D(1920), '个': D(23040)})
        for text in ('1920套 23040个 未确认', '-1920套', '1920套 1921套', '1.2.3个'):
            with self.subTest(text=text), self.assertRaises(ValueError):
                quantities_with_units(text)

    def test_1920_sets_use_23040_pieces(self):
        report, _ = fixture()
        item = report.documents['TEST001']['装箱单'].items[0]
        item.quantity = D(1920)
        item.quantity_text = '1920套 23040个'
        item.quantity_units = quantities_with_units(item.quantity_text)
        self.assertEqual(legal_quantity(item, D(1920), D('1092.48'), '套', '个'), 23040)
        self.assertEqual(legal_quantity(item, D(960), D('546.24'), '套', '个'), 11520)

    def test_both_warehouses_keep_pricing_and_allocate_legal_counts(self):
        report, packing = fixture()
        item = report.documents['TEST001']['装箱单'].items[0]
        item.quantity_text = '8套 96个'
        item.quantity_units = {'套': D(8), '个': D(96)}
        item.quantity_unit = '套'
        for warehouse, template, cols in [('yixing', ASN_TEMPLATE, (31, 32, 33, 34)),
                                           ('zhongtong', ZHONGTONG_ASN_TEMPLATE, (17, 18, 19, 20))]:
            options = dict(generation_options(report, packing), warehouse=warehouse,
                           legal_units={item.code: LegalUnits('个', '千克')})
            plan = build_asn_plan(report, packing, template, options)
            self.assertEqual([row[cols[0]] for row in plan['rows']], [48, 48])
            self.assertEqual([plan['rows'][0][i] for i in cols], [48, '个', 2.7, '千克'])
            if warehouse == 'zhongtong':
                self.assertEqual(plan['rows'][0][3], '000987-SKU')
                self.assertEqual(plan['rows'][0][12:15], [4, '套', 7.5])
                self.assertEqual(plan['headers']['H15'], 'TEST001')
                self.assertEqual(plan['headers']['H24'], 2)
                self.assertIsNone(plan['headers']['H14'])
                self.assertIsNone(plan['headers']['F19'])

    def test_bare_number_uses_first_unit_and_missing_conversion_stops(self):
        report, _ = fixture()
        item = report.documents['TEST001']['装箱单'].items[0]
        self.assertEqual(legal_quantity(item, D(4), D('2.7'), '套', '个', allow_bare_quantity=True), 4)
        item.quantity_units = {'套': D(8)}
        with self.assertRaisesRegex(ValueError, '换算依据'):
            legal_quantity(item, D(4), D('2.7'), '套', '个', allow_bare_quantity=True)
        item.quantity_units = {'套': D(8), '个': D(9)}
        with self.assertRaisesRegex(ValueError, '不是整数'):
            legal_quantity(item, D(4), D('2.7'), '套', '个')

    def test_missing_lookup_and_preflight_cannot_generate(self):
        report, packing = fixture()
        options = generation_options(report, packing)
        with self.assertRaisesRegex(ValueError, '尚未成功查询'):
            build_asn_plan(report, packing, ASN_TEMPLATE, options)
        plan = build_asn_plan(report, packing, ASN_TEMPLATE, options, validate_only=True)
        with self.assertRaisesRegex(ValueError, '预检查'):
            write_asn_workbook(plan, Path('unused'))

    def test_no_second_unit_leaves_both_second_fields_empty(self):
        report, packing = fixture()
        options = dict(generation_options(report, packing), legal_units={'9503008390': LegalUnits('个')})
        plan = build_asn_plan(report, packing, ASN_TEMPLATE, options)
        self.assertEqual(plan['rows'][0][33:35], [None, ''])
