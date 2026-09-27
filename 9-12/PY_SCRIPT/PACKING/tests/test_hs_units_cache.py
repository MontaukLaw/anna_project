import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from services.hs_units_service import LegalUnitError, LegalUnits, load_legal_units
from services.hs_units_cache import HSUnitCache


class HSUnitCacheTests(unittest.TestCase):
    def setUp(self):
        temporary = TemporaryDirectory()
        self.addCleanup(temporary.cleanup)
        self.path = Path(temporary.name) / 'hs_units_cache.json'
        setting = patch('services.hs_units_service.HS_UNIT_CACHE_FILE', self.path)
        setting.start()
        self.addCleanup(setting.stop)

    def read(self):
        return json.loads(self.path.read_text(encoding='utf-8'))

    def test_persistent_success_avoids_network_and_keeps_other_codes(self):
        with patch('services.hs_units_service.query_legal_units', side_effect=[
                LegalUnits('个', '千克'), LegalUnits('套')]) as query:
            load_legal_units(['9503002100'])
            load_legal_units(['9503008900'])
            self.assertEqual(query.call_count, 2)
        data = self.read()
        self.assertEqual(len(data['entries']), 2)
        self.assertEqual(len(data['queries']), 2)
        self.assertTrue(data['entries']['9503002100']['queried_at'])
        self.assertEqual(data['entries']['9503002100']['url'],
                         'https://www.hsbianma.com/search?keywords=9503002100')
        log, failure = Mock(), Mock()
        # A fresh call reads the persisted file, independent of any in-memory results.
        with patch('services.hs_units_service.query_legal_units', side_effect=AssertionError('Unexpected network')) as query:
            result = load_legal_units(['9503002100', '9503008900'], log, failure)
            query.assert_not_called()
            failure.assert_not_called()
        self.assertEqual(result['9503002100'].first, '个')
        self.assertEqual(result['9503008900'].second, '')
        self.assertTrue(all('本地缓存' in call.args[0] for call in log.call_args_list))
        self.assertEqual(self.read(), data)

    def test_each_retry_is_recorded_but_failures_are_not_cache_hits(self):
        with patch('services.hs_units_service.query_legal_units', side_effect=LegalUnitError('timeout')):
            with self.assertRaises(LegalUnitError):
                load_legal_units(['9503002100'])
        self.assertEqual(self.read()['entries'], {})
        with patch('services.hs_units_service.query_legal_units', side_effect=[
                LegalUnitError('offline'), LegalUnits('个', '千克')]) as query:
            load_legal_units(['9503002100'], on_failure=lambda *_: 'retry')
            self.assertEqual(query.call_count, 2)
        data = self.read()
        self.assertEqual([row['status'] for row in data['queries']], ['error', 'error', 'success'])
        self.assertEqual(data['queries'][0]['error'], 'timeout')
        self.assertEqual(data['entries']['9503002100']['status'], 'success')

    def test_all_cached_codes_are_resolved_before_first_network_request(self):
        with patch('services.hs_units_service.query_legal_units', return_value=LegalUnits('个', '千克')):
            load_legal_units(['9503008900'])
        messages = []
        def query(code):
            self.assertEqual(code, '9503002100')
            self.assertTrue(any('使用本地缓存：商品编码 9503008900' in message for message in messages))
            self.assertTrue(any('命中 1 个，待联网 1 个' in message for message in messages))
            return LegalUnits('套', '千克')
        with patch('services.hs_units_service.query_legal_units', side_effect=query) as request:
            result = load_legal_units(['9503002100', '9503008900', '9503002100'], messages.append)
            request.assert_called_once_with('9503002100')
        self.assertEqual(list(result), ['9503002100', '9503008900'])
        # A fresh cache instance must read once and never request the network.
        original_read = HSUnitCache.read
        reads = []
        def read(cache):
            reads.append(cache.path)
            return original_read(cache)
        with patch.object(HSUnitCache, 'read', read), \
             patch('services.hs_units_service.query_legal_units', side_effect=AssertionError('Network forbidden')):
            load_legal_units(['9503002100', '9503008900', '9503002100'])
        self.assertEqual(reads, [self.path])

    def test_manual_choice_never_becomes_a_network_cache_entry(self):
        with patch('services.hs_units_service.query_legal_units', side_effect=LegalUnitError('offline')):
            load_legal_units(['9503002100'], on_failure=lambda *_: LegalUnits('个', '千克', manual=True))
        self.assertEqual(self.read()['entries'], {})
        self.assertEqual(self.read()['queries'][0]['status'], 'error')

    def test_corrupt_json_is_backed_up_and_recreated_after_query(self):
        damaged = '{invalid JSON'
        self.path.write_text(damaged, encoding='utf-8')
        log = Mock()
        with patch('services.hs_units_service.query_legal_units', return_value=LegalUnits('个')) as query:
            load_legal_units(['9503002100'], log)
            query.assert_called_once()
        backup = next(self.path.parent.glob('*.bak'))
        self.assertEqual(backup.read_text(encoding='utf-8'), damaged)
        self.assertEqual(self.read()['entries']['9503002100']['first'], '个')
        self.assertTrue(any('警告' in call.args[0] for call in log.call_args_list))

    def test_invalid_unit_entry_is_replaced_by_fresh_result(self):
        self.path.write_text(json.dumps({'version': 1, 'entries': {'9503002100': {
            'status': 'success', 'code': '9503002100', 'first': ['个'], 'second': '千克',
            'url': 'https://www.hsbianma.com/search?keywords=9503002100'}}, 'queries': []}), encoding='utf-8')
        with patch('services.hs_units_service.query_legal_units', return_value=LegalUnits('套', '千克')) as query:
            result = load_legal_units(['9503002100'])
            query.assert_called_once()
        self.assertEqual(result['9503002100'].first, '套')
        self.assertEqual(self.read()['entries']['9503002100']['first'], '套')

    def test_unwritable_cache_warns_and_keeps_successful_query(self):
        log = Mock()
        with patch('services.hs_units_cache.NamedTemporaryFile', side_effect=PermissionError('read only')), \
             patch('services.hs_units_service.query_legal_units', return_value=LegalUnits('个')):
            result = load_legal_units(['9503002100'], log)
        self.assertEqual(result['9503002100'].first, '个')
        self.assertTrue(any('保存失败' in call.args[0] for call in log.call_args_list))
