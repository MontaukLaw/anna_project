from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from services.asn_log_service import DailyAuditLog
from services.asn_service import AuditReport, result_messages


class DailyAuditLogTests(unittest.TestCase):
    def test_same_day_appends_across_runs_and_midnight_changes_file(self):
        with TemporaryDirectory() as temporary:
            directory = Path(temporary) / 'logs'
            current = datetime(2026, 9, 26, 23, 59, 59)
            first = DailyAuditLog(directory, now=lambda: current)
            first.write('合同 A：开始核对', 'START')
            second = DailyAuditLog(directory, now=lambda: current)
            second.write('合同 B：表内=123 (Sheet1!E4)', 'INFO')
            first_path = directory / 'asn-2026-09-26.log'
            content = first_path.read_text(encoding='utf-8')
            self.assertIn('合同 A', content)
            self.assertIn('合同 B', content)
            self.assertIn(first.run_id, content)
            self.assertIn(second.run_id, content)
            self.assertEqual(len(list(directory.glob('*.log'))), 1)
            current = datetime(2026, 9, 27, 0, 0, 1)
            second.write('检查结束', 'SUCCESS')
            next_path = directory / 'asn-2026-09-27.log'
            self.assertEqual(first_path.read_text(encoding='utf-8'), content)
            self.assertIn('检查结束', next_path.read_text(encoding='utf-8'))
            self.assertEqual(second.paths, [first_path, next_path])

    def test_write_failure_is_exposed_without_interrupting_the_audit(self):
        with TemporaryDirectory() as temporary:
            unavailable = Path(temporary) / 'file_instead_of_directory'
            unavailable.write_text('occupied', encoding='utf-8')
            journal = DailyAuditLog(unavailable)
            journal.write('anything')
            self.assertIsNotNone(journal.error)
            self.assertEqual(journal.paths, [])
            journal.write('audit can continue')


class ResultSummaryTests(unittest.TestCase):
    def test_mixed_contracts_only_expose_results_and_actionable_errors(self):
        report = AuditReport(
            entries=[('读取每一行的冗长细节', 'INFO')],
            documents={'A': {}, 'B': {}}, checks=30, failures=1,
            issues=[('装箱单B.xls 合同号核对', '文件名合同号=B；表内=WRONG (Sheet1!E4)')],
            contract_failures={'A': 0, 'B': 1})
        messages = result_messages(report)
        self.assertEqual(len(messages), 3)
        self.assertEqual(messages[0][1], 'SUCCESS')
        self.assertEqual(messages[1][1], 'ERROR')
        self.assertIn('装箱单B.xls', messages[2][0])
        self.assertIn('WRONG', messages[2][0])
        self.assertNotIn('Sheet1!', messages[2][0])
        self.assertNotIn('冗长细节', str(messages))
