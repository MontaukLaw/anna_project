from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
from unittest.mock import patch, Mock
import unittest

from services.packing_job import run_packing_job
from services.order_pdf_service import read_order_number
from models.packing import PackingDataError


class AllPdfPackingTests(unittest.TestCase):
    def test_scan_all_files_without_recognition_and_continue_after_bad_pdf(self):
        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            (root / 'sub').mkdir()
            files = [root / 'a.pdf', root / 'b.PDF', root / 'sub' / 'c.pdf']
            for path in files:
                path.touch()
            (root / 'ignore.txt').touch()
            events = Queue()
            with patch('services.packing_job.read_order_number', side_effect=[
                PackingDataError('bad PDF'), 'POHK-26-12345-001', 'POHK-26-12345-001'
            ]) as read, patch('services.packing_job.generate_packing_lists', side_effect=[
                ([root / 'output.xlsx'], [], []), ([], [], ['POHK-26-12345-001'])
            ]) as generate:
                run_packing_job(None, None, None, None, root, events, pdf_directory=root)
            self.assertEqual([call.args[0] for call in read.call_args_list], files)
            self.assertEqual(generate.call_count, 2)
            self.assertEqual(generate.call_args_list[0].args[0], {'POHK-26-12345-001': [files[1]]})
            messages = list(events.queue)
            self.assertIn(('packing_pdf_files', files), messages)
            self.assertIn(('packing_finished', ([root / 'output.xlsx'], [str(files[0])], ['POHK-26-12345-001'])), messages)
            self.assertEqual(messages[-1], ('packing_done', None))

    def test_order_number_comes_from_pdf_content(self):
        with patch('services.order_pdf_service.PdfReader') as reader:
            reader.return_value.pages = [Mock(extract_text=lambda: 'PURCHASE ORDER NO.: POHK-26-12345-001')]
            self.assertEqual(read_order_number(Path('arbitrary.pdf')), 'POHK-26-12345-001')
            reader.return_value.pages = [Mock(extract_text=lambda: '')]
            with self.assertRaises(PackingDataError):
                read_order_number(Path('empty.pdf'))

    def test_empty_directory_finishes(self):
        with TemporaryDirectory() as temporary:
            events = Queue()
            root = Path(temporary)
            run_packing_job(None, None, None, None, root, events, pdf_directory=root)
            self.assertIn(('packing_finished', ([], [], [])), list(events.queue))
