from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
from unittest.mock import patch, Mock
import os
import unittest

from services.packing_job import run_packing_job
from services.order_pdf_service import read_order_number
from models.packing import PackingDataError


class AllPdfPackingTests(unittest.TestCase):
    def test_scan_all_files_and_continue_after_bad_pdf(self):
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


@unittest.skipUnless(os.name == 'nt', 'Requires the Windows Tk desktop runtime')
class PdfPackingUITests(unittest.TestCase):
    def test_directory_processing_restores_controls_after_success_and_error(self):
        from ui.app import PackingApp

        def worker(*, target, args, kwargs=None, **options):
            return Mock(start=lambda: target(*args, **(kwargs or {})))

        with TemporaryDirectory() as temporary:
            root = Path(temporary)
            settings = root / 'settings.json'
            (root / 'sub').mkdir()
            pdf = root / 'sub' / 'order.pdf'
            pdf.touch()
            with patch('ui.app.SETTINGS_FILE', settings), \
                 patch('ui.app.PACKING_RULES_FILE', root / 'packing_rules.json'), \
                 patch('ui.asn_tab.SETTINGS_FILE', settings), \
                 patch('ui.factory_tab.SETTINGS_FILE', settings), \
                 patch('ui.factory_tab.FactoryTab.restore'), \
                 patch('ui.app.Thread', side_effect=worker), \
                 patch('ui.app.messagebox.showinfo') as notice, \
                 patch('services.packing_job.read_order_number', return_value='POHK-26-12345-001'), \
                 patch('services.packing_job.generate_packing_lists', return_value=([], [], [])) as generate:
                app = PackingApp()
                app.withdraw()
                try:
                    # Missing source data should still produce a useful prompt.
                    app.all_pdf_button.invoke()
                    generate.assert_not_called()
                    self.assertFalse(app.packing_busy)

                    app.order_search_directory = root
                    app.product_catalog = object()
                    app.order_schedule = object()
                    app.excel_output_directory = root / 'output'
                    app.all_pdf_button.invoke()
                    self.assertTrue(app.packing_busy)
                    self.assertEqual(app.all_pdf_button.cget('state'), 'disabled')
                    app._poll_events()
                    self.assertFalse(app.packing_busy)
                    self.assertEqual(generate.call_args.args[0], {'POHK-26-12345-001': [pdf]})
                    self.assertIn('order.pdf', app.table_text.get('1.0', 'end'))
                    self.assertEqual(app.table_info.cget('text'), '1 个 PDF')
                    notice.assert_called_once()

                    with patch('services.packing_job.scan_order_directory', side_effect=OSError('cannot read directory')):
                        app.all_pdf_button.invoke()
                        app._poll_events()
                    self.assertFalse(app.packing_busy)
                    self.assertIn('cannot read directory', app.logs.textbox.get('1.0', 'end'))
                    for control in (app.all_pdf_button, app.order_directory_button,
                                    app.product_button, app.schedule_button):
                        self.assertEqual(control.cget('state'), 'normal')
                finally:
                    for timer in app.tk.call('after', 'info'):
                        app.tk.call('after', 'cancel', timer)
                    app.destroy()
