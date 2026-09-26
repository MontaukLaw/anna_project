import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from test_asn_generation import fixture
from services.settings_service import read_settings


@unittest.skipUnless(os.name == 'nt', 'Requires Windows Tk')
class GenerationUITests(unittest.TestCase):
    def test_button_cancel_success_and_failed_fresh_audit(self):
        import customtkinter as ctk
        from ui.asn_tab import AsnTab

        def worker(*, target, args, **kwargs):
            return Mock(start=lambda: target(*args))

        report, packing = fixture()
        with TemporaryDirectory() as temporary:
            folder = Path(temporary)
            root = ctk.CTk()
            root.withdraw()
            try:
                with patch('ui.asn_tab.SETTINGS_FILE', folder / 'settings.json'), \
                     patch('ui.asn_tab.ASN_LOG_DIR', folder / 'logs'), \
                     patch('ui.asn_tab.Thread', side_effect=worker), \
                     patch('ui.asn_tab.audit_directory', return_value=report), \
                     patch('ui.asn_tab.read_customs_packing', return_value=packing), \
                     patch('ui.asn_tab.write_asn_workbook', return_value=folder / 'ASN.xls') as writer, \
                     patch('ui.asn_tab.filedialog.askdirectory', return_value='') as picker, \
                     patch('ui.asn_tab.AsnGenerationDialog') as dialog:
                    tab = AsnTab(root)
                    self.assertEqual(tab.generate_button.cget('state'), 'disabled')
                    tab.report = report
                    tab.update_generate_state()
                    self.assertEqual(tab.generate_button.cget('state'), 'disabled')
                    tab.customs_data = packing
                    tab.directory = folder
                    tab.customs_path = packing.source
                    tab.update_generate_state()
                    self.assertEqual(tab.generate_button.cget('state'), 'normal')
                    tab.generate()
                    writer.assert_not_called()
                    dialog.assert_not_called()
                    self.assertFalse(tab.busy)
                    picker.return_value = str(folder)
                    tab.output_button.invoke()
                    self.assertEqual(tab.output_directory, folder.resolve())
                    self.assertEqual(read_settings(folder / 'settings.json')['asn_output_directory'], str(folder.resolve()))
                    restored = AsnTab(root)
                    self.assertEqual(restored.output_directory, folder.resolve())
                    self.assertEqual(restored.output_path_var.get(), str(folder.resolve()))
                    picker.return_value = ''
                    tab.output_button.invoke()
                    self.assertEqual(tab.output_directory, folder.resolve())
                    picker.reset_mock()
                    tab.generate()
                    picker.assert_not_called()
                    self.assertTrue(tab.busy)
                    self.assertEqual(tab.generate_button.cget('state'), 'disabled')
                    self.assertEqual(tab.output_button.cget('state'), 'disabled')
                    tab.poll()
                    self.assertFalse(tab.busy)
                    self.assertIn('2 行明细', tab.status.cget('text'))
                    self.assertEqual(writer.call_count, 1)
                    self.assertEqual(writer.call_args.args[1], folder.resolve())
                    self.assertEqual(tab.output_button.cget('state'), 'normal')
                    fresh_report, _ = fixture()
                    fresh_report.failures = 1
                    with patch('ui.asn_tab.audit_directory', return_value=fresh_report):
                        tab.generate()
                        tab.poll()
                    self.assertEqual(writer.call_count, 1)
                    self.assertIn('未生成', tab.status.cget('text'))
                    self.assertIn('核对尚未通过', tab.logs.textbox.get('1.0', 'end'))
                    self.assertEqual(len(list((folder / 'logs').glob('*.log'))), 1)
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()
