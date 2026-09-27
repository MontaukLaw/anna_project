import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from test_asn_generation import fixture
from services.settings_service import read_settings
from services.hs_units_service import LegalUnits, LegalUnitError


@unittest.skipUnless(os.name == 'nt', 'Requires Windows Tk')
class GenerationUITests(unittest.TestCase):
    def test_xinghui_generates_confirmed_pallet_fields(self):
        import customtkinter as ctk
        from ui.asn_tab import AsnTab
        report, packing = fixture()
        report.documents['TEST001']['香港合同'].items[0].specification = '3|0|玩具|Disney牌'
        for row in packing.rows:
            row.values['packaging'] = '1箱/卡板'
        def worker(*, target, args, **kwargs):
            return Mock(start=lambda: target(*args))
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
                     patch('ui.asn_tab.load_legal_units', return_value={'9503008390': LegalUnits('套', '千克')}), \
                     patch('ui.asn_tab.write_asn_workbook', return_value=folder / 'ASN.xlsm') as writer:
                    tab = AsnTab(root)
                    tab.wait_window = Mock()
                    tab.report, tab.customs_data = report, packing
                    tab.directory, tab.customs_path, tab.output_directory = folder, packing.source, folder
                    tab.warehouse_var.set('xinghui')
                    tab.generate()
                    tab.poll()
                    plan = writer.call_args.args[0]
                    self.assertEqual(plan['warehouse'], 'xinghui')
                    self.assertEqual(plan['headers']['D16'], '是')
                    self.assertEqual(plan['rows'][0][4], '000987-SKU')
                    self.assertEqual(plan['rows'][0][26], 'Disney牌')
                    self.assertEqual(plan['rows'][0][29:31], [1, 1])
                    self.assertFalse(tab.busy)
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()

    def test_xinghui_choice_is_saved_and_disabled_while_busy(self):
        import customtkinter as ctk
        from config import XINGHUI_ASN_TEMPLATE
        from ui.asn_tab import AsnTab
        with TemporaryDirectory() as temporary:
            root = ctk.CTk()
            root.withdraw()
            try:
                with patch('ui.asn_tab.SETTINGS_FILE', Path(temporary) / 'settings.json'):
                    tab = AsnTab(root)
                    self.assertEqual(set(tab.warehouse_buttons), {'yixing', 'zhongtong', 'xinghui'})
                    tab.warehouse_buttons['xinghui'].invoke()
                    self.assertEqual(tab.warehouse_template(tab.warehouse_var.get()), XINGHUI_ASN_TEMPLATE)
                    restored = AsnTab(root)
                    self.assertEqual(restored.warehouse_var.get(), 'xinghui')
                    tab.busy = True
                    tab.update_generate_state()
                    self.assertEqual(tab.warehouse_buttons['xinghui'].cget('state'), 'disabled')
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()

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
                     patch('ui.asn_tab.load_legal_units', return_value={'9503008390': LegalUnits('个', '千克')}) as lookup, \
                     patch('ui.asn_tab.write_asn_workbook', return_value=folder / 'ASN.xls') as writer, \
                     patch('ui.asn_tab.filedialog.askdirectory', return_value='') as picker, \
                     patch('ui.asn_tab.AsnGenerationDialog') as dialog:
                    tab = AsnTab(root)
                    self.assertEqual(tab.warehouse_var.get(), 'yixing')
                    tab.warehouse_buttons['zhongtong'].invoke()
                    self.assertEqual(tab.warehouse_var.get(), 'zhongtong')
                    self.assertEqual(read_settings(folder / 'settings.json')['asn_warehouse'], 'zhongtong')
                    warehouse_restored = AsnTab(root)
                    self.assertEqual(warehouse_restored.warehouse_var.get(), 'zhongtong')
                    tab.warehouse_buttons['yixing'].invoke()
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
                    self.assertTrue(all(control.cget('state') == 'disabled'
                                        for control in tab.warehouse_buttons.values()))
                    tab.poll()
                    self.assertFalse(tab.busy)
                    self.assertIn('2 行明细', tab.status.cget('text'))
                    self.assertEqual(writer.call_count, 1)
                    self.assertEqual(writer.call_args.args[1], folder.resolve())
                    self.assertEqual(tab.output_button.cget('state'), 'normal')
                    self.assertTrue(all(control.cget('state') == 'normal'
                                        for control in tab.warehouse_buttons.values()))
                    tab.warehouse_buttons['zhongtong'].invoke()
                    tab.generate()
                    tab.poll()
                    self.assertEqual(writer.call_count, 2)
                    self.assertEqual(writer.call_args.args[0]['warehouse'], 'zhongtong')
                    self.assertEqual(writer.call_args.args[0]['rows'][0][3], '000987-SKU')
                    self.assertEqual(writer.call_args.args[0]['rows'][0][18], '个')
                    self.assertFalse(tab.busy)
                    lookup.side_effect = LegalUnitError('网络连接失败，请检查网络后重新生成')
                    tab.generate()
                    tab.poll()
                    self.assertEqual(writer.call_count, 2)
                    self.assertIn('网络连接失败', tab.logs.textbox.get('1.0', 'end'))
                    self.assertEqual(tab.generate_button.cget('state'), 'normal')
                    self.assertFalse(tab.busy)
                    lookup.side_effect = None
                    tab.warehouse_buttons['yixing'].invoke()
                    fresh_report, _ = fixture()
                    fresh_report.failures = 1
                    with patch('ui.asn_tab.audit_directory', return_value=fresh_report):
                        tab.generate()
                        tab.poll()
                    self.assertEqual(writer.call_count, 2)
                    self.assertIn('未生成', tab.status.cget('text'))
                    self.assertIn('核对尚未通过', tab.logs.textbox.get('1.0', 'end'))
                    self.assertEqual(len(list((folder / 'logs').glob('*.log'))), 1)
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()
