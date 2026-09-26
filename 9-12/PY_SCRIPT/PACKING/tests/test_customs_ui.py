from decimal import Decimal
import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from models.customs_packing import CustomsColumn, CustomsPackingData, CustomsPackingRow
from services.settings_service import read_settings


@unittest.skipUnless(os.name == 'nt', 'Requires Windows Tk')
class CustomsUITests(unittest.TestCase):
    def test_selection_rows_logging_cancel_and_failed_replacement(self):
        import customtkinter as ctk
        from ui.asn_tab import AsnTab

        def worker(*, target, args, **kwargs):
            return Mock(start=lambda: target(*args))

        with TemporaryDirectory() as temporary:
            folder = Path(temporary)
            source = folder / 'packing.xlsx'
            columns = [CustomsColumn('quantity', '数量', 8), CustomsColumn('cartons', '箱数', 10)]
            data = CustomsPackingData(source, [CustomsPackingRow('Sheet1', r, columns,
                {'quantity': Decimal(12), 'cartons': Decimal(3)}, {'quantity': 12, 'cartons': 3}) for r in (4, 5)])
            root = ctk.CTk()
            root.withdraw()
            try:
                with patch('ui.asn_tab.SETTINGS_FILE', folder / 'settings.json'), \
                     patch('ui.asn_tab.ASN_LOG_DIR', folder / 'logs'), \
                     patch('ui.asn_tab.Thread', side_effect=worker), \
                     patch('ui.asn_tab.read_customs_packing', return_value=data) as reader, \
                     patch('ui.asn_tab.filedialog.askopenfilename', return_value=str(source)) as picker:
                    tab = AsnTab(root)
                    tab.customs_button.invoke()
                    self.assertTrue(tab.busy)
                    self.assertEqual(tab.select_button.cget('state'), 'disabled')
                    tab.poll()
                    self.assertFalse(tab.busy)
                    self.assertIs(tab.customs_data, data)
                    self.assertIn('2 行明细', tab.status.cget('text'))
                    text = tab.logs.textbox.get('1.0', 'end')
                    self.assertEqual(text.count('序号 | 数量 | 箱数'), 1)
                    self.assertIn('\n1 | 12 | 3\n2 | 12 | 3\n', text)
                    self.assertEqual(tab.logs.textbox._textbox.cget('wrap'), 'none')
                    self.assertIn('合计箱数 6', text)
                    self.assertEqual(read_settings(folder / 'settings.json')['asn_customs_file'], str(source.resolve()))
                    logs = list((folder / 'logs').glob('*.log'))
                    self.assertEqual(len(logs), 1)
                    self.assertIn('2 | 12 | 3', logs[0].read_text(encoding='utf-8'))
                    self.assertIn('Sheet1 第 5 行原始值', logs[0].read_text(encoding='utf-8'))
                    picker.return_value = ''
                    tab.customs_button.invoke()
                    self.assertIs(tab.customs_data, data)
                    self.assertEqual(tab.logs.textbox.get('1.0', 'end'), text)
                    reader.side_effect = ValueError('损坏文件')
                    tab.start_customs_read(folder / 'bad.xlsx')
                    tab.poll()
                    self.assertIsNone(tab.customs_data)
                    self.assertIn('读取失败', tab.status.cget('text'))
                    self.assertIn('损坏文件', tab.logs.textbox.get('1.0', 'end'))
                    self.assertEqual(tab.customs_button.cget('state'), 'normal')
                    self.assertEqual(tab.select_button.cget('state'), 'normal')
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()
