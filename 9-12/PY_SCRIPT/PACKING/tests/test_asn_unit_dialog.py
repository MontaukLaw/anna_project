import os
from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
from threading import Event, Thread
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from services.hs_units_service import LegalUnitError, LegalUnits
from test_asn_generation import fixture


@unittest.skipUnless(os.name == 'nt', 'Requires Windows Tk')
class UnitDialogTests(unittest.TestCase):
    @staticmethod
    def close_root(root):
        for timer in root.tk.call('after', 'info'):
            root.tk.call('after', 'cancel', timer)
        root.destroy()

    def test_no_default_units_and_explicit_no_second(self):
        import customtkinter as ctk
        from ui.asn_unit_dialog import AsnUnitDialog
        root = ctk.CTk()
        root.withdraw()
        try:
            dialog = AsnUnitDialog(root, '9503002100', '网络超时')
            dialog.confirm()
            self.assertIsNone(dialog.result)
            self.assertTrue(dialog.winfo_exists())
            dialog.first.set('个')
            dialog.confirm()
            self.assertIsNone(dialog.result)
            self.assertIn('法2', dialog.error_label.cget('text'))
            dialog.second.set(dialog.NO_SECOND)
            dialog.confirm()
            self.assertEqual(dialog.result, LegalUnits('个', manual=True))
        finally:
            self.close_root(root)

    def test_both_warehouses_retry_manual_or_cancel_in_real_worker(self):
        import customtkinter as ctk
        from ui.asn_tab import AsnTab
        from ui.asn_unit_dialog import AsnUnitDialog
        for warehouse in ('yixing', 'zhongtong'):
            for action in ('retry', 'manual', 'cancel'):
                with self.subTest(warehouse=warehouse, action=action), TemporaryDirectory() as temporary:
                    folder = Path(temporary)
                    report, packing = fixture()
                    root = ctk.CTk()
                    root.withdraw()
                    dialog_threads = []
                    def show_dialog(*args):
                        from threading import current_thread
                        dialog_threads.append(current_thread().name)
                        dialog = AsnUnitDialog(*args)
                        def choose():
                            if action == 'manual':
                                dialog.first.set('个')
                                dialog.second.set('千克')
                                dialog.confirm()
                            elif action == 'retry':
                                dialog.retry()
                            else:
                                # Closing the window has the same effect as Cancel.
                                dialog.destroy()
                        dialog.after(30, choose)
                        return dialog
                    try:
                        with patch('ui.asn_tab.SETTINGS_FILE', folder / 'settings.json'), \
                             patch('ui.asn_tab.ASN_LOG_DIR', folder / 'logs'), \
                             patch('services.hs_units_service.HS_UNIT_CACHE_FILE', folder / 'cache.json'), \
                             patch('ui.asn_tab.audit_directory', return_value=report), \
                             patch('ui.asn_tab.read_customs_packing', return_value=packing), \
                             patch('services.hs_units_service.query_legal_units', side_effect=[
                                 LegalUnitError('网络连接超时'), LegalUnits('个', '千克')]) as lookup, \
                             patch('ui.asn_tab.AsnUnitDialog', side_effect=show_dialog) as popup, \
                             patch('ui.asn_tab.write_asn_workbook', return_value=folder / 'ASN.xls') as writer:
                            tab = AsnTab(root)
                            tab.report, tab.customs_data = report, packing
                            tab.directory, tab.customs_path = folder, packing.source
                            tab.output_directory = folder
                            tab.warehouse_var.set(warehouse)
                            tab.generate()
                            def finish_when_ready():
                                if tab.busy:
                                    root.after(25, finish_when_ready)
                                else:
                                    root.quit()
                            root.after(25, finish_when_ready)
                            root.after(10000, root.quit)
                            root.mainloop()
                            self.assertFalse(tab.busy, f'Worker did not resume: {tab.status.cget("text")}')
                            self.assertEqual(dialog_threads, ['MainThread'])
                            popup.assert_called_once()
                            self.assertIn('9503008390', popup.call_args.args)
                            self.assertIn('装箱单总数量', popup.call_args.args[3])
                            self.assertEqual(tab.generate_button.cget('state'), 'normal')
                            self.assertEqual(tab.output_button.cget('state'), 'normal')
                            if action == 'cancel':
                                writer.assert_not_called()
                                self.assertIn('已取消', tab.status.cget('text'))
                            else:
                                writer.assert_called_once()
                                plan = writer.call_args.args[0]
                                self.assertEqual(plan['warehouse'], warehouse)
                                columns = (31, 32, 33, 34) if warehouse == 'yixing' else (17, 18, 19, 20)
                                self.assertEqual([plan['rows'][0][col] for col in columns], [4, '个', 2.7, '千克'])
                            self.assertEqual(lookup.call_count, 2 if action == 'retry' else 1)
                            log = next((folder / 'logs').glob('*.log')).read_text(encoding='utf-8')
                            self.assertIn('网络连接超时', log)
                            if action == 'manual':
                                self.assertIn('用户手动确认（仅本次生成）', log)
                    finally:
                        self.close_root(root)

    def test_closing_app_unblocks_pending_worker(self):
        from ui.asn_tab import AsnTab
        tab = SimpleNamespace(events=Queue(), closed=Event())
        result = []
        worker = Thread(target=lambda: result.append(AsnTab.request_legal_units(tab, '9503002100', 'offline', '')))
        worker.start()
        tab.events.get(timeout=1)
        tab.closed.set()
        worker.join(timeout=1)
        self.assertFalse(worker.is_alive())
        self.assertEqual(result, [None])
