import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from openpyxl import Workbook


@unittest.skipUnless(os.name == 'nt', 'Requires the Windows Tk desktop runtime')
class FactoryUITests(unittest.TestCase):
    def test_saved_files_reload_and_controls_enable_without_reselection(self):
        import customtkinter as ctk
        from ui.factory_tab import FactoryTab

        # Execute preload synchronously to verify event handling without timing races.
        def worker(*, target, args, **kwargs):
            return Mock(start=lambda: target(*args))

        with TemporaryDirectory() as temp:
            directory = Path(temp)
            product = directory / 'chosen_product.xlsx'
            schedule = directory / 'chosen_schedule.xlsx'
            book = Workbook()
            book.active.append(['ITEM NO.', '产品名称'])
            book.active.append(['28483', '振动球'])
            book.save(product)
            book = Workbook()
            book.active.append(['生产订单号', '包装要求'])
            book.active.append(['POHK-26-13545', '包装'])
            book.save(schedule)
            root = ctk.CTk()
            root.withdraw()
            try:
                with patch('ui.factory_tab.SETTINGS_FILE', directory / 'settings.json'), \
                     patch('ui.factory_tab.Thread', side_effect=worker):
                    tab = FactoryTab(root)
                    tab.set_path('product', product)
                    tab.set_path('schedule', schedule)
                    tab.set_path('orders', directory)
                    tab.set_path('customers', directory)
                    tab.set_path('output', directory)
                    tab.poll()
                    self.assertEqual(tab.generate_button.cget('state'), 'normal')
                    # A second instance represents the next application launch.
                    reopened = FactoryTab(root)
                    reopened.restore()
                    reopened.poll()
                    self.assertEqual(reopened.catalog.source, product)
                    self.assertEqual(reopened.schedule.source, schedule)
                    self.assertEqual(reopened.path_vars['product'].get(), str(product))
                    self.assertEqual(reopened.generate_button.cget('state'), 'normal')
                    reopened.busy = True
                    reopened.update_state()
                    self.assertTrue(all(control.cget('state') == 'disabled' for control in reopened.buttons))
            finally:
                root.destroy()
