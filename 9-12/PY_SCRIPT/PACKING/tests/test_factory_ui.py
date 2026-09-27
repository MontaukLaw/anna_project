import os
import json
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
            (directory / 'settings.json').write_text(json.dumps({
                'factory_customers_file': str(directory / 'removed_customer_directory')
            }), encoding='utf-8')
            (directory / 'order.pdf').touch()
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
                    self.assertNotIn('customers', tab.path_vars)
                    tab.set_path('product', product)
                    tab.set_path('schedule', schedule)
                    tab.set_path('orders', directory)
                    tab.set_path('output', directory)
                    tab.poll()
                    self.assertEqual(tab.generate_button.cget('state'), 'disabled')
                    tab.pdf_selection.select_all_button.invoke()
                    self.assertEqual(tab.generate_button.cget('state'), 'normal')
                    # A second instance represents the next application launch.
                    reopened = FactoryTab(root)
                    reopened.restore()
                    reopened.poll()
                    self.assertEqual(reopened.catalog.source, product)
                    self.assertEqual(reopened.schedule.source, schedule)
                    self.assertEqual(reopened.path_vars['product'].get(), str(product))
                    self.assertNotIn('customers', reopened.paths)
                    self.assertNotIn('removed_customer_directory', reopened.logs.textbox.get('1.0', 'end'))
                    self.assertEqual(reopened.pdf_selection.selected_files(), ())
                    reopened.pdf_selection.select_all_button.invoke()
                    self.assertEqual(reopened.generate_button.cget('state'), 'normal')
                    reopened.busy = True
                    reopened.update_state()
                    self.assertTrue(all(control.cget('state') == 'disabled' for control in reopened.buttons))
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()

    def test_filter_checkboxes_bulk_selection_and_only_checked_generation(self):
        import customtkinter as ctk
        from ui.factory_tab import FactoryTab
        from decimal import Decimal
        from types import SimpleNamespace

        def worker(*, target, args, **kwargs):
            return Mock(start=lambda: target(*args))

        with TemporaryDirectory() as temp:
            folder = Path(temp)
            files = [folder / 'COUTURE_POHK-26-16759-004_JPLLC_7.13.2026_0.pdf',
                     folder / 'POHK-26-75999-001.PDF', folder / 'other.pdf',
                     folder / '167' / 'other.pdf']
            files[-1].parent.mkdir()
            for path in files:
                path.touch()
            root = ctk.CTk()
            root.withdraw()
            try:
                with patch('ui.factory_tab.SETTINGS_FILE', folder / 'settings.json'), \
                     patch('ui.factory_tab.Thread', side_effect=worker), \
                     patch('ui.factory_tab.collect_factory_rows', return_value=[SimpleNamespace(cartons=Decimal(3))]) as collect, \
                     patch('ui.factory_tab.write_factory_workbook', return_value=folder / '3.xlsx'), \
                     patch('ui.factory_tab.open_in_excel'):
                    tab = FactoryTab(root)
                    self.assertNotIn('customers', tab.path_vars)
                    tab.catalog = tab.schedule = object()
                    for key in ('orders', 'output'):
                        tab.set_path(key, folder, remember=False)
                    tab.poll()
                    table = tab.pdf_selection
                    self.assertEqual(set(table.files), set(files))
                    tab.generate()
                    collect.assert_not_called()
                    table.filter_var.set('167')
                    self.assertEqual(table.visible_files, files[:1])
                    table.select_all_button.invoke()
                    table.filter_var.set('759')
                    self.assertEqual(set(table.visible_files), set(files[:2]))
                    table.entries[files[1]][1].toggle()
                    table.filter_var.set('other')
                    self.assertEqual(set(table.selected_files()), set(files[:2]))
                    self.assertIn('隐藏项已勾选 2 个', table.summary.cget('text'))
                    table.select_none_button.invoke()
                    self.assertEqual(set(table.selected_files()), set(files[:2]))
                    table.filter_var.set('nothing-matches')
                    self.assertEqual(table.visible_files, [])
                    tab.generate()
                    self.assertEqual(set(collect.call_args.kwargs['order_files']), set(files[:2]))
                    self.assertEqual(table.filter_entry.cget('state'), 'disabled')
                    tab.poll()
                    self.assertFalse(tab.busy)
                    table.filter_var.set('')
                    table.select_none_button.invoke()
                    self.assertEqual(table.selected_files(), ())
                    self.assertEqual(tab.generate_button.cget('state'), 'disabled')
                    table.select_all_button.invoke()
                    self.assertEqual(set(table.selected_files()), set(files))
                    tab.set_path('orders', files[-1].parent, remember=False)
                    tab.poll()
                    self.assertEqual(table.files, [files[-1]])
                    self.assertEqual(table.selected_files(), ())
                    self.assertEqual(table.filter_var.get(), '')
            finally:
                for timer in root.tk.call('after', 'info'):
                    root.tk.call('after', 'cancel', timer)
                root.destroy()
