import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import Mock, patch

from services.settings_service import read_settings, save_settings


@unittest.skipUnless(os.name == 'nt', 'Requires the Windows Tk desktop runtime')
class AsnUITests(unittest.TestCase):
    def test_home_page_global_fonts_persistence_and_directory_check(self):
        from ui.app import PackingApp
        from ui import theme as t

        def worker(*, target, args, **kwargs):
            return Mock(start=lambda: target(*args))

        def close(app):
            # Tk timers belong to the interpreter; cancel before creating the
            # second application in this same test process.
            for timer in app.tk.call('after', 'info'):
                app.tk.call('after', 'cancel', timer)
            app.destroy()

        with TemporaryDirectory() as temporary:
            settings = Path(temporary) / 'settings.json'
            save_settings(settings, {'ui_font_size': 16, 'unrelated_preference': 'keep me'})
            with patch('ui.app.SETTINGS_FILE', settings), patch('ui.asn_tab.SETTINGS_FILE', settings), \
                 patch('ui.asn_tab.ASN_LOG_DIR', Path(temporary) / 'logs'), \
                 patch('ui.factory_tab.SETTINGS_FILE', settings), patch('ui.factory_tab.FactoryTab.restore'), \
                 patch('ui.app.PackingApp._restore_product_catalog') as restore, \
                 patch('ui.asn_tab.Thread', side_effect=worker):
                app = PackingApp()
                app.withdraw()
                try:
                    app.update_idletasks()
                    self.assertEqual(app.tabs.get(), '入仓预申报ASN生成器')
                    restore.assert_not_called()
                    self.assertEqual(t.get_font_size(), 16)
                    fonts = [app.table_text.cget('font'), app.factory_tab.status.cget('font'),
                             app.asn_tab.status.cget('font'), app.tabs._segmented_button.cget('font')]
                    before = [font.cget('size') for font in fonts]
                    app.font_controls.plus.invoke()
                    self.assertTrue(all(font.cget('size') > size for font, size in zip(fonts, before)))
                    app.font_controls.minus.invoke()
                    self.assertEqual([font.cget('size') for font in fonts], before)
                    app.font_controls.plus.invoke()
                    self.assertEqual(read_settings(settings)['ui_font_size'], 17)
                    self.assertEqual(read_settings(settings)['unrelated_preference'], 'keep me')
                    app.asn_tab.start_check(Path(temporary))
                    self.assertTrue(app.asn_tab.busy)
                    self.assertEqual(app.asn_tab.select_button.cget('state'), 'disabled')
                    app.asn_tab.poll()
                    self.assertFalse(app.asn_tab.busy)
                    self.assertIn('核对未通过', app.asn_tab.status.cget('text'))
                    self.assertTrue(app.asn_tab.logs.textbox.tag_ranges('ERROR'))
                    self.assertEqual(app.asn_tab.select_button.cget('state'), 'normal')
                    sample = Path(__file__).resolve().parents[4] / 'asn'
                    if (sample / '装箱单SCKTY260914.xls').exists():
                        app.asn_tab.start_check(sample)
                        app.asn_tab.poll()
                        self.assertTrue(app.asn_tab.report.passed)
                        self.assertIn('核对通过', app.asn_tab.status.cget('text'))
                        self.assertTrue(app.asn_tab.logs.textbox.tag_ranges('SUCCESS'))
                        self.assertFalse(app.asn_tab.logs.textbox.tag_ranges('ERROR'))
                        visible = app.asn_tab.logs.textbox.get('1.0', 'end').strip()
                        self.assertEqual(len(visible.splitlines()), 1)
                        self.assertNotIn('[SUCCESS]', visible)
                        self.assertNotIn('Sheet1!', visible)
                        self.assertNotIn('必需表头', visible)
                        log_files = list((Path(temporary) / 'logs').glob('*.log'))
                        self.assertEqual(len(log_files), 1)
                        details = log_files[0].read_text(encoding='utf-8')
                        self.assertIn('Sheet1!E4', details)
                        self.assertIn('未找到 形式发票', details)
                        for message, level in app.asn_tab.report.entries:
                            self.assertIn(f'  {message}\n', details)
                    app.tabs.set('自动装箱生成')
                    app._on_tab_change()
                    restore.assert_called_once()
                finally:
                    close(app)
                reopened = PackingApp()
                reopened.withdraw()
                try:
                    self.assertEqual(t.get_font_size(), 17)
                    self.assertEqual(reopened.tabs.get(), '入仓预申报ASN生成器')
                    self.assertEqual(reopened.asn_tab.status.cget('font').cget('size'), round(17 * 17 / 14))
                    t.set_font_size(t.MAX_FONT_SIZE)
                    reopened.font_controls.refresh()
                    self.assertEqual(reopened.font_controls.plus.cget('state'), 'disabled')
                    t.set_font_size(t.MIN_FONT_SIZE)
                    reopened.font_controls.refresh()
                    self.assertEqual(reopened.font_controls.minus.cget('state'), 'disabled')
                finally:
                    close(reopened)
                    t.set_font_size(t.DEFAULT_FONT_SIZE)


class FontPreferenceTests(unittest.TestCase):
    def test_invalid_font_sizes_use_default(self):
        from ui import theme as t
        for value in (None, True, '20', -1, 100, float('nan')):
            self.assertEqual(t.valid_font_size(value), t.DEFAULT_FONT_SIZE)
