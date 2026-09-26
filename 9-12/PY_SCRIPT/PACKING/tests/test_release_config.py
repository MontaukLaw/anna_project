from pathlib import Path
from tempfile import TemporaryDirectory
import runpy
import sys
import unittest
from unittest.mock import patch

from services.settings_service import ensure_settings, read_settings, save_settings


class ReleaseConfigTests(unittest.TestCase):
    def test_missing_config_created_and_existing_preferences_preserved(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / 'settings.json'
            ensure_settings(path, {'ui_font_size': 14})
            self.assertEqual(read_settings(path), {'ui_font_size': 14})
            save_settings(path, {'ui_font_size': 20, 'asn_output_directory': directory})
            ensure_settings(path, {'ui_font_size': 14})
            self.assertEqual(read_settings(path)['ui_font_size'], 20)
            self.assertEqual(read_settings(path)['asn_output_directory'], directory)

    def test_frozen_resources_and_writable_paths_are_separate(self):
        source = Path(__file__).resolve().parents[1] / 'config.py'
        with TemporaryDirectory() as directory, patch.object(sys, 'frozen', True, create=True), \
             patch.object(sys, 'executable', str(Path(directory) / 'LoveAnna.exe')):
            cfg = runpy.run_path(str(source))
            self.assertEqual(cfg['PROJECT_DIR'], Path(directory).resolve())
            self.assertEqual(cfg['SETTINGS_FILE'].parent, cfg['PROJECT_DIR'])
            self.assertEqual(cfg['ASN_LOG_DIR'].parent, cfg['PROJECT_DIR'])
            self.assertEqual(cfg['ASN_TEMPLATE'].parent.parent, cfg['RESOURCE_DIR'])
            self.assertNotEqual(cfg['RESOURCE_DIR'], cfg['PROJECT_DIR'])
