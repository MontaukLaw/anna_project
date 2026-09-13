import json
from pathlib import Path
from queue import Queue
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook
from services.product_catalog_service import read_product_catalog, load_product_catalog_job
from services.settings_service import read_settings, get_product_path, save_product_path


class ProductCatalogTests(unittest.TestCase):
    def test_read_and_remember_unicode_path_and_duplicate_items(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "产品资料.xlsx"
            config = root / "settings.json"
            config.write_text(json.dumps({"other_setting": 123}), encoding="utf8")
            book = Workbook()
            book.active.append(["产品资料表"])
            book.active.append(["ITEM NO.", "净重"])
            book.active.append(["000123", 2])
            book.active.append(["000123", 3])
            book.save(source)
            book.close()
            events = Queue()
            load_product_catalog_job(source, config, True, events)
            messages = list(events.queue)
            self.assertEqual([kind for kind, _ in messages], ["product_loaded", "product_remembered", "product_done"])
            catalog = messages[0][1]
            self.assertEqual(len(catalog.items["000123"]), 2)
            self.assertEqual(catalog.items["000123"][0]["values"]["净重"], 2)
            settings = read_settings(config)
            self.assertTrue(settings["jp_product_confirmed"])
            self.assertEqual(settings["other_setting"], 123)
            self.assertEqual(get_product_path(settings, config), source.resolve())

    def test_invalid_workbook_does_not_replace_saved_path(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            config = root / "settings.json"
            save_product_path(config, root / "previous.xlsx")
            previous = config.read_bytes()
            broken = root / "broken.xlsx"
            broken.write_text("not a workbook")
            events = Queue()
            load_product_catalog_job(broken, config, True, events)
            self.assertEqual([kind for kind, _ in events.queue], ["product_error", "product_done"])
            self.assertEqual(config.read_bytes(), previous)

    def test_missing_item_column_is_rejected(self):
        with TemporaryDirectory() as directory:
            source = Path(directory) / "wrong.xlsx"
            book = Workbook()
            book.active.append(["Something else"])
            book.save(source)
            book.close()
            with self.assertRaises(ValueError):
                read_product_catalog(source)
