from pathlib import Path
import sys

APP_NAME = "自动装箱生成系统"
RESOURCE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else RESOURCE_DIR
WORKSPACE_DIR = PROJECT_DIR if getattr(sys, 'frozen', False) else PROJECT_DIR.parent.parent
OUTPUT_DIR = PROJECT_DIR / "output"
PACKING_OUTPUT_DIR = OUTPUT_DIR / "packing_lists"
PACKING_TEMPLATE = RESOURCE_DIR / "templates" / "packing_list.xlsx"
FACTORY_TEMPLATE = RESOURCE_DIR / "templates" / "factory_notice.xlsx"
SETTINGS_FILE = PROJECT_DIR / "settings.json"
ASN_LOG_DIR = PROJECT_DIR / "logs"
ASN_TEMPLATE = RESOURCE_DIR / 'templates' / 'asn_template.xls'
PACKING_RULES_FILE = PROJECT_DIR / 'packing_rules.json'
DEFAULT_JP_FILE = PROJECT_DIR / "JP产品净重毛重外箱尺寸表20241209.xlsx"
