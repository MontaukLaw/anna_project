from pathlib import Path

APP_NAME = "自动装箱生成系统"
PROJECT_DIR = Path(__file__).resolve().parent
WORKSPACE_DIR = PROJECT_DIR.parent.parent
OUTPUT_DIR = PROJECT_DIR / "output"
PACKING_OUTPUT_DIR = OUTPUT_DIR / "packing_lists"
PACKING_TEMPLATE = PROJECT_DIR / "templates" / "packing_list.xlsx"
SETTINGS_FILE = PROJECT_DIR / "settings.json"
DEFAULT_JP_FILE = PROJECT_DIR / "JP产品净重毛重外箱尺寸表20241209.xlsx"
IMAGE_FILE_TYPES = [("图片文件", "*.png *.jpg *.jpeg *.bmp *.webp *.tif *.tiff"), ("所有文件", "*.*")]
