"""读取和原子保存用户选择的资料路径。"""
import json
from pathlib import Path
from tempfile import NamedTemporaryFile
from threading import Lock
from fnmatch import fnmatch

_settings_lock = Lock()


def read_settings(path: Path) -> dict:
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    if not isinstance(data, dict):
        raise ValueError("配置文件内容必须是 JSON 对象")
    return data


def get_product_path(settings: dict, config_path: Path) -> Path | None:
    return get_saved_path(settings, config_path, "jp_product_file")


def get_saved_path(settings: dict, config_path: Path, key: str) -> Path | None:
    value = settings.get(key)
    if not isinstance(value, str) or not value.strip():
        return None
    path = Path(value)
    return (path if path.is_absolute() else config_path.parent / path).resolve()


def find_local_product_files(directory: Path) -> list[Path]:
    """定位 main.py 旁的资料表，不依赖启动命令的工作目录。"""
    return sorted((path.resolve() for path in directory.iterdir()
                   if path.is_file() and path.name.startswith("JP产品净重毛重外箱尺寸表")
                   and path.suffix.lower() == ".xlsx"), key=lambda path: path.name.casefold())


def save_product_path(config_path: Path, product_path: Path) -> None:
    save_file_path(config_path, product_path, "jp_product")


def find_local_schedule_files(directory: Path) -> list[Path]:
    return sorted((path.resolve() for path in directory.iterdir()
                   if path.is_file() and not path.name.startswith("~$")
                   and fnmatch(path.name.lower(), "*订单分类排期汇*总*.xlsx")),
                  key=lambda path: path.name.casefold())


def save_file_path(config_path: Path, selected_path: Path, prefix: str) -> None:
    # Both Excel preload workers can finish at once; serialize read/modify/write.
    with _settings_lock:
        _save_file_path(config_path, selected_path, prefix)


def _save_file_path(config_path: Path, selected_path: Path, prefix: str) -> None:
    try:
        settings = read_settings(config_path)
    except (ValueError, OSError):
        settings = {}
    settings.update({f"{prefix}_file": str(selected_path.resolve()), f"{prefix}_confirmed": True})
    config_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with NamedTemporaryFile(mode="w", encoding="utf-8", dir=config_path.parent,
                                suffix=".tmp", delete=False) as stream:
            temporary = Path(stream.name)
            json.dump(settings, stream, ensure_ascii=False, indent=2)
            stream.write("\n")
        temporary.replace(config_path)
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)
