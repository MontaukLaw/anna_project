from pathlib import Path
from PIL import Image, ImageOps


def load_preview(path: Path, max_size: tuple[int, int] = (900, 140)) -> tuple[Image.Image, tuple[int, int]]:
    """验证图片，按比例缩小；关闭源文件，避免占用文件句柄。"""
    with Image.open(path) as source:
        size = source.size
        preview = ImageOps.exif_transpose(source)
        preview.thumbnail(max_size, Image.Resampling.LANCZOS)
        return preview.convert("RGB"), size
