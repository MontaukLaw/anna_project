from pathlib import Path
from pypdf import PdfReader


def read_pdf_text(path: Path) -> list[str]:
    """返回文字型 PDF 的各页文本；不支持读取扫描版 PDF。"""
    with path.open("rb") as stream:
        reader = PdfReader(stream)
        return [page.extract_text() or "" for page in reader.pages]
