from pathlib import Path
from pypdf import PdfReader


def read_pdf_text(path: Path) -> list[str]:
    """返回各页文本。扫描版 PDF 需后续接入 OCR，本函数不执行识别。"""
    with path.open("rb") as stream:
        reader = PdfReader(stream)
        return [page.extract_text() or "" for page in reader.pages]
