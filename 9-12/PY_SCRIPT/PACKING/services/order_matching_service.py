"""从识别表格提取采购订单号，并匹配包含该订单号的 PDF 文件名。"""
from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata

from models.table import TableData


def normalize_order(value: str) -> str:
    value = unicodedata.normalize("NFKC", str(value)).casefold()
    value = value.translate(str.maketrans({char: "-" for char in "‐‑‒–—−"}))
    return re.sub(r"\s+", "", value)


def _header(value: str) -> str:
    return re.sub(r"[^\w]", "", normalize_order(value)).replace("_", "")


@dataclass
class OrderMatches:
    column: str
    files_by_order: dict[str, list[Path]]
    blank_rows: int


def match_order_pdfs(table: TableData, pdf_files: list[Path]) -> OrderMatches:
    table.validate()
    headers = [_header(value) for value in table.columns]
    # Purchase Order takes priority over Customer PO when both occur in a table.
    priorities = [
        {"purchaseorderno", "purchaseordernumber", "purchaseorder", "采购订单号", "采购单号"},
        {"orderno", "ordernumber", "订单号", "定单号", "pono", "ponumber", "po"},
        {"customerpono", "customerponumber", "客户订单号"},
    ]
    index = next((i for names in priorities for i, name in enumerate(headers) if name in names), None)
    if index is None:
        raise ValueError("未找到订单号列（如 Purchase Order No. / 订单号），无法匹配订单 PDF。")
    orders = {}
    blank_rows = 0
    for row in table.rows:
        value = str(row[index]).strip()
        normalized = normalize_order(value)
        if not normalized:
            blank_rows += 1
            continue
        orders.setdefault(normalized, value)
    files = [(path, normalize_order(path.stem)) for path in dict.fromkeys(pdf_files) if path.suffix.lower() == ".pdf"]
    matches = {}
    for normalized, original in orders.items():
        # Allow filename prefixes/suffixes separated by punctuation, but not a longer ID.
        pattern = re.compile(r"(?<![a-z0-9])" + re.escape(normalized) + r"(?![a-z0-9])")
        matches[original] = [path for path, name in files if pattern.search(name)]
    return OrderMatches(table.columns[index], matches, blank_rows)
