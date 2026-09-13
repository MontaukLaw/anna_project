import unicodedata
from models.table import TableData


def _width(text: str) -> int:
    return sum(2 if unicodedata.east_asian_width(char) in "WF" else 1 for char in text)


def format_table(table: TableData) -> str:
    """等宽纯文本表格，兼顾中文宽度与源单元格的多行文本。"""
    table.validate()
    rows = [[str(value).replace("\r", " ").replace("\n", " ").replace("\t", " ") for value in row]
            for row in [table.columns, *table.rows]]
    widths = [max(_width(row[i]) for row in rows) for i in range(len(table.columns))]
    lines = [" | ".join(value + " " * (widths[i] - _width(value)) for i, value in enumerate(row)) for row in rows]
    lines.insert(1, "-+-".join("-" * width for width in widths))
    return "\n".join(lines)
