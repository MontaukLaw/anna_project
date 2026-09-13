"""从 OCR 坐标重建常规列对齐表格，支持无边框和单元格内换行。"""
from bisect import bisect_left
from dataclasses import dataclass
from math import ceil
from statistics import median

from models.table import TableData


@dataclass(frozen=True)
class TextBox:
    left: float
    top: float
    right: float
    bottom: float
    text: str

    @property
    def center_y(self) -> float:
        return (self.top + self.bottom) / 2


def _join(parts: list[TextBox]) -> str:
    result = ""
    for box in sorted(parts, key=lambda box: (round(box.center_y / 3), box.left)):
        text = box.text.strip()
        # A wrapped order number ending in '-' must not acquire an extra space.
        separator = "" if not result or result.endswith(("-", "/")) else " "
        result += separator + text
    return result


def reconstruct_table(boxes: list[TextBox]) -> TableData:
    boxes = [box for box in boxes if box.text.strip()]
    if not boxes:
        raise ValueError("未识别到文字，请选择清晰的表格图片。")
    height = median(box.bottom - box.top for box in boxes)
    # Full-height whitespace gutters separate columns, including right-aligned numbers.
    intervals: list[list[float]] = []
    for box in sorted(boxes, key=lambda box: box.left):
        if intervals and box.left <= intervals[-1][1] + height * 0.12:
            intervals[-1][1] = max(intervals[-1][1], box.right)
        else:
            intervals.append([box.left, box.right])
    if len(intervals) < 2:
        raise ValueError("识别到文字，但未能分离表格列。请裁剪到表格区域并选择更清晰的图片。")
    column_edges = [(a[1] + b[0]) / 2 for a, b in zip(intervals, intervals[1:])]

    def column(box):
        return bisect_left(column_edges, (box.left + box.right) / 2)

    # First group physical text lines, then use populated lines as logical row baselines.
    lines: list[list[TextBox]] = []
    for box in sorted(boxes, key=lambda box: box.center_y):
        if lines and abs(box.center_y - median(b.center_y for b in lines[-1])) <= height * 0.55:
            lines[-1].append(box)
        else:
            lines.append([box])
    minimum_columns = max(2, ceil(len(intervals) * 0.6))
    baselines = [line for line in lines if len({column(box) for box in line}) >= minimum_columns]
    if len(baselines) < 2:
        raise ValueError("未找到完整的表头与数据行。请使用只包含一张表格的正向图片。")
    row_edges = [max(box.bottom for box in line) + height * 0.2 for line in baselines[:-1]]
    cells = [[[] for _ in intervals] for _ in baselines]
    for box in boxes:
        cells[bisect_left(row_edges, box.center_y)][column(box)].append(box)
    rows = [[_join(cell) for cell in row] for row in cells]
    table = TableData(columns=[value or f"列{index + 1}" for index, value in enumerate(rows[0])], rows=rows[1:])
    table.validate()
    return table
