from dataclasses import dataclass, field


@dataclass
class TableData:
    """列名与文本单元格；保留订单编号中的前导零。"""

    columns: list[str] = field(default_factory=list)
    rows: list[list[str]] = field(default_factory=list)

    def validate(self) -> None:
        if not self.columns:
            raise ValueError("表格没有列名")
        if any(len(row) != len(self.columns) for row in self.rows):
            raise ValueError("数据行的列数与表头不一致")
