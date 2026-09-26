"""Ordered source rows for future ASN generation: one source row per ASN row."""
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


@dataclass
class CustomsColumn:
    key: str
    label: str
    column: int


@dataclass
class CustomsPackingRow:
    sheet: str
    source_row: int
    columns: list[CustomsColumn]
    values: dict[str, Any]
    raw_values: dict[str, Any]
    warnings: list[str] = field(default_factory=list)


@dataclass
class CustomsPackingData:
    source: Path
    rows: list[CustomsPackingRow] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
