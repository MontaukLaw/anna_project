from dataclasses import dataclass
from pathlib import Path


@dataclass
class ProductCatalog:
    source: Path
    sheets: dict[str, list[tuple]]
    # An item may have multiple records; retain all of them for later matching.
    items: dict[str, list[dict]]
