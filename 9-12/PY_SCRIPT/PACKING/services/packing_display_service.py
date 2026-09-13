"""包装提示文字格式化；尺寸换算仅用于对照，不修改资料表尺寸。"""
from decimal import Decimal
import re


def packing_with_cm(packing: str) -> str:
    number = r'\d+(?:\.\d+)?'
    match = re.search(rf'({number})\s*[x×]\s*({number})\s*[x×]\s*({number})\s*(?:INCH(?:ES)?\b|IN\b|")', packing, re.I)
    if not match:
        return packing
    def cm(value):
        text = format(Decimal(value) * Decimal('2.54'), 'f')
        return text.rstrip('0').rstrip('.') if '.' in text else text
    dimensions = ' × '.join(cm(value) for value in match.groups())
    return f'{packing}（{dimensions} CM）'


def packaging_keywords(text: str):
    return list(re.finditer(r'\b(?:PDQ|POLY\s?BAGS?|BAGS?|SLIP\s+SHEET)\b', text, re.I))


def packing_dimensions_cm(packing: str):
    number = r'\d+(?:\.\d+)?'
    match = re.search(rf'({number})\s*[x×]\s*({number})\s*[x×]\s*({number})\s*(INCH(?:ES)?\b|IN\b|"|CM\b)', packing, re.I)
    if not match:
        return None
    factor = Decimal(1) if match[4].upper() == 'CM' else Decimal('2.54')
    dimensions = tuple(Decimal(match[i]) * factor for i in (1, 2, 3))
    return dimensions if all(value > 0 for value in dimensions) else None
