"""为资料行给出可核对的推荐理由，不代替用户选择。"""
from dataclasses import dataclass
from decimal import Decimal, InvalidOperation
import re
from config import PROJECT_DIR
from services.settings_service import read_settings
from services.packing_display_service import packing_dimensions_cm, packaging_keywords


@dataclass(frozen=True)
class Recommendation:
    level: str
    reason: str


def read_dimension_tolerance():
    path = PROJECT_DIR / 'packing_rules.json'
    settings = read_settings(path)
    try:
        percent = Decimal(str(settings['dimension_tolerance_percent']))
        if not percent.is_finite() or not 0 <= percent <= 100:
            raise ValueError()
    except (KeyError, InvalidOperation, ValueError):
        raise ValueError(f'{path.name} 中 dimension_tolerance_percent 必须为 0 到 100 的数字') from None
    return percent


def recommend_candidate(item, record, tolerance_percent=None):
    if tolerance_percent is None:
        tolerance_percent = read_dimension_tolerance()
    tolerance = Decimal(str(tolerance_percent)) / 100
    values = record['values']
    note = str(values.get('备注') or '')
    required = {'SLIP SHEET' if 'SLIP' in match.group().upper() else ('PDQ' if match.group().upper() == 'PDQ' else 'BAG') for match in packaging_keywords(item.packaging)}
    present = set()
    if re.search(r'(?<![A-Za-z])PDQ(?![A-Za-z])', note, re.I):
        present.add('PDQ')
    if re.search(r'(?<![A-Za-z])(?:POLY\s?BAGS?|BAGS?)(?![A-Za-z])|袋', note, re.I):
        present.add('BAG')
    shared = required & present
    packaging_match = bool(required) and required <= present
    target = packing_dimensions_cm(item.packing)
    dimensions_match = False
    valid_dimensions = False
    try:
        actual = tuple(Decimal(str(values.get(key))) for key in ('长', '宽', '高'))
        valid_dimensions = all(value.is_finite() and value > 0 for value in actual)
        dimensions_match = target is not None and valid_dimensions and all(abs(a - b) <= b * tolerance for a,b in zip(actual,target))
    except (InvalidOperation, ValueError, TypeError):
        pass
    reasons = []
    if dimensions_match:
        reasons.append(f'尺寸一致（误差≤{tolerance_percent}%）')
    if packaging_match:
        reasons.append(f"装箱方式一致（{' / '.join(sorted(required))}）")
    elif shared:
        reasons.append(f"部分包装关键词匹配（{' / '.join(sorted(shared))}）")
    highlighted = dimensions_match or bool(shared)
    if highlighted and not dimensions_match:
        reasons.append('尺寸未匹配' if target and valid_dimensions else '尺寸待核对')
    if not reasons:
        reasons.append('未匹配推荐条件')
    try:
        if Decimal(str(values.get('装箱数量'))) != item.case_pack:
            reasons.append('装箱数量不同，请核对')
    except (InvalidOperation, ValueError, TypeError):
        reasons.append('装箱数量需人工核对')
    level = 'strong' if dimensions_match and packaging_match else ('partial' if highlighted else 'normal')
    return Recommendation(level, '；'.join(reasons))
