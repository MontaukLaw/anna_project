"""Parse explicit quantities and convert legal quantities without guessing unit ratios."""
from decimal import Decimal
from fractions import Fraction
import re
import unicodedata

from services.hs_units_service import unit_name


def quantities_with_units(value):
    if not isinstance(value, str):
        return {}
    text = unicodedata.normalize('NFKC', value).replace(',', '')
    if not re.search(r'[\u4e00-\u9fffA-Za-z]', text):
        return {}
    pattern = re.compile(r'([+]?(?:\d+(?:\.\d+)?|\.\d+))\s*([\u4e00-\u9fffA-Za-z]+)')
    quantities = {}
    position = 0
    for match in pattern.finditer(text):
        if text[position:match.start()].strip(' \t\r\n/;；、='):
            raise ValueError(f'总数量无法完整解析：{value}')
        unit = unit_name(match[2])
        if unit in quantities:
            raise ValueError(f'总数量中 {unit} 重复，无法确定：{value}')
        quantities[unit] = Decimal(match[1])
        position = match.end()
    if not quantities or text[position:].strip(' \t\r\n/;；、='):
        raise ValueError(f'总数量无法完整解析：{value}')
    return quantities


def legal_quantity(item, quantity, net_weight, invoice_unit, legal_unit, *, allow_bare_quantity=False):
    legal_unit = unit_name(legal_unit)
    invoice_unit = unit_name(invoice_unit)
    if not legal_unit:
        return None
    # PL weights are explicitly in kilograms, independently of the pricing unit.
    if legal_unit in ('千克', '克', '吨'):
        return net_weight * {'千克': Decimal(1), '克': Decimal(1000), '吨': Decimal('.001')}[legal_unit]
    if legal_unit == invoice_unit:
        return quantity
    explicit = item.quantity_units
    if legal_unit not in explicit:
        if allow_bare_quantity and not explicit:
            return quantity
        raise ValueError(f'{item.name}：法定单位为“{legal_unit}”，计价单位为“{invoice_unit}”，'
                         f'装箱单总数量“{item.quantity_text or item.quantity}”没有对应数量或换算依据，请补充后重试。')
    if item.quantity is None or item.quantity <= 0:
        raise ValueError(f'{item.name}：缺少有效计价总数量，不能分配法定数量')
    fraction = Fraction(quantity) * Fraction(explicit[legal_unit]) / Fraction(item.quantity)
    denominator = fraction.denominator
    for factor in (2, 5):
        while denominator % factor == 0:
            denominator //= factor
    if denominator != 1:
        raise ValueError(f'{item.name}：按 PL 行分配 {legal_unit} 数量无法精确计算，请确认各行法定数量')
    value = Decimal(fraction.numerator) / Decimal(fraction.denominator)
    if legal_unit in ('个', '件', '套', '台', '只', '双', '辆', '本', '张', '条', '支') and value != value.to_integral_value():
        raise ValueError(f'{item.name}：该 PL 行换算结果为 {value}{legal_unit}，不是整数，请确认装箱分配')
    return value
