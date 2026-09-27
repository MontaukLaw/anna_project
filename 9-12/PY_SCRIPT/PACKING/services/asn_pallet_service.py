"""Explicit pallet counts from PL packaging; no guessed carton-to-pallet conversion."""
from decimal import Decimal
import re
import unicodedata


def pallet_counts(row):
    packaging = unicodedata.normalize('NFKC', str(row.values.get('packaging') or '')).strip()
    if not packaging:
        raise ValueError(f'货号 {row.values.get("item")} 的包装说明为空，无法判断是否打板，请核实')
    # A shipment can contain both loose cartons and palletized rows.
    # Do not require a carton-to-pallet conversion for ordinary carton packaging.
    if not re.search(r'卡板|托盘|托板|栈板|棧板|滑片纸|\bpallets?\b|\bslip\s*sheets?\b', packaging, re.I):
        return None, None
    matches = {Decimal(value.replace(',', '')) for value in
               re.findall(r'(?<![\d.\-])([\d,]+(?:\.\d+)?)\s*箱\s*/\s*(?:卡板|托盘|托板)', packaging)}
    if len(matches) != 1:
        raise ValueError(f'货号 {row.values.get("item")} 的每板箱数缺失或不唯一，请核实包装说明')
    per_pallet = next(iter(matches))
    cartons = row.values.get('cartons')
    if (not per_pallet.is_finite() or per_pallet <= 0 or per_pallet != per_pallet.to_integral_value()
            or not isinstance(cartons, Decimal) or not cartons.is_finite() or cartons <= 0):
        raise ValueError(f'货号 {row.values.get("item")} 的板货数量无效')
    pallets = cartons / per_pallet
    if pallets != pallets.to_integral_value():
        raise ValueError(f'货号 {row.values.get("item")} 的箱数 {cartons} 无法按每板 {per_pallet} 箱整除')
    return int(pallets), int(per_pallet)
