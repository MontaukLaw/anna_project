"""生成同一货物的单箱与滑托两种包装记录。"""
import re
from decimal import Decimal
from models.packing import PackingDataError, PackingRow
from services.product_weight_service import product_weights
from services.packing_lookup_service import number


def is_slip_item(item):
    return bool(re.search(r'\bSLIP\s+SHEET\b', item.packaging, re.I))


def build_slip_rows(item, candidates, code, delivery_date, hint, choose, log):
    matches = re.findall(r'([\d,]+)\s*(?:ctns?|cnts?|cartons?)\s*=\s*1\s*slip(?:\s+sheet)?', item.shipping_remarks, re.I)
    counts = {Decimal(value.replace(',', '')) for value in matches}
    if len(counts) != 1 or next(iter(counts)) <= 0:
        raise PackingDataError(f'{item.item_no}：REMARKS 缺少或存在冲突的每托箱数')
    per_slip = next(iter(counts))
    cartons = item.quantity / item.case_pack
    if cartons % per_slip:
        raise PackingDataError(f'{item.item_no}：{cartons} 箱不能按每托 {per_slip} 箱整除，请核对')
    result = []
    for pallet, label in [(False, 'cargo'), (True, 'cargo together with slip sheet')]:
        records = [r for r in candidates if (product_weights(r['values'])[2] == '整个卡板') == pallet]
        if not records:
            raise PackingDataError(f'{item.item_no}：缺少{label}资料行')
        if len(records) > 1:
            choice = choose(item, records, f'请选择 {label} 的资料行\n{hint}')
            if choice is None:
                raise PackingDataError('用户取消选择装箱资料')
            record = choice[0]
        else:
            record = records[0]
        v = record['values']
        net, gross, source = product_weights(v)
        row = PackingRow(item, code, number(net, source+'净重'), number(gross, source+'毛重'),
                         number(v.get('长'), '长'), number(v.get('宽'), '宽'), number(v.get('高'), '高'),
                         item.description, f"{record['sheet']} / 行 {record['row']}", delivery_date,
                         cargo_label=label)
        if row.gross < row.net:
            raise PackingDataError(f'{item.item_no}：毛重小于净重')
        if pallet:
            row.display_quantity = cartons
            row.display_case_pack = per_slip
            row.slip_note = f'{item.item_no}: {per_slip} ctns = 1 slip sheet'
        result.append(row)
        log(f"ITEM {item.item_no}：{label} 使用 {row.source}；净重 {net} KG，毛重 {gross} KG")
    log(f'ITEM {item.item_no}：{item.quantity} 件 / 每箱 {item.case_pack} 件 = {cartons} 箱 / 每托 {per_slip} 箱 = {cartons / per_slip} 托')
    return result
