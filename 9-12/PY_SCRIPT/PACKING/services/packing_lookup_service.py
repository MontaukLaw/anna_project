"""一次建立资料索引，按订单与 ITEM 查找排期日期码和装箱参数。"""
import re
from decimal import Decimal, InvalidOperation
from models.packing import PackingDataError
from services.date_service import parse_schedule_date


def clean(value):
    return re.sub(r"\s+", "", str(value or "")).upper()


def number(value, label):
    try:
        result = Decimal(str(value).replace(",", "").strip())
    except (InvalidOperation, TypeError, ValueError):
        raise PackingDataError(f"{label}不是有效数值：{value}")
    if not result.is_finite() or result <= 0:
        raise PackingDataError(f"{label}必须大于零：{value}")
    return result


class PackingLookup:
    def __init__(self, catalog, schedule):
        self.date_epoch = schedule.epoch
        self.products = {}
        for key, records in catalog.items.items():
            for token in re.split(r"[/、,，\s]+", key.strip()):
                token = token.lstrip("#").upper()
                if token:
                    self.products.setdefault(token, []).extend(records)
        self.schedules = {}
        for sheet, rows in schedule.sheets.items():
            headers = {}
            for row in rows[:8]:
                for col, value in enumerate(row):
                    if value is not None:
                        headers.setdefault(clean(value), col)
            date_col = headers.get("产品日期码")
            order_col = headers.get("生产订单号")
            if date_col is None or order_col is None:
                continue
            item_col = headers.get("长编号")
            packing_col = headers.get("包装要求")
            delivery_col = headers.get("询单要求交期")
            for row_no, row in enumerate(rows, 1):
                for order in re.findall(r"POHK-\d+-\d+(?:-\d+)?", clean(row[order_col])):
                    self.schedules.setdefault(order, []).append((sheet, row_no, row[date_col],
                        clean(row[item_col]) if item_col is not None else "",
                        row[packing_col] if packing_col is not None else None,
                        row[delivery_col] if delivery_col is not None else None))

    def _schedule_matches(self, order, item):
        order = clean(order)
        records = self.schedules.get(order)
        if not records:
            records = self.schedules.get(re.sub(r"-\d+$", "", order), [])
        base = item.split("-")[0]
        exact = [r for r in records if item in r[3]]
        return exact or [r for r in records if re.search(r"(?<!\d)"+re.escape(base)+r"(?!\d)", r[0])]

    def packaging_hint(self, order, item):
        matches = self._schedule_matches(order, item)
        lines = [f"订单：{order}    ITEM：{item}"]
        if not matches:
            lines.append("未找到该订单与 ITEM 对应的排期记录，请核对装箱资料。")
        else:
            if len(matches) > 1:
                lines.append(f"找到 {len(matches)} 条排期记录，请综合核对以下包装要求：")
            for sheet, row, _, _, requirement, _ in matches:
                text = str(requirement).strip() if requirement is not None else ""
                lines.append(f"【{sheet} · 第 {row} 行】\n{text or '包装要求为空或未找到该列，请自行核对。'}")
        return '\n\n'.join(lines)

    def delivery_date(self, order, item):
        matches = self._schedule_matches(order, item)
        dates = {parse_schedule_date(row[5], self.date_epoch) for row in matches}
        if len(dates) != 1:
            raise PackingDataError(f"{order} / {item} 的询单要求交期缺失或冲突：{sorted(dates)}")
        return next(iter(dates))

    def date_code(self, order, item):
        matches = self._schedule_matches(order, item)
        codes = {str(r[2]).strip() for r in matches if r[2] is not None and str(r[2]).strip()}
        if len(codes) != 1:
            raise PackingDataError(f"{order} / {item} 的产品日期码缺失或冲突：{sorted(codes)}")
        return next(iter(codes))

    def product_candidates(self, item):
        records = self.products.get(item.upper()) or self.products.get(item.split("-")[0], [])
        unique = {(r['sheet'], r['row']): r for r in records}
        if not unique:
            raise PackingDataError(f"产品资料表未找到 ITEM {item}")
        return list(unique.values())
