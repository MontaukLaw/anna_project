"""Combine all purchase orders into one factory notice, with explicit user choices."""
from dataclasses import dataclass
from decimal import Decimal
from pathlib import Path
import os
import re

from pypdf import PdfReader

from models.packing import OrderItem, PackingDataError
from services.order_pdf_service import read_order_items, read_order_number
from services.packing_lookup_service import PackingLookup, clean, number
from services.pdf_remarks_service import extract_remarks


@dataclass
class FactoryRow:
    order: str
    customer: str
    item: OrderItem
    product: dict
    packaging: str
    booking: str
    cartons_per_slip: Decimal | None

    @property
    def cartons(self):
        return self.item.quantity / self.item.case_pack


def scan_pdfs(directory):
    """Fail on inaccessible subdirectories rather than silently omit an order."""
    directory = Path(directory)
    if not directory.is_dir():
        raise PackingDataError(f"目录不存在：{directory}")
    result = []

    def fail(error):
        raise error

    for root, folders, files in os.walk(directory, onerror=fail, followlinks=False):
        folders.sort(key=str.casefold)
        result.extend(Path(root) / name for name in sorted(files, key=str.casefold)
                      if Path(name).suffix.lower() == '.pdf')
    return result


class FactoryScheduleLookup:
    """Packaging lookup does not depend on date-code or delivery-date columns."""
    def __init__(self, schedule):
        self.orders = {}
        for sheet, rows in schedule.sheets.items():
            headers = {}
            for row in rows[:8]:
                for col, value in enumerate(row):
                    if value is not None:
                        headers.setdefault(clean(value), col)
            order_col, pack_col = headers.get('生产订单号'), headers.get('包装要求')
            if order_col is None or pack_col is None:
                continue
            item_col = next((headers[k] for k in ('长编号', 'ITEMNO.', 'ITEMNO', '货号', '产品编号') if k in headers), None)
            for row_no, row in enumerate(rows, 1):
                if max(order_col, pack_col) >= len(row):
                    continue
                item = clean(row[item_col]) if item_col is not None and item_col < len(row) else ''
                for order in re.findall(r'POHK-\d+-\d+(?:-\d+)?', clean(row[order_col])):
                    self.orders.setdefault(order, []).append({
                        'sheet': sheet, 'row': row_no, 'item': item,
                        'packaging': str(row[pack_col] or '').strip(),
                    })

    def matches(self, order, item):
        order, item = clean(order), clean(item)
        records = self.orders.get(order)
        if records is None:
            records = self.orders.get(re.sub(r'-\d+$', '', order), [])
        exact = [r for r in records if item in re.split(r'[/、,，\s]+', r['item'])]
        if exact:
            return exact
        base = item.split('-')[0]
        # Never match a different full item merely because the sheet has the same base.
        return [r for r in records if r['item'] == base or
                (not r['item'] and re.search(r'(?<!\d)' + re.escape(base) + r'(?!\d)', r['sheet']))]

    def packaging(self, order, item, choose):
        matches = self.matches(order, item.item_no)
        if not matches or any(not r['packaging'] for r in matches):
            raise PackingDataError(f'{order} / {item.item_no}：排期表未找到包装要求或包含空白，请补齐')
        options = {}
        for record in matches:
            options.setdefault(record['packaging'], []).append(record)
        if len(options) == 1:
            return next(iter(options))
        labels = [f"{value}\n来源：" + '、'.join(f"{r['sheet']} 第 {r['row']} 行" for r in records)
                  for value, records in options.items()]
        selected = choose('包装要求冲突', f'{order} / {item.item_no}\nPDF Packaging：{item.packaging}', labels)
        if selected is None:
            raise PackingDataError('已取消本次合并生成')
        return list(options)[selected]


def slip_cartons(item, remarks):
    if not re.search(r'\bSLIP\s+SHEETS?\b', item.packaging, re.I):
        return None
    base = item.item_no.split('-')[0]
    pattern = (r'(?<!\d)' + re.escape(base) +
               r'\s*[-–—:]\s*[\d,.]+\s*pcs?\.?\s*=\s*([\d,]+(?:\.\d+)?)\s*'
               r'(?:ctns?|cnts?|cartons?)\.?\s*=\s*1\s*slip\s*sheet\b')
    counts = {Decimal(v.replace(',', '')) for v in re.findall(pattern, remarks, re.I)}
    if len(counts) != 1 or next(iter(counts)) <= 0:
        raise PackingDataError(f'{item.item_no}：REMARKS 中每托箱数缺失或冲突')
    count = next(iter(counts))
    if (item.quantity / item.case_pack) % count:
        raise PackingDataError(f'{item.item_no}：箱数不能按每托 {count} 箱整除，请核对')
    return count


def collect_factory_rows(pdf_directory, catalog, schedule,
                         choose_product, choose_option, log=lambda message: None, *, order_files=None):
    # A supplied selection is a snapshot: never rescan and add unchecked/new files.
    paths = scan_pdfs(pdf_directory) if order_files is None else list(dict.fromkeys(map(Path, order_files)))
    if not paths:
        raise PackingDataError('请至少勾选一个订单 PDF' if order_files is not None else '订单目录中没有 PDF 文件')
    if order_files is not None:
        directory = Path(pdf_directory).resolve()
        for path in paths:
            if not path.resolve().is_relative_to(directory) or path.suffix.lower() != '.pdf' or not path.is_file():
                raise PackingDataError(f'已勾选的订单 PDF 不存在或不在所选目录中，请重新选择目录：{path}')
    lookup = PackingLookup(catalog, schedule)
    packaging = FactoryScheduleLookup(schedule)
    rows, seen = [], {}
    for path in paths:
        log(f'读取订单 PDF：{path}')
        order = read_order_number(path)
        if order in seen:
            raise PackingDataError(f'采购订单 {order} 重复：{seen[order]} / {path}。请保留需要使用的一份，避免重复计数')
        seen[order] = path
        items, customer = read_order_items(path, order, customer_from_name=True)
        pages = [page.extract_text() or '' for page in PdfReader(path).pages]
        remarks = extract_remarks(pages)
        if not items:
            log(f'{order}：没有非零数量明细，跳过')
        for item in items:
            requirement = packaging.packaging(order, item, choose_option)
            candidates = lookup.product_candidates(item.item_no)
            if len(candidates) > 1:
                selected = choose_product(item, candidates, f'订单：{order}\n包装要求：{requirement}')
                if selected is None:
                    raise PackingDataError('已取消本次合并生成')
                record = selected[0]
            else:
                record = candidates[0]
            values = record['values']
            for key in ('长', '宽', '高', '整箱净重kg', '整箱毛重kg'):
                number(values.get(key), f'{order} / {item.item_no} {key}')
            if number(values['整箱毛重kg'], '毛重') < number(values['整箱净重kg'], '净重'):
                raise PackingDataError(f'{item.item_no}：整箱毛重小于净重')
            if not str(values.get('产品名称') or '').strip():
                raise PackingDataError(f'{item.item_no}：所选资料缺少产品名称')
            if str(values.get('装箱数量')) != str(item.case_pack):
                log(f'提示：{item.item_no} 资料装箱数量 {values.get("装箱数量")}；箱数仍按 PDF 每箱 {item.case_pack} 件计算')
            row = FactoryRow(order, customer, item, record, requirement, '',
                             slip_cartons(item, remarks))
            rows.append(row)
            log(f'{order} / {item.item_no}：{row.cartons} 箱，'
                f'卡板 {row.cartons / row.cartons_per_slip if row.cartons_per_slip else "无"}；'
                f'资料 {record["sheet"]} 第 {record["row"]} 行')
    if not rows:
        raise PackingDataError('没有可生成的非零订单明细')
    return rows
