"""Build an ASN write plan from reconciled sources and explicitly confirmed field choices."""
from collections import defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP
import json
from pathlib import Path
import re
import subprocess
from tempfile import TemporaryDirectory

import xlrd
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string

from services.asn_service import normalized, money_equal


ASN_SHEET = '报关资料与ASN'
HEADER_FIELDS = {
    'C5': '申报日期（YYYY-MM-DD）', 'E5': '联系人', 'G5': '联系人电话',
    'C6': '境内发货人', 'F6': '发货人统一社会信用代码', 'H6': '发货人海关代码',
    'C7': '发货人电话', 'E7': '发货人传真', 'G7': '发货人地址',
    'F8': '境外收货人代码', 'C9': '生产销售单位', 'F9': '销售单位信用代码', 'H9': '销售单位海关代码',
    'C10': '销售合同买方', 'F10': '买方地址', 'C11': '买方电话', 'F11': '买方传真',
    'C12': '出境关别', 'E12': '征免性质', 'C13': '海关代码', 'E13': '贸易方式',
    'C14': '手册备案号', 'E14': '成交方式', 'G14': '合同签约地点',
    'C15': '许可证号', 'E15': '征免方式', 'G15': '装运期',
    'C16': '贸易国', 'E16': '离境口岸', 'G16': '发票号',
    'C17': '集装箱号', 'E17': '运输工具', 'G17': '提运单号',
    'C18': '指运港', 'E18': '包装种类代码', 'C19': '运输方式', 'E19': '是否退税', 'G19': '随附单据',
    'C20': '运费', 'C21': '保险费', 'C22': '杂费', 'G20': '标记唛头及备注',
    'C23': '特殊关系确认', 'E23': '价格影响确认', 'G23': '支付特许权使用费确认', 'H23': '其他确认',
}


def template_defaults(template):
    book = xlrd.open_workbook(str(template), on_demand=True)
    try:
        sheet = book.sheet_by_name(ASN_SHEET)
        def value(address):
            col, row = coordinate_from_string(address)
            result = sheet.cell_value(row - 1, column_index_from_string(col) - 1)
            if isinstance(result, float) and result.is_integer():
                return str(int(result))
            return str(result).strip()
        return {cell: value(cell) for cell in HEADER_FIELDS}
    finally:
        book.release_resources()


def product_key(row):
    return json.dumps([str(row.values.get(key) or '') for key in ('item', 'chinese_name', 'price')], ensure_ascii=False)


def matching_products(document, row):
    """Match the complete Chinese name (or final / component), never a fuzzy substring."""
    name = normalized(row.values.get('chinese_name') or '')
    return [i for i, item in enumerate(document.items)
            if name and name in (normalized(item.name), normalized(item.name.rsplit('/', 1)[-1]))
            and item.price == row.values.get('price')]


def generation_options(report, packing):
    contract = next(iter(report.documents)) if len(report.documents) == 1 else None
    mapping = {}
    if contract:
        for row in packing.rows:
            candidates = matching_products(report.documents[contract]['装箱单'], row)
            if len(candidates) == 1:
                mapping[product_key(row)] = candidates[0]
    return {'contract': contract, 'product_mapping': mapping}


def output_filename(contract):
    if not contract or re.search(r'[<>:"/\\|?*\x00-\x1f]', contract) or contract.endswith(('.', ' ')):
        raise ValueError('合同号不能用于文件名，请检查合同号')
    return f'ZLC盐田综合保税区仓库 -入仓预申报ASN表格-{contract}.xls'


def build_asn_plan(report, packing, template, options):
    if not report.passed:
        raise ValueError('三单核对尚未通过，不能生成 ASN')
    if not packing.rows or packing.warnings:
        raise ValueError('海关装箱单尚未读取或有未处理的数据问题')
    contract = options.get('contract')
    if contract not in report.documents:
        raise ValueError('请选择本次 ASN 使用的合同号')
    if len(packing.rows) > 65511:
        raise ValueError('.xls 最多支持 65536 行，当前明细超过模板可用行数')
    docs = report.documents[contract]
    declaration = docs['装箱单']
    current_items = [(item.name, str(item.price), item.code) for item in declaration.items]
    if 'source_items' in options and options['source_items'] != current_items:
        raise ValueError('三单商品信息已发生变化，请重新核对并确认商品对应关系')
    sales = docs['香港合同']
    invoice = docs['形式发票']
    # The consignee comes from PL Customer; other company/customs details stay as in the sample.
    customers = set()
    for source in packing.rows:
        customer = str(source.values.get('customer') or '').strip()
        if not customer:
            raise ValueError(f'海关装箱单 {source.sheet} 第 {source.source_row} 行 Customer 为空，无法填写境外收货人')
        customers.add(customer)
    if len(customers) != 1:
        raise ValueError('海关装箱单存在多个不同的 Customer，无法确定本份 ASN 的境外收货人，请核实或拆分装箱单')
    headers = template_defaults(template)
    headers.update({'C5': date.today().isoformat(), 'C8': next(iter(customers)),
                    'G12': contract, 'G13': sales.signed_date})
    defaults = {'AI': '千克', 'AR': '中国', 'AS': '达州市', 'AT': '51169', 'U': None}
    po_key = 'customer_order'
    totals = defaultdict(lambda: defaultdict(Decimal))
    rows = []
    formulas = []
    for index, source in enumerate(packing.rows):
        values = source.values
        mapped = options.get('product_mapping', {}).get(product_key(source))
        if not isinstance(mapped, int) or isinstance(mapped, bool) or not 0 <= mapped < len(declaration.items):
            raise ValueError(f'请确认货号 {values.get("item")} 对应的三单商品')
        item = declaration.items[mapped]
        candidates = [entry for entry in sales.items if normalized(entry.name) == normalized(item.name) and entry.price == item.price]
        specifications = {entry.specification for entry in candidates if entry.specification}
        if len(specifications) != 1:
            raise ValueError(f'{item.name} 的申报要素缺失或存在多个版本，请先核实合同规格')
        invoice_items = [entry for entry in invoice.items
                         if normalized(entry.name) == normalized(item.name) and entry.price == item.price]
        units = {entry.unit.strip() for entry in invoice_items}
        if len(units) != 1 or not next(iter(units)):
            raise ValueError(f'{item.name} 的形式发票单位缺失或不唯一，请核实，不能沿用模板单位')
        unit = next(iter(units))
        for key in ('quantity', 'cartons', 'per_carton', 'length', 'width', 'height', 'volume', 'net_weight', 'gross_weight', 'price'):
            value = values.get(key)
            if not isinstance(value, Decimal) or not value.is_finite() or value < 0:
                raise ValueError(f'海关装箱单 {source.sheet} 第 {source.source_row} 行 {key} 缺失或无效')
        if values['cartons'] <= 0 or values['cartons'] != values['cartons'].to_integral_value():
            raise ValueError(f'海关装箱单第 {source.source_row} 行箱数须为正整数')
        if values['quantity'] != values['cartons'] * values['per_carton']:
            raise ValueError(f'海关装箱单第 {source.source_row} 行数量与箱数×每箱数量不一致')
        if values['price'] != item.price:
            raise ValueError(f'货号 {values.get("item")}：PL单价 {values["price"]} 与三单单价 {item.price} 不一致')
        for key in ('so', 'item', po_key):
            if not str(values.get(key) or '').strip():
                raise ValueError(f'海关装箱单第 {source.source_row} 行 {key} 未填写')
        row = [None] * 47
        fields = dict(defaults)
        fields.update({'A': mapped + 1, 'B': values['so'], 'E': values[po_key], 'F': values['item'],
            'J': values['cartons'], 'K': values['per_carton'], 'L': values['quantity'],
            'P': values['length'], 'Q': values['width'], 'R': values['height'], 'S': values['volume'], 'T': 'CTN',
            'W': item.code, 'Y': item.name, 'Z': next(iter(specifications)), 'AA': values['quantity'],
            'AC': item.price, 'AE': 'USD', 'AF': values['quantity'],
            'M': unit, 'AB': unit, 'AG': unit, 'AH': values['net_weight'],
            'AJ': values['gross_weight'], 'AK': values['net_weight'], 'AL': item.country, 'AP': 'kg', 'AQ': 'CBM'})
        for column, value in fields.items():
            row[column_index_from_string(column) - 1] = float(value) if isinstance(value, Decimal) else value
        rows.append(row)
        excel_row = index + 26
        formulas.append({'cell': f'AD{excel_row}', 'value': f'=ROUND(AA{excel_row}*AC{excel_row},2)'})
        for key in ('quantity', 'cartons', 'net_weight', 'gross_weight'):
            totals[mapped][key] += values[key]
        totals[mapped]['amount'] += (values['quantity'] * item.price).quantize(Decimal('.01'), rounding=ROUND_HALF_UP)
    for index, item in enumerate(declaration.items):
        actual = totals[index]
        for key in ('quantity', 'cartons', 'net_weight', 'gross_weight', 'amount'):
            expected = getattr(item, key)
            if expected is None:
                continue
            matches = actual[key] == expected if key in ('quantity', 'cartons') else money_equal(actual[key], expected)
            if not matches:
                raise ValueError(f'{item.name} {key} 汇总不一致：海关装箱单={actual[key]}；三单={expected}')
    return {'template': str(Path(template).resolve()), 'sheet': ASN_SHEET, 'scope': 'asn_only',
            'contract': contract, 'headers': headers, 'rows': rows, 'formulas': formulas}


def write_asn_workbook(plan, directory):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    destination = directory / output_filename(plan['contract'])
    if destination.exists():
        raise FileExistsError(f'文件已存在，未覆盖：{destination}')
    with TemporaryDirectory(prefix='asn-', dir=directory) as temporary:
        temp = Path(temporary)
        generated = temp / destination.name
        payload = dict(plan, destination=str(generated.resolve()))
        config = temp / 'plan.json'
        config.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        script = Path(__file__).with_name('asn_excel_writer.ps1')
        result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                                 '-File', str(script), '-PlanPath', str(config)],
                                capture_output=True, text=True, encoding='utf-8', errors='replace',
                                creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode != 0 or not generated.exists():
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or 'Microsoft Excel 未能保存 ASN')
        # Rename only after Excel has closed the complete BIFF8 file. Never overwrite an existing result.
        generated.rename(destination)
    return destination
