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
from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string, column_index_from_string, get_column_letter

from services.asn_service import normalized, money_equal
from services.asn_warehouse import warehouse_profile
from services.asn_quantity_service import legal_quantity
from services.hs_units_service import LegalUnits, unit_name
from services.asn_template_preservation import preserve_template_metadata
from services.asn_pallet_service import pallet_counts


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


def output_filename(contract, warehouse='yixing'):
    profile = warehouse_profile(warehouse)
    if not contract or re.search(r'[<>:"/\\|?*\x00-\x1f]', contract) or contract.endswith(('.', ' ')):
        raise ValueError('合同号不能用于文件名，请检查合同号')
    if warehouse == 'yixing':
        return f'ZLC盐田综合保税区仓库 -入仓预申报ASN表格-{contract}.xls'
    return f'{profile.label}-入仓预申报ASN表格-{contract}{profile.extension}'


def contract_brand(specification, item_name):
    """Use the complete, explicit brand field in the contract's declaration elements."""
    brands = {re.sub(r'^品牌\s*[:：]\s*', '', part.strip())
              for part in specification.split('|') if part.strip().endswith('牌')}
    if len(brands) != 1:
        raise ValueError(f'{item_name} 的合同品牌缺失或无法唯一确定，请核实合同申报要素')
    return next(iter(brands))


def build_asn_plan(report, packing, template, options, *, validate_only=False):
    warehouse = options.get('warehouse', 'yixing')
    profile = warehouse_profile(warehouse)
    if not report.passed:
        raise ValueError('三单核对尚未通过，不能生成 ASN')
    if not packing.rows or packing.warnings:
        raise ValueError('海关装箱单尚未读取或有未处理的数据问题')
    contract = options.get('contract')
    if contract not in report.documents:
        raise ValueError('请选择本次 ASN 使用的合同号')
    max_rows = (profile.data_last_row - 25 if profile.data_last_row else
                (65511 if warehouse == 'yixing' else 1048551))
    if len(packing.rows) > max_rows:
        raise ValueError(f'{profile.label}模板最多可填写 {max_rows} 行明细，当前有 {len(packing.rows)} 行')
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
    if warehouse == 'yixing':
        book = xlrd.open_workbook(str(template), on_demand=True)
        try:
            if profile.sheet not in book.sheet_names():
                raise ValueError('以星仓模板缺少报关资料与ASN工作表')
        finally:
            book.release_resources()
        # Fixed template headers stay untouched, including their original data types.
        headers = {'C5': date.today().isoformat(), 'C8': next(iter(customers)),
                   'G12': contract, 'G13': sales.signed_date}
    else:
        # Fixed company/customs data stay in the warehouse's own template.
        book = load_workbook(template, read_only=True)
        try:
            if profile.sheet not in book.sheetnames:
                raise ValueError(f'{profile.label}模板缺少 {profile.sheet} 工作表')
            if warehouse == 'xinghui':
                source_sheet = book[profile.sheet]
                origin = source_sheet['AS26'].value
                production_area = source_sheet['AT26'].value
                if not origin:
                    raise ValueError('星辉仓模板的境内货源地为空，请核实模板')
        finally:
            book.close()
        headers = {'D3': date.today().isoformat(), 'D6': next(iter(customers)),
                   'H11': contract, 'H12': sales.signed_date, 'H14': None,
                   'H15': contract, 'F16': None, 'F19': None}
        if warehouse == 'xinghui':
            headers = {'D5': date.today().isoformat(), 'G8': next(iter(customers)),
                       'H11': contract, 'H12': sales.signed_date, 'H14': None,
                       'H15': contract, 'F19': None,
                       'D16': '否'}
    defaults = {'AI': '千克', 'AR': '中国', 'AS': '达州市', 'AT': 51169, 'U': None}
    po_key = 'customer_order'
    totals = defaultdict(lambda: defaultdict(Decimal))
    rows = []
    formulas = []
    legal_totals = defaultdict(lambda: defaultdict(Decimal))
    legal_log = []
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
        for key in ('so', 'item', po_key) + (('customer_item',) if warehouse in ('zhongtong', 'xinghui') else ()):
            if not str(values.get(key) or '').strip():
                raise ValueError(f'海关装箱单第 {source.source_row} 行 {key} 未填写')
        units = options.get('legal_units', {}).get(item.code)
        if not isinstance(units, LegalUnits) or not units.first:
            if not validate_only:
                raise ValueError(f'商品编码 {item.code} 尚未成功查询法定单位，请检查网络后重新生成')
            first_quantity = second_quantity = None
            first_unit = second_unit = ''
        else:
            first_unit, second_unit = units.first, units.second
            first_quantity = legal_quantity(item, values['quantity'], values['net_weight'], unit,
                                            first_unit, allow_bare_quantity=True)
            second_quantity = legal_quantity(item, values['quantity'], values['net_weight'], unit,
                                             second_unit, allow_bare_quantity=False)
            for legal_unit, value in ((first_unit, first_quantity), (second_unit, second_quantity)):
                if legal_unit:
                    legal_totals[mapped][unit_name(legal_unit)] += value
            legal_log.append(f'{source.sheet} 第 {source.source_row} 行 / {item.code} / {item.name}：'
                             f'装箱单总数量={item.quantity_text or item.quantity}；PL计价数量={values["quantity"]}{unit}；'
                             f'法1={first_quantity}{first_unit}；法2={second_quantity if second_unit else "无"}{second_unit}')
        row = [None] * profile.columns
        fields = dict(defaults)
        fields.update({'A': mapped + 1, 'B': values['so'], 'E': values[po_key], 'F': values['item'],
            'J': values['cartons'], 'K': values['per_carton'], 'L': values['quantity'],
            'P': values['length'], 'Q': values['width'], 'R': values['height'], 'S': values['volume'], 'T': 'CTN',
            'W': item.code, 'Y': item.name, 'Z': next(iter(specifications)), 'AA': values['quantity'],
            'AC': item.price, 'AE': 'USD', 'AF': first_quantity,
            'M': unit, 'AB': unit, 'AG': first_unit, 'AH': second_quantity, 'AI': second_unit,
            'AJ': values['gross_weight'], 'AK': values['net_weight'], 'AL': item.country, 'AP': 'kg', 'AQ': 'CBM'})
        if warehouse == 'zhongtong':
            fields = {'A': mapped + 1, 'B': values['so'], 'C': values[po_key], 'D': values['customer_item'],
                      'E': values['quantity'], 'F': values['per_carton'], 'G': values['cartons'],
                      'H': values['gross_weight'], 'I': values['net_weight'], 'J': item.code,
                      'K': item.name, 'L': next(iter(specifications)), 'M': values['quantity'],
                      'N': unit, 'O': item.price, 'Q': 'USD', 'R': first_quantity, 'S': first_unit,
                      'T': second_quantity, 'U': second_unit, 'V': '中国', 'W': item.country, 'X': 51169}
        elif warehouse == 'xinghui':
            specification = next(iter(specifications))
            fields = {'A': mapped + 1, 'B': values['so'], 'D': values[po_key], 'E': values['customer_item'],
                      'I': values['quantity'], 'J': unit, 'K': values['per_carton'], 'L': values['cartons'],
                      'M': values['gross_weight'], 'N': values['net_weight'], 'O': item.code,
                      'P': item.name, 'Q': specification, 'R': values['quantity'], 'S': unit,
                      'T': item.price, 'V': 'USD', 'W': first_quantity, 'X': first_unit,
                      'Y': second_quantity, 'Z': second_unit, 'AA': contract_brand(specification, item.name),
                      'AB': item.country, 'AF': values['length'], 'AG': values['width'], 'AH': values['height'],
                      'AI': 'CTN', 'AJ': values['volume'], 'AM': '中国', 'AN': '千克', 'AO': 'CBM',
                      'AS': origin, 'AT': production_area}
            fields['AD'], fields['AE'] = pallet_counts(source)
            if fields['AD'] is not None:
                headers['D16'] = '是'
        for column, value in fields.items():
            row[column_index_from_string(column) - 1] = float(value) if isinstance(value, Decimal) else value
        rows.append(row)
        excel_row = index + 26
        amount_col, quantity_col, price_col = {'yixing': ('AD', 'AA', 'AC'),
                                               'zhongtong': ('P', 'M', 'O'),
                                               'xinghui': ('U', 'R', 'T')}[warehouse]
        formulas.append({'cell': f'{amount_col}{excel_row}',
                         'value': f'=ROUND({quantity_col}{excel_row}*{price_col}{excel_row},2)'})
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
        for unit, quantity in legal_totals[index].items():
            if unit in item.quantity_units:
                expected = item.quantity_units[unit]
                same = money_equal(quantity, expected) if unit in ('千克', '克', '吨') else quantity == expected
                if not same:
                    raise ValueError(f'{item.name} 法定数量合计不一致：ASN={quantity}{unit}；装箱单={expected}{unit}')
    if warehouse == 'zhongtong':
        headers['H24'] = int(sum((row.values['cartons'] for row in packing.rows), Decimal(0)))
    return {'template': str(Path(template).resolve()), 'sheet': profile.sheet, 'scope': 'all',
            'warehouse': warehouse, 'consignee': next(iter(customers)), 'legal_log': legal_log,
            'validated_legal_units': not validate_only,
            'contract': contract, 'headers': headers, 'rows': rows, 'formulas': formulas}


def write_asn_workbook(plan, directory):
    if plan.get('validated_legal_units') is False:
        raise ValueError('当前仅完成预检查，尚未联网查询法定单位，不能保存 ASN')
    warehouse = plan.get('warehouse', 'yixing')
    profile = warehouse_profile(warehouse)
    if plan['sheet'] != profile.sheet:
        raise ValueError('ASN 工作表与所选仓库不一致，请重新生成')
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True)
    destination = directory / output_filename(plan['contract'], warehouse)
    if destination.exists():
        raise FileExistsError(f'文件已存在，未覆盖：{destination}')
    with TemporaryDirectory(prefix='asn-', dir=directory) as temporary:
        temp = Path(temporary)
        generated = temp / destination.name
        payload = dict(plan, destination=str(generated.resolve()), columns=profile.columns,
                       last_column=get_column_letter(profile.columns),
                       clear_last_column=get_column_letter(profile.clear_columns),
                       clear_cells=list(profile.clear_cells), date_cells=list(profile.date_cells),
                       file_format=profile.file_format, data_last_row=profile.data_last_row,
                       extend_print_area=profile.extend_print_area)
        config = temp / 'plan.json'
        config.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        script = Path(__file__).with_name('asn_excel_writer.ps1')
        result = subprocess.run(['powershell.exe', '-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass',
                                 '-File', str(script), '-PlanPath', str(config)],
                                capture_output=True, text=True, encoding='utf-8', errors='replace',
                                creationflags=subprocess.CREATE_NO_WINDOW)
        if result.returncode != 0 or not generated.exists():
            raise RuntimeError(result.stderr.strip() or result.stdout.strip() or 'Microsoft Excel 未能保存 ASN')
        if profile.extension in ('.xlsx', '.xlsm'):
            preserve_template_metadata(plan['template'], generated, preserve_vba=profile.extension == '.xlsm')
        # Rename only after Excel has closed the complete BIFF8 file. Never overwrite an existing result.
        generated.rename(destination)
    return destination
