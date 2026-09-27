"""解析 Just Play 采购订单明细；保留完整 ITEM 编号，排除 ASSORTMENT 子项。"""
import re
from decimal import Decimal
from pathlib import Path
from pypdf import PdfReader
from models.packing import OrderItem, PackingDataError
from services.pdf_remarks_service import extract_remarks, filter_packaging_remarks


def read_order_number(path: Path) -> str:
    reader = PdfReader(path)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    numbers = set(re.findall(r"PURCHASE ORDER NO\.\s*:?\s*(POHK-\d+-\d+-\d+)", text, re.I))
    numbers = {number.upper() for number in numbers}
    if len(numbers) != 1:
        raise PackingDataError(f"PDF 未找到唯一采购订单号：{sorted(numbers)}")
    return numbers.pop()


def factory_customer_name(text, path, pages=None):
    """Prefer CUSTOMER NAME; support Ultimate Consignee in the other PO layout."""
    matches = list(re.finditer(r'(?im)\bCUSTOMER[ \t]+NAME\b[ \t]*[:：]?[ \t]*([^\r\n]*)', text))
    if not matches:
        if pages is not None:
            from services.consignee_service import first_consignee_line
            try:
                return first_consignee_line(pages)
            except PackingDataError as exc:
                raise PackingDataError(f'{Path(path).name}：未找到 CUSTOMER NAME；{exc}') from exc
        raise PackingDataError(f'{Path(path).name}：未找到 CUSTOMER NAME 或可读取的 Ultimate Consignee 客户名称')
    names = set()
    for match in matches:
        name = match[1].strip()
        if not name:
            name = next((line.strip() for line in text[match.end():].splitlines() if line.strip()), '')
        # Never turn the next field heading into a customer when the value is absent.
        if not name or re.match(r'(?i)^(?:CUSTOMER\s+(?:NAME|ORDER)|ULTIMATE\s+CONSIGNEE|'
                                r'PURCHASE\s+ORDER|ITEM\s+NO|ADDRESS|TEL\b|FAX\b)', name):
            raise PackingDataError(f'{Path(path).name}：CUSTOMER NAME 后的客户名称为空，请核实订单 PDF')
        names.add(name)
    if len(names) != 1:
        raise PackingDataError(f'{Path(path).name}：CUSTOMER NAME 后未找到唯一客户名称：{sorted(names)}')
    return next(iter(names))


def read_order_items(path: Path, order: str, *, customer_from_name=False) -> tuple[list[OrderItem], str]:
    reader = PdfReader(path)
    pages = [page.extract_text() or "" for page in reader.pages]
    text = "\n".join(pages)
    shipping_remarks = filter_packaging_remarks(extract_remarks(pages))
    numbers = set(re.findall(r"PURCHASE ORDER NO\.\s*:?\s*(POHK-\d+-\d+-\d+)", text, re.I))
    if numbers != {order.upper()}:
        raise PackingDataError(f"PDF 内订单号与 {order} 不一致，或 PDF 无可读取文字：{numbers}")
    # Main item occupies its own line. Secondary carton IDs on quantity lines are excluded.
    starts = list(re.finditer(r"(?m)^\s*(\d{5,}-[A-Z0-9]+-[A-Z0-9]+-\d+-[A-Z0-9]+)\s*$", text))
    if not starts:
        raise PackingDataError("PDF 未找到完整 ITEM NO 明细")
    items = []
    for index, start in enumerate(starts):
        block = text[start.end():starts[index + 1].start() if index + 1 < len(starts) else len(text)]
        qty = re.search(r"([^\n]+?)\s+([\d,]+(?:\.\d+)?)\s+PCS\b", block)
        if not qty:
            raise PackingDataError(f"ITEM {start[1]} 缺少 QTY")
        quantity = Decimal(qty[2].replace(",", ""))
        # Cancelled lines need no packing, customer PO, date or product validation.
        if quantity == 0:
            continue
        pack = re.search(r"(?m)^\s*([\d,]+(?:\.\d+)?)\s*(?:PCS)?\s*/[^\n]*Packing:", block, re.I)
        customer = re.search(r"(?m)^\s*([^\n]+?)Customer Order No\.:\s*$", block)
        cargo = re.search(r"(?m)^\s*([^\n]+?)Cargo Ready Date:\s*$", block)
        if not qty or not pack or not customer:
            raise PackingDataError(f"ITEM {start[1]} 缺少 QTY / Packing / Customer Order No")
        case_pack = Decimal(pack[1].replace(",", ""))
        if quantity <= 0 or case_pack <= 0 or quantity % case_pack:
            raise PackingDataError(f"ITEM {start[1]} 数量或装箱量无效，或不能整除：{quantity}/{case_pack}")
        first_line = next((line.strip() for line in block.splitlines() if line.strip()), '')
        description = re.split(r'\s+[\d,]+(?:\.\d+)?\s+PCS\b', first_line, maxsplit=1)[0].strip()
        if not description:
            raise PackingDataError(f"ITEM {start[1]} 缺少 DESCRIPTION 第一行")
        packaging = re.search(r'(?mi)^\s*([^\n]+?)Packaging:\s*$', block)
        packaging_forward = re.search(r'(?mi)^\s*Packaging:\s*([^\n]+)', block)
        packaging_text = packaging[1].strip() if packaging else (packaging_forward[1].strip() if packaging_forward else '')
        packing_text = re.sub(r'Packing:\s*$', '', pack[0], flags=re.I).strip()
        customer_item = re.search(r'(?mi)^\s*([^\n]+?)Customer Item No\.:\s*$', block)
        items.append(OrderItem(start[1], quantity, case_pack, customer[1].strip(), description,
            cargo[1].strip() if cargo else "", packaging_text, packing_text, shipping_remarks,
            customer_item[1].strip() if customer_item else ''))
    customer_name = re.search(r"CUSTOMER NAME\s*:\s*([^\n]+)", text)
    if customer_from_name:
        name = factory_customer_name(text, path, reader.pages)
    elif customer_name:
        name = customer_name[1].strip()
    elif "Wal-Mart Store Inc." in text:
        name = "Wal-Mart Store Inc."
    else:
        raise PackingDataError("PDF 未找到客户名称")
    return items, name
