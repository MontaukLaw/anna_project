from dataclasses import dataclass
from decimal import Decimal
from datetime import date


@dataclass
class OrderItem:
    item_no: str
    quantity: Decimal
    case_pack: Decimal
    customer_po: str
    description: str
    cargo_date: str
    packaging: str = ''
    packing: str = ''
    shipping_remarks: str = ''


@dataclass
class PackingRow:
    item: OrderItem
    date_code: str
    net: Decimal
    gross: Decimal
    length: Decimal
    width: Decimal
    height: Decimal
    english_name: str
    source: str
    delivery_date: date | None = None
    cargo_label: str = ''
    slip_note: str = ''
    display_quantity: Decimal | None = None
    display_case_pack: Decimal | None = None


class PackingDataError(ValueError):
    """该订单缺少或存在冲突数据，不生成不完整装箱单。"""
