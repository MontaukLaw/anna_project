"""Warehouse-specific ASN layouts; business validation is shared."""
from dataclasses import dataclass


@dataclass(frozen=True)
class AsnWarehouse:
    label: str
    sheet: str
    extension: str
    file_format: int
    columns: int
    clear_columns: int
    clear_cells: tuple[str, ...] = ()
    date_cells: tuple[str, ...] = ()
    data_last_row: int | None = None
    extend_print_area: bool = True


WAREHOUSES = {
    'yixing': AsnWarehouse('以星仓', '报关资料与ASN', '.xls', 56, 47, 48,
                           clear_cells=('I24',), date_cells=('C5', 'G13')),
    'zhongtong': AsnWarehouse('中通仓', 'ASN', '.xlsx', 51, 30, 30, date_cells=('D3', 'H12')),
    'xinghui': AsnWarehouse('星辉仓', '报关资料与ASN', '.xlsm', 52, 47, 47,
                           date_cells=('D5', 'H12'), data_last_row=996, extend_print_area=False),
}


def warehouse_profile(key='yixing'):
    try:
        return WAREHOUSES[key]
    except (KeyError, TypeError):
        raise ValueError('请选择以星仓、中通仓或星辉仓') from None
