"""将排期中的 Excel 日期、序列值及明确年月日文本转为日期。"""
from datetime import date, datetime
import re
from openpyxl.utils.datetime import from_excel, WINDOWS_EPOCH
from models.packing import PackingDataError


def parse_schedule_date(value, epoch=WINDOWS_EPOCH):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        try:
            result = from_excel(value, epoch)
            if isinstance(result, datetime):
                return result.date()
        except (ValueError, OverflowError):
            pass
    text = str(value or '').strip()
    match = re.fullmatch(r'(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?(?:\s+00:00:00)?', text)
    if match:
        try:
            return date(*map(int, match.groups()))
        except ValueError:
            pass
    for pattern in ('%b %d, %Y', '%B %d, %Y', '%Y%m%d'):
        try:
            return datetime.strptime(text, pattern).date()
        except ValueError:
            pass
    raise PackingDataError(f'询单要求交期为空或日期格式无法确定：{value}')
