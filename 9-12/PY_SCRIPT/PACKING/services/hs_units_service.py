"""Read legal units from the exact HS code row on the configured public search page."""
from dataclasses import dataclass
from html.parser import HTMLParser
from http.client import HTTPException
import re
import unicodedata
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

from config import HS_UNIT_CACHE_FILE
from services.hs_units_cache import HSUnitCache


SEARCH_URL = 'https://www.hsbianma.com/search'
# Unknown labels stop generation instead of treating login/error text as a unit.
SUPPORTED_UNITS = set('千克 克 吨 个 件 套 台 只 双 辆 本 张 条 支 副 架 艘 把 根 头 匹 卷 '
                      '米 平方米 立方米 升 毫升 千个 千只 千瓦时'.split())


class LegalUnitError(ValueError):
    pass


class LegalUnitCancelled(Exception):
    """The user cancelled generation while resolving an unavailable lookup."""


@dataclass(frozen=True)
class LegalUnits:
    first: str
    second: str = ''
    url: str = ''
    manual: bool = False


def unit_name(value):
    text = re.sub(r'\s+', '', unicodedata.normalize('NFKC', str(value)))
    return {'kg': '千克', 'kgs': '千克', '公斤': '千克'}.get(text.lower(), text)


class _Tables(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self.table = None
        self.row = None
        self.cell = None
        self.ignored = 0

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'):
            self.ignored += 1
        elif tag == 'table':
            self.table = []
        elif tag == 'tr' and self.table is not None:
            self.row = []
        elif tag in ('td', 'th') and self.row is not None:
            self.cell = []

    def handle_data(self, data):
        if self.cell is not None and not self.ignored:
            self.cell.append(data)

    def handle_endtag(self, tag):
        if tag in ('script', 'style'):
            self.ignored = max(0, self.ignored - 1)
        elif tag in ('td', 'th') and self.cell is not None:
            self.row.append(''.join(self.cell).strip())
            self.cell = None
        elif tag == 'tr' and self.row is not None:
            self.table.append(self.row)
            self.row = None
        elif tag == 'table' and self.table is not None:
            self.tables.append(self.table)
            self.table = None


def parse_legal_units(html, code, url=''):
    parser = _Tables()
    parser.feed(html)
    matches = set()
    for table in parser.tables:
        headers = None
        for row in table:
            normalized = [re.sub(r'\s+', '', value) for value in row]
            if '商品编码' in normalized and '计量单位' in normalized:
                headers = (normalized.index('商品编码'), normalized.index('计量单位'))
                continue
            if headers is None or len(row) <= max(headers):
                continue
            # The website formats 9503002100 as "9503 0021.00".
            candidate = re.sub(r'[\s.]', '', row[headers[0]])
            if candidate == code:
                matches.add(unit_name(row[headers[1]]))
    if len(matches) != 1:
        raise LegalUnitError(f'商品编码 {code} 未查到唯一的计量单位（网页可能变化、无结果或要求验证），请核实后重试。')
    parts = next(iter(matches)).split('/')
    if not 1 <= len(parts) <= 2 or any(unit_name(p) not in SUPPORTED_UNITS for p in parts):
        raise LegalUnitError(f'商品编码 {code} 的计量单位格式无法识别，请核实网页。')
    return LegalUnits(unit_name(parts[0]), unit_name(parts[1]) if len(parts) == 2 else '', url)


def query_legal_units(code, timeout=15):
    code = str(code).strip()
    if not re.fullmatch(r'\d{10}', code):
        raise LegalUnitError(f'商品编码 {code} 须为完整的 10 位数字，不能自动截断或猜测。')
    url = SEARCH_URL + '?' + urlencode({'keywords': code})
    request = Request(url, headers={'User-Agent': 'Mozilla/5.0', 'Accept': 'text/html'})
    try:
        with urlopen(request, timeout=timeout) as response:
            data = response.read(2_000_001)
            if len(data) > 2_000_000:
                raise LegalUnitError(f'商品编码 {code} 的查询响应过大，无法读取。')
            html = data.decode(response.headers.get_content_charset() or 'utf-8')
    except HTTPError as exc:
        raise LegalUnitError(f'商品编码 {code} 的法定单位查询失败（HTTP {exc.code}），请检查网站后重试。') from exc
    except (URLError, OSError, TimeoutError, UnicodeError, LookupError, HTTPException) as exc:
        raise LegalUnitError(f'商品编码 {code} 的法定单位查询失败：网络连接、超时或响应异常。请检查网络后重新生成。') from exc
    return parse_legal_units(html, code, url)


def manual_legal_units(first, second):
    first, second = unit_name(first), unit_name(second)
    if first not in SUPPORTED_UNITS:
        raise LegalUnitError('请选择法1单位。')
    if second and second not in SUPPORTED_UNITS:
        raise LegalUnitError('请选择法2单位，或明确选择“无第二法定单位”。')
    if first == second:
        raise LegalUnitError('法1和法2单位不能相同，请核实后选择。')
    return LegalUnits(first, second, manual=True)


def load_legal_units(codes, log=None, on_failure=None):
    """on_failure returns 'retry', manually confirmed LegalUnits, or None to cancel."""
    result = {}
    cache = HSUnitCache(HS_UNIT_CACHE_FILE, log)
    codes = list(dict.fromkeys(codes))
    if log:
        log(f'正在读取本地缓存：{cache.path}')
    entries = cache.get_many(codes)
    missing = []
    for code in codes:
        url = SEARCH_URL + '?' + urlencode({'keywords': code})
        entry = entries.get(code)
        if entry is not None:
            if (re.fullmatch(r'\d{10}', str(code).strip())
                    and isinstance(entry, dict) and entry.get('status') == 'success'
                    and entry.get('code') == code and entry.get('url') == url
                    and isinstance(entry.get('first'), str) and entry['first'] in SUPPORTED_UNITS
                    and isinstance(entry.get('second'), str)
                    and entry['second'] in SUPPORTED_UNITS | {''}):
                result[code] = LegalUnits(entry['first'], entry['second'], entry['url'])
                if log:
                    log(f'使用本地缓存：商品编码 {code}；法1={entry["first"]}；法2={entry["second"] or "无"}；'
                        f'查询时间={entry.get("queried_at", "未记录")}；来源={url}')
                continue
            cache.warn(f'中商品编码 {code} 的记录无效，将重新联网查询')
        missing.append(code)
    if log:
        log(f'本地缓存检查完成：命中 {len(result)} 个，待联网 {len(missing)} 个。'
            + ('全部使用本地记录，无须联网。' if not missing else '仅查询本地缺失或无效的编码。'))
    for code in missing:
        url = SEARCH_URL + '?' + urlencode({'keywords': code})
        while True:
            if log:
                log(f'正在联网查询商品编码 {code} 的法定单位…')
            try:
                units = query_legal_units(code)
                cache.record(code, url, units=units)
                break
            except LegalUnitError as exc:
                if re.fullmatch(r'\d{10}', str(code).strip()):
                    cache.record(code, url, error=exc)
                if on_failure is None or not re.fullmatch(r'\d{10}', str(code).strip()):
                    raise
                if log:
                    log(f'查询未完成：{exc}')
                choice = on_failure(code, str(exc))
                if choice is None:
                    raise LegalUnitCancelled('用户取消了 ASN 生成。') from exc
                if choice == 'retry':
                    continue
                if not isinstance(choice, LegalUnits):
                    raise LegalUnitError('未获得有效的法定单位选择。') from exc
                units = manual_legal_units(choice.first, choice.second)
                break
        result[code] = units
        if log:
            source = '用户手动确认（仅本次生成）' if units.manual else units.url
            log(f'商品编码 {code}：法1={units.first}；法2={units.second or "无"}；来源={source}')
    return {code: result[code] for code in codes}
