"""后台读取订单 PDF 的 REMARKS，附加到识别表格最右侧。"""
import re
from pypdf import PdfReader
from models.table import TableData
from services.order_matching_service import normalize_order


def extract_remarks(pages: list[str]) -> str:
    sections = []
    started = False
    for page in pages:
        heading = re.search(r'\bREMARKS\s*[:：]', page, re.I)
        if heading:
            body = page[heading.end():]
            started = True
        elif started:
            # Repeated purchase-order headers end at the table header on continuation pages.
            header = re.search(r'ITEM NO\.[\s\S]*?PRICE\s*\(USD\)', page, re.I)
            body = page[header.end():] if header else page
        else:
            continue
        footer = re.search(r'Couture Toy Limited\s*Just Play|Just Play \(HK\) Ltd\s*Couture Toy Limited|Accepted By|THIS IS A COMPUTER-GENERATED', body, re.I)
        if footer:
            body = body[:footer.start()]
        text = body.strip()
        if text:
            sections.append(text)
        if footer:
            break
    return '\n'.join(sections)


def filter_packaging_remarks(text: str) -> str:
    """只保留件数/箱数/滑托换算及散箱出货指示。"""
    matches = []
    pattern = r'\b[\d,]+(?:\.\d+)?\s*PCS?\.?\s*=\s*[\d,]+(?:\.\d+)?\s*(?:CTNS?|CNTS?|CARTONS?)\.?\s*=\s*1\s*SLIP(?:\s*SHEETS?)?\b'
    for match in re.finditer(pattern, text, re.I):
        matches.append((match.start(), re.sub(r'\s+', ' ', match.group()).strip()))
    for match in re.finditer(r'\bSHIP\s+LOOSE\s+CARTONS?\b', text, re.I):
        # A negative instruction is not permission to ship loose cartons.
        prefix = text[max(0, match.start()-20):match.start()]
        if not re.search(r'\b(?:NOT|NEVER|NO)\s*$', prefix, re.I):
            matches.append((match.start(), '散箱出货/不打托'))
    return '\n'.join(dict.fromkeys(value for _, value in sorted(matches)))


def read_pdf_remarks(path, order):
    reader = PdfReader(path)
    pages = [page.extract_text() or '' for page in reader.pages]
    numbers = re.findall(r'PURCHASE ORDER NO\.\s*:?\s*(POHK-\d+-\d+-\d+)', '\n'.join(pages), re.I)
    if not numbers or {normalize_order(number) for number in numbers} != {normalize_order(order)}:
        raise ValueError('PDF 内的订单号不匹配或没有可读取文字')
    return filter_packaging_remarks(extract_remarks(pages))


def append_remarks(table, order_column, remarks):
    keep = [i for i, name in enumerate(table.columns) if name.strip().upper() != 'REMARKS']
    order_index = table.columns.index(order_column)
    values = [remarks.get(normalize_order(row[order_index]), '') for row in table.rows]
    show = any(values)
    return TableData([table.columns[i] for i in keep] + (['REMARKS'] if show else []),
        [[row[i] for i in keep] + ([value] if show else []) for row, value in zip(table.rows, values)])


def read_remarks_job(table, order_column, matches, generation, events):
    remarks = {}
    for order, files in matches.items():
        results = []
        if not files:
            events.put(('remarks_log', (generation, f'{order}：未找到匹配 PDF。', 'WARNING')))
        for path in files:
            try:
                content = read_pdf_remarks(path, order)
                if not content:
                    events.put(('remarks_log', (generation, f'{order} / {path.name}：未匹配到指定的包装备注。', 'INFO')))
                else:
                    events.put(('remarks_log', (generation, f'{order} / {path.name}：REMARKS 已读取。', 'INFO')))
                if content:
                    results.append(f'【{path.name}】\n{content}' if len(files) > 1 else content)
            except Exception as exc:
                events.put(('remarks_log', (generation, f'{order} / {path.name}：REMARKS 读取失败：{exc}', 'ERROR')))
        remarks[normalize_order(order)] = '\n\n'.join(results)
    try:
        events.put(('remarks_ready', (generation, append_remarks(table, order_column, remarks))))
    except Exception as exc:
        events.put(('remarks_log', (generation, f'REMARKS 表格更新失败：{exc}', 'ERROR')))
    finally:
        events.put(('remarks_done', generation))
