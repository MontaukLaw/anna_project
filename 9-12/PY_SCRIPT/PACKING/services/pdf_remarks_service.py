"""提取订单 PDF 的 REMARKS 及包装指示。"""
import re


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
