"""Read the first Ultimate Consignee line without confusing adjacent PDF columns."""
import re

from models.packing import PackingDataError


HEADING = re.compile(r'Ultimate\s+Consignee\s*[:：]?', re.I)


def first_consignee_line(pages):
    names = set()
    for page in pages:
        fragments, headings = [], []
        last_show = None

        def position(cm, tm):
            return (tm[4] * cm[0] + tm[5] * cm[2] + cm[4],
                    tm[4] * cm[1] + tm[5] * cm[3] + cm[5])

        def before(operator, operands, cm, tm):
            nonlocal last_show
            if operator in (b'Tj', b'TJ', b"'", b'"'):
                last_show = position(cm, tm)

        def text(value, cm, tm, font, size):
            if not value.strip():
                return
            x, y = position(cm, tm)
            match = HEADING.search(value)
            if match:
                # pypdf can combine a previous column and this label in one callback.
                # The last text-show operator locates the label itself in that case.
                hx, hy = last_show if last_show is not None else (x, y)
                inline = value[match.end():].split('\n')[0].strip(' :：\t')
                headings.append((hx, hy, size, inline))
            else:
                first = next((line.strip(' :：\t') for line in value.splitlines()
                              if line.strip(' :：\t')), '')
                if first:
                    fragments.append((x, y, first))

        page.extract_text(visitor_operand_before=before, visitor_text=text)
        for x, y, size, inline in headings:
            if inline:
                names.add(inline)
                continue
            # Standard PO layout has the label on the left and company on its right.
            same_line = [(fx, value) for fx, fy, value in fragments
                         if fx > x + size and abs(fy - y) <= max(2, size * .35)]
            if same_line:
                names.add(min(same_line)[1])
                continue
            # Also accept a company immediately below a standalone heading.
            below = [(y - fy, fx, value) for fx, fy, value in fragments
                     if x - 2 <= fx <= x + size * 2 and 0 < y - fy <= size * 2.5]
            if below:
                names.add(min(below)[2])
    if len(names) != 1:
        raise PackingDataError(f'PDF Ultimate Consignee 首行缺失或冲突：{sorted(names)}')
    return names.pop()
