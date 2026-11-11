"""Verify each online upload using normal rich clipboard read-back, without data logs."""
from lxml import etree

from clipboard_html import table_fragment
from excel_clipboard import _TableShape, _check_html_errors, inspect_copy
from update_excel import address
from copy_range import column_name


class _MergedShape(_TableShape):
    def __init__(self, max_cells):
        super().__init__(max_cells)
        self.merges = []

    def start(self, tag, attrs):
        super().start(tag, attrs)
        if tag in ('td', 'th') and self.in_table:
            cs, rs = int(attrs.get('colspan', '1')), int(attrs.get('rowspan', '1'))
            if cs > 1 or rs > 1:
                self.merges.append((self.col - cs, self.row, self.col - 1, self.row + rs - 1))


def snapshot(clip, ref, max_cells, *, strict_rows=False):
    """Keep relative merge positions, allowing comparison at a shifted destination."""
    config = {'source': {'max_cells': max_cells}, 'export': {'target': {'mode': 'cells'}}}
    origin = ref.split(':')[0]
    grid = inspect_copy(clip, config, origin, 'selection', expected_range=ref, batch=True,
                        strict_rows=strict_rows)
    target = _MergedShape(max_cells)
    parser = etree.HTMLParser(target=target, no_network=True)
    html = table_fragment(clip['html'])
    for offset in range(0, len(html), 65536):
        parser.feed(html[offset:offset + 65536])
        _check_html_errors(parser)
    parser.close()
    _check_html_errors(parser)
    return dict(rows=grid['rows'], row_count=grid['row_count'], width=grid['width'],
                merges=set(target.merges), unverified_rows=grid.get('unverified_rows', 0))


_SPACES = dict.fromkeys(map(ord, '\u00a0\u3000\u200b\ufeff'), None)


def normalize_text(value):
    """Compare-friendly text: 空白差异与多余的成对外层引号不算数据差异。

    Excel和WPS对同一单元格的纯文本可能只差首尾空格、不间断空格或换行符
    （例如“100”对“100 ”）；WPS导出以引号开头的文本时也可能把外层引号吃掉。
    这些都不应让整批上传失败。"""
    text = str(value).translate(_SPACES).replace('\r\n', '\n').replace('\r', '\n')
    text = '\n'.join(part.strip() for part in text.split('\n')).strip('\n')
    if len(text) >= 2 and text.count('"') == 2 and text.startswith('"') and text.endswith('"'):
        text = text[1:-1]
    return text


def differences(expected, actual, origin, limit=None):
    """List the cells that differ between the source batch and what table D holds.

    limit keeps the scan cheap when only the first difference is needed; without
    a limit the count tells the upload loop whether WPS is still filling the
    batch in (fewer differing cells) or the paste has really gone wrong.
    Blank cells, blank rows and whitespace-only differences are normal data."""
    if (expected['row_count'], expected['width']) != (actual['row_count'], actual['width']):
        return [dict(cell='', kind='shape',
                     expected=f"{expected['row_count']:,}行×{expected['width']:,}列",
                     actual=f"{actual['row_count']:,}行×{actual['width']:,}列")]
    col, row = address(origin)
    found = []
    for y, (wanted, got) in enumerate(zip(expected['rows'], actual['rows'])):
        for x in range(expected['width']):
            a = wanted[x] if x < len(wanted) else ''
            b = got[x] if x < len(got) else ''
            if a == b or normalize_text(a) == normalize_text(b):
                continue
            found.append(dict(cell=f'{column_name(col+x)}{row+y}', kind='value',
                              expected=a, actual=b))
            if limit is not None and len(found) >= limit:
                return found
    return found


def mismatch(expected, actual, origin):
    found = differences(expected, actual, origin, limit=1)
    if not found:
        return None
    item = found[0]
    if item['kind'] == 'shape':
        return f"上传后的行列数与本批内容不一致（期望{item['expected']}，实际{item['actual']}）"
    want = str(item['expected'])[:40]
    got = str(item['actual'])[:40]
    return f"表D单元格{item['cell']}与本批内容不一致（期望“{want}”，实际“{got}”）"


def merge_ref(merge, origin):
    col, row = address(origin)
    left, top, right, bottom = merge
    return f'{column_name(col+left)}{row+top}:{column_name(col+right)}{row+bottom}'


def expected_merges(reference, row, count, columns):
    return {(c1-1, r1-row, c2-1, r2-row) for c1,r1,c2,r2 in reference.get('merges', ())
            if 1 <= c1 <= c2 <= columns and row <= r1 <= r2 < row+count}
