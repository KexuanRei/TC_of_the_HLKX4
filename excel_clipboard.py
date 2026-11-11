"""Paste the native rich clipboard into a temporary Excel workbook, then commit."""
import datetime as dt
import hashlib
import io
import os
import posixpath
import re
from pathlib import Path
import tempfile
import time
import zipfile
from lxml import etree

from file_source import source_path
from format_reference import load as load_format_reference
from paste_source import parse, PasteShapeError
from update_excel import REL, address, column, tag, xml
from read_limits import check_cells
from copy_range import column_name
from clipboard_html import CopyIntegrityError, decode_context, table_fragment
from omitted_columns import check_omitted_target, omitted_in_range, expand_values, column_list


def _check_writable_target(path):
    """Fail fast when 表A1 cannot be replaced, before minutes of copying.

    The common cause is the workbook being open in WPS/Excel, which blocks the
    atomic replace at the end with PermissionError (reported as FILE_ACCESS_DENIED)."""
    try:
        with path.open('r+b'):
            pass
    except PermissionError as exc:
        raise ValueError('表A1正被其他程序占用，无法写入（通常是WPS/Excel已打开表A1，'
                         '或云同步、杀毒软件正在锁定该文件）。请先关闭表A1后重新复制；本次没有修改表A1') from exc
    except FileNotFoundError as exc:
        raise ValueError('找不到表A1文件：' + str(path) + '；请检查配置中的表A1路径') from exc
    except OSError as exc:
        raise ValueError('表A1无法打开写入（' + type(exc).__name__ + '）：' + str(exc)[:120] +
                         '；请检查文件属性、所在文件夹权限后重试，本次没有修改表A1') from exc
    if not os.access(path, os.W_OK):
        raise ValueError('表A1是只读文件或当前账户没有写入权限，无法保存；请在文件属性中取消只读后重试')
    try:
        fd, probe = tempfile.mkstemp(prefix='.a1_probe_', suffix='.tmp', dir=path.parent)
        os.close(fd)
        os.unlink(probe)
    except OSError as exc:
        raise ValueError('表A1所在文件夹无法写入（可能只读、没有权限或磁盘已满），请更换保存位置后重试；本次没有修改表A1') from exc


def _replace_locked(tmp, path):
    """Replace the workbook, tolerating short-lived locks from sync/antivirus."""
    delay = 0.4
    for attempt in range(4):
        try:
            os.replace(tmp, path)
            return
        except PermissionError as exc:
            if attempt == 3:
                raise ValueError('表A1无法保存：文件正被其他程序占用（多为WPS/Excel已打开表A1，'
                                 '也可能是只读、云同步或杀毒软件锁定）。本次没有替换原文件，'
                                 '原表A1和备份都保持原样；请关闭表A1后重新复制') from exc
            time.sleep(delay)
            delay *= 2


def _read_clipboard_once(cb, win32con):
    cb.OpenClipboard()
    try:
        text = cb.GetClipboardData(win32con.CF_UNICODETEXT) if cb.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT) else ''
        fmt = cb.RegisterClipboardFormat('HTML Format')
        rich = cb.GetClipboardData(fmt) if cb.IsClipboardFormatAvailable(fmt) else b''
        if rich: rich=decode_context(rich)
        sequence = cb.GetClipboardSequenceNumber()
    finally:
        cb.CloseClipboard()
    return dict(text=text,html=rich,sequence=sequence)


def read_clipboard():
    try:
        import win32clipboard as cb
        import win32con
    except ImportError as exc:
        raise ValueError('缺少Excel复制组件，请运行安装依赖.bat安装pywin32并重启') from exc
    from retry_com import retry_call
    # 复制（Excel Ctrl+C）是异步的：刚按下时 GetClipboardData 可能拿不到句柄而报
    # “(0, 'GetClipboardData', ...)”，剪贴板也可能被其他程序短暂占用，这里重试。
    return retry_call(lambda: _read_clipboard_once(cb, win32con), attempts=8, delay=0.25)


class _TableShape:
    """Validate HTML spans with O(column count) state, without building a DOM."""
    def __init__(self, max_cells):
        self.max_cells = max_cells
        self.tables = 0
        self.in_table = self.in_row = False
        self.row = -1
        self.col = self.height = self.width = 0
        self.spans = []  # exclusive final row of a rowspan, per column
        # 记录含文字的最后一个单元格：Excel/WPS的纯文本会省略末尾空白行/列，
        # 补齐空白前先用它确认被省略的确实是空白，而不是复制被截断。
        self.last_text_row = self.last_text_col = -1
        # 每行最后一个含文字的列（0基，-1表示该行没有文字）：网页会把每行末尾的
        # 空单元格省掉，用它可以确认省掉的只是空白，而不是数据被挪了位置。
        self.row_ends = []
        self.cell_row = self.cell_col = self.cell_span = 0

    def start(self, tag, attrs):
        if tag in ('img', 'object', 'iframe'):
            raise ValueError('复制内容含图片或嵌入对象，当前不能验证其完整性，表A1未修改')
        if tag == 'table':
            self.tables += 1
            if self.tables != 1:
                raise ValueError('剪贴板不是单一表格')
            self.in_table = True
        elif tag == 'tr' and self.in_table:
            if self.in_row: raise ValueError('富文本表格行结构无效')
            self.row += 1
            self.col = 0
            self.in_row = True
            self.row_ends.append(-1)
        elif tag in ('td', 'th') and self.in_table:
            if not self.in_row: raise ValueError('富文本单元格不在表格行中')
            while self.col < len(self.spans) and self.spans[self.col] > self.row:
                self.col += 1
            rs, cs = int(attrs.get('rowspan', '1')), int(attrs.get('colspan', '1'))
            self.cell_row, self.cell_col, self.cell_span = self.row, self.col, cs
            bottom, right = self.row + rs, self.col + cs
            if rs < 1 or cs < 1 or bottom > 1048576 or right > 16384:
                raise ValueError('富文本表格范围无效或过大')
            check_cells(max(self.height,bottom), max(self.width,right), self.max_cells, '富文本范围')
            if right > len(self.spans): self.spans.extend([0] * (right - len(self.spans)))
            for c in range(self.col, right):
                if self.spans[c] > self.row: raise ValueError('富文本合并范围重叠')
                self.spans[c] = bottom
            self.height = max(self.height, bottom)
            self.width = max(self.width, right)
            self.col = right

    def end(self, tag):
        if tag == 'tr': self.in_row = False
        elif tag == 'table': self.in_table = False

    def data(self, text):
        if self.in_table and text and text.strip():
            self.last_text_row = max(self.last_text_row, self.cell_row)
            self.last_text_col = max(self.last_text_col, self.cell_col + self.cell_span - 1)
            if 0 <= self.cell_row < len(self.row_ends):
                self.row_ends[self.cell_row] = max(self.row_ends[self.cell_row],
                                                   self.cell_col + self.cell_span - 1)

    def close(self):
        if self.tables != 1: raise ValueError('剪贴板不是单一表格')
        return self.height, self.width


def _check_html_errors(parser):
    """Unknown HTML extension tags do not imply loss of table rows or cells.

    Office clipboard HTML can use o:p/x:* and other extension elements.
    libxml2 reports these as HTML_UNKNOWN_TAG at ERROR level, even though
    parsing the complete table succeeds. Keep all other errors actionable.
    Never include parser.message: it can contain copied business data.
    """
    for error in parser.feed_error_log:
        if error.level_name not in ('ERROR','FATAL'):
            continue
        if (error.level_name=='ERROR' and error.domain_name=='HTML' and
                error.type_name=='HTML_UNKNOWN_TAG'):
            continue
        # Only a size/resource failure is helped by smaller copy batches.
        resource_error=(error.type_name in ('ERR_RESOURCE_LIMIT','ERR_NO_MEMORY') or
                        (error.type_name=='ERR_ATTRIBUTE_NOT_FINISHED' and
                         error.message.lower().strip()=='attribute value too long') or
                        (error.type_name=='ERR_INTERNAL_ERROR' and
                         any(s in error.message.lower() for s in
                             ('resource limit','buffer size limit','huge input lookup','excessive depth'))))
        code=re.sub(r'[^A-Z0-9_]', '', error.type_name)[:80] or 'HTML_PARSE_ERROR'
        hint='将自动尝试缩小批次' if resource_error else '这是格式解析错误，继续缩小行数无效；请提供此错误代码'
        raise CopyIntegrityError(f'富文本解析失败（{code}，HTML第{error.line}行／第{error.column}列）；'
                                 f'{hint}，表A1未修改',can_split=resource_error)


def _visible_to_logical(index, omitted, first_col):
    """可见列序号（0基）→ 完整区域内的列序号（0基）。"""
    logical=index
    for column_number in omitted:
        if column_number-first_col<=logical:
            logical+=1
        else:
            break
    return logical


def _text_keeps_omitted(grid, shape, omitted, first_col):
    """网页的纯文本是不是仍带着那些被舍弃的隐藏列。

    两种约定都可能出现：纯文本跟富文本一样跳过隐藏列（要按列位补空白），
    或者纯文本仍按原列位给出隐藏列（列数已经够，不能再插空白）。
    用“每行最后一个有文字的列”区分：纯文本末列序号等于富文本末列序号→跳过；
    等于“富文本末列换算到完整区域后的列序号”→没跳过。两者都不符时不改，交给后续校验。"""
    text_ends=grid.get('row_ends') or []
    html_ends=getattr(shape,'row_ends',None) or []
    pairs=min(len(text_ends),len(html_ends))
    if not pairs:
        return False
    identity=expanded=True
    for index in range(pairs):
        text_end=text_ends[index]
        html_end=html_ends[index]
        if text_end!=html_end:
            identity=False
        if text_end!=_visible_to_logical(html_end,omitted,first_col):
            expanded=False
        if not identity and not expanded:
            return False
    # expanded 成立 = 纯文本末列位置是“富文本末列换算到完整区域”后的位置，
    # 说明纯文本仍按原列位带着隐藏列；identity 成立说明两边一样跳过了隐藏列。
    return expanded and not identity


def _trailing_columns_blank(grid, shape, width):
    """纯文本在 width 之后的列是不是全为空白，并且没有整体错位。

    网页复制时会把每行末尾的空单元格省掉，富文本因此比应有的列数窄（例如范围内隐藏了
    1 列、应有 79 列，富文本却只有 77 列），而纯文本仍保留这些列（空字段）。只有同时
    满足下面两条时才能按空白补齐：
    1. 纯文本在富文本宽度之后的列全部为空（这些列确实没有数据）；
    2. 每行最后一个有文字的列位置与纯文本完全一致（列没有整体错位）。
    两条都成立时，补上去的列在表A1里本来就是空白，不会覆盖任何内容。"""
    nonempty=grid.get('column_nonempty') or b''
    text_ends=grid.get('row_ends')
    html_ends=getattr(shape,'row_ends',None)
    if text_ends is None or html_ends is None:
        return False
    if any(nonempty[width:]):
        return False
    return all(text_ends[index]==html_ends[index]
               for index in range(min(len(text_ends),len(html_ends))))


_FILTER_SPACES=dict.fromkeys(map(ord,'\u00a0\u3000\u200b\ufeff'),'')


def _filter_text(value):
    """筛选比较用的文本：Excel/WPS给出的纯文本常常只差首尾空格、NBSP、全角空格或换行，
    这些都不应让“筛选 BW=在建”匹配不到行。"""
    text=str(value).translate(_FILTER_SPACES).replace('\r\n','\n').replace('\r','\n')
    return '\n'.join(part.strip() for part in text.split('\n')).strip()


def inspect_copy(clip, config, origin, mode, *, retain_rows=True, expected_range=None, batch=False,
                 strict_rows=False):
    first_col,first_row = address(origin)
    configured_omitted=check_omitted_target(config)
    if configured_omitted and (mode!='selection' or not expected_range):
        raise ValueError('设置舍弃隐藏列后，必须使用指定区域复制，才能保留原列位置')
    if not batch and (first_row != 1 or (mode == 'all' and origin != 'A1')):
        raise ValueError('复制区域必须包含第1行表头，整表模式从A1开始')
    if not clip.get('html'):
        raise ValueError('剪贴板没有提供带格式的表格（网页或Excel没有给出表格数据），已停止；不会降级为纯文本覆盖表A1')
    # 先解析纯文本（便宜），超过读取上限时不必再解析整段富文本。
    text=clip.get('text') or ''
    grid=None
    shape_error=None
    if text.strip():
        try:
            grid=parse(text,origin,config['source']['max_cells'],retain_rows=retain_rows,allow_blank=True)
        except PasteShapeError as exc:
            # 引号或控制字符让“带引号解析”的结果不能写进单元格：稍后用字面解析再试，
            # 富文本行列数是最终的判断依据，宁可保留原文也不整批失败。
            shape_error=exc
    # Validate the actual complete table, not surrounding clipboard headers or
    # duplicate document wrappers. The original clipboard is left untouched.
    html=table_fragment(clip['html'])
    # Parse in chunks; keep only the current row spans, not millions of cell tuples.
    shape=_TableShape(config['source']['max_cells'])
    parser=etree.HTMLParser(target=shape,no_network=True)
    try:
        for offset in range(0,len(html),65536):
            parser.feed(html[offset:offset+65536])
            _check_html_errors(parser)
        height,width=parser.close()
        _check_html_errors(parser)
    except (etree.LxmlError,TypeError) as exc:
        _check_html_errors(parser)
        raise CopyIntegrityError('复制内容不是可验证的富文本数据表（HTML_PARSE_ERROR），表A1未修改',can_split=False) from exc
    if height<1 or width<1:
        # 整片空白选区在网页上可能连表格行列都没有：只要纯文本也为空、且请求区域明确，
        # 就按请求区域生成空白数据继续粘贴，不再让“末尾空白行”中断复制。
        if expected_range and not text.strip():
            left,right=expected_range.split(':')
            c1,r1=address(left);c2,r2=address(right)
            logical_width=c2-c1+1
            check_cells(r2-r1+1,logical_width,config['source']['max_cells'],'完整复制范围')
            omitted=omitted_in_range(config,c1,logical_width)
            expected_width=logical_width-len(omitted)
            if expected_width<1:
                raise ValueError('指定区域的全部列都已设为舍弃，请保留至少一列数据')
            blank_rows=[['']*logical_width for _ in range(r2-r1+1)] if retain_rows else []
            return dict(first_column=c1, first_row=r1, width=logical_width,
                        row_count=r2-r1+1, header=['']*logical_width, nonempty_count=0,
                        rows=blank_rows, clipboard_width=expected_width,
                        omitted_columns=omitted, synthetic=True)
        # 其它情况说明复制被截断或选区本身无效；空白单元格仍会有空的td。
        raise CopyIntegrityError('复制结果没有任何单元格（可能选区为空或复制被截断），表A1未修改',can_split=True)
    # 指定区域时先拿剪贴板实际矩形与请求区域对比：复制被截断时直接说清“少了几行几列”，
    # 而不是让用户看到“富文本与文本行列数不一致”这种次要现象。
    # 唯一放行的情况是行数少、列数正好（末尾空白行），后面按空白补齐。
    expected_width=target_height=None
    if expected_range:
        left,right=expected_range.split(':')
        c1,r1=address(left);c2,r2=address(right)
        if (c1,r1)!=(first_col,first_row) or c2<c1 or r2<r1:
            raise ValueError('请求区域与复制起点不一致')
        check_cells(r2-r1+1,c2-c1+1,config['source']['max_cells'],'完整复制范围')
        expected_omitted=omitted_in_range(config,c1,c2-c1+1)
        expected_width=c2-c1+1-len(expected_omitted)
        target_height=r2-r1+1
        if width<expected_width and height==target_height and grid is not None \
                and grid['width']>width and _trailing_columns_blank(grid,shape,width):
            # 行数完整、只是富文本少了每行末尾的空单元格：按应有列数继续，
            # 隐藏列仍按“舍弃”处理，缺掉的末尾列在表A1里保持空白。
            width=expected_width
        if (height,width)!=(target_height,expected_width):
            blank_tail=(height<target_height and width==expected_width and not strict_rows)
            if not blank_tail:
                end=f'{column_name(first_col+width-1)}{first_row+height-1}'
                detail=(f'请求{expected_range}（{target_height:,}行 × {c2-c1+1:,}列')
                if expected_omitted:
                    detail+=f'，舍弃{column_list(expected_omitted)}后应为{expected_width:,}列'
                detail+=f'），剪贴板识别为{origin}:{end}（{height:,}行 × {width:,}列'
                disagree=grid is not None and (grid['row_count'],grid['width'])!=(height,width)
                if disagree:
                    detail+=f'，纯文本{grid["row_count"]:,}行 × {grid["width"]:,}列'
                detail+='）'
                if disagree:
                    # 富文本与纯文本本身对不上：单独点出来，便于区分“网页少给列”与“复制被截断”。
                    detail+='；富文本与文本行列数不一致'
                if height<target_height:
                    hint=('剪贴板行数少于请求：多半是网页还没把这些行渲染出来、有行被筛选或隐藏，'
                          '也可能是复制被截断；将自动缩小每批行数重试')
                elif width<expected_width:
                    hint=('剪贴板列数少于请求，不能用空白列补齐未明确舍弃的缺失数据：'
                          '请核对网页的隐藏列／筛选状态，或把隐藏列填入“舍弃的隐藏列”')
                else:
                    hint='剪贴板范围大于请求区域，请重新确认网页选区后重试'
                raise CopyIntegrityError(f'复制结果与指定区域大小不一致：{detail}。{hint}，表A1未修改',
                                         can_split=height<target_height,
                                         kind='rows' if height<target_height else 'columns')
    # 空白单元格、空白行、整批空白都不算错误：只要富文本表格的行列数可以确认，
    # 就按“空白内容”正常粘贴，不再要求表头或数据非空。
    if text.strip() and (grid is None or (grid['row_count'],grid['width'])!=(height,width)):
        # 带引号解析失败或行数对不上时，按“整行=制表符切分”的字面方式再解析一次；
        # 只有它的行列数与富文本完全一致才采用，避免用错误的行切分覆盖表A1。
        try:
            literal=parse(text,origin,config['source']['max_cells'],retain_rows=retain_rows,
                          allow_blank=True,literal=True)
        except ValueError:
            literal=None
        if literal is not None and (literal['row_count'],literal['width'])==(height,width):
            grid=literal
        elif grid is None:
            raise shape_error or ValueError('粘贴内容无法解析，表A1未修改')
    if grid is not None:
        if (grid['row_count'],grid['width'])!=(height,width):
            # Excel/WPS的纯文本剪贴板会省略末尾的空白行和空白列（中间空白仍保留）。
            # 列数一致地少、或末尾列在富文本里没有文字时，按空白补齐；
            # 行数少且缺失行在富文本里是空白（末尾空白行）时，按空白补齐并标记未核对；
            # 缺失行在富文本里仍有文字说明是中段丢行，绝不能补齐（会把后面行整体上移）。
            if grid['width']>width and _trailing_columns_blank(grid,shape,width):
                # 纯文本多出来的只是空白列（网页省掉了富文本末尾的空单元格）：裁掉不算丢数据。
                if retain_rows:
                    grid['rows']=[row[:width] for row in grid['rows']]
                grid['header']=list(grid['header'])[:width]
                grid['width']=width
            if grid['width']>width or grid['row_count']>height:
                raise CopyIntegrityError(f'富文本与文本行列数不一致：富文本{height:,}行 × {width:,}列，'
                                         f'文本{grid["row_count"]:,}行 × {grid["width"]:,}列；'
                                         '无法确认完整复制，表A1未修改',
                                         can_split=width==0 or width==grid['width'])
            if grid['width']<width:
                if width>0 and shape.last_text_col>=grid['width']:
                    raise CopyIntegrityError(f'富文本与文本行列数不一致：富文本{height:,}行 × {width:,}列，'
                                             f'文本{grid["row_count"]:,}行 × {grid["width"]:,}列；'
                                             '被省略的列在富文本里仍有文字，无法确认完整复制，表A1未修改',
                                             can_split=False)
                if retain_rows:
                    grid['rows']=[row+['']*(width-grid['width']) for row in grid['rows']]
                if len(grid['header'])<width:
                    grid['header']=list(grid['header'])+['']*(width-len(grid['header']))
                grid['width']=width
            if grid['row_count']<height:
                missing=height-grid['row_count']
                if shape.last_text_row>=grid['row_count']:
                    # 富文本在纯文本缺失的行里仍有文字：说明纯文本不是把末尾空白行省掉，
                    # 而是中段丢了行（例如个别单元格的引号把后续若干行并进一个单元格，
                    # 或网页复制/渲染被截断）。此时按末尾空白补齐会把后面的行整体上移，
                    # 主键与行号全部错位，绝不能静默补齐；交给调用方缩小批次重试。
                    raise CopyIntegrityError(
                        f'纯文本剪贴板比富文本少{missing:,}行，且缺失的行在富文本里仍有内容：'
                        '复制结果不完整，无法安全对齐行，表A1未修改', can_split=True, kind='rows')
                grid['unverified_rows']=missing
                if retain_rows:
                    grid['rows']=grid['rows']+[['']*width for _ in range(missing)]
                grid['row_count']=height
            grid['clipboard_width']=grid['width']
    else:
        # Excel/WPS对全空白选区的纯文本可能是空的，此时按富文本确认的行列生成空白数据。
        grid=dict(first_column=first_col, first_row=first_row, width=width, row_count=height,
                  header=['']*width, nonempty_count=0,
                  rows=[['']*width for _ in range(height)] if retain_rows else [])
    logical_width=width
    omitted=()
    if expected_range:
        left,right=expected_range.split(':')
        c1,r1=address(left);c2,r2=address(right)
        if (c1,r1)!=(first_col,first_row) or c2<c1 or r2<r1:
            raise ValueError('请求区域与复制起点不一致')
        logical_width=c2-c1+1
        check_cells(r2-r1+1,logical_width,config['source']['max_cells'],'完整复制范围')
        omitted=omitted_in_range(config,c1,logical_width)
        expected_width=logical_width-len(omitted)
        if expected_width<1:
            raise ValueError('指定区域的全部列都已设为舍弃，请保留至少一列数据')
        target_height=r2-r1+1
        if height<target_height and width==expected_width:
            # 网页复制时会把末尾的空白行省略（整片空白时甚至没有表格）。
            # 行宽完全一致、只是行数变少，说明缺的是末尾空白行：按空白补齐后继续粘贴。
            extra=target_height-height
            if retain_rows:
                grid['rows'].extend([['']*width for _ in range(extra)])
            if len(grid['header'])<width:
                grid['header']=list(grid['header'])+['']*(width-len(grid['header']))
            grid['row_count']=target_height
            height=target_height
        if (height,width)!=(target_height,expected_width):
            if omitted:
                raise CopyIntegrityError(f'复制结果与指定区域大小不一致：请求{expected_range}（{r2-r1+1:,}行 × {logical_width:,}列），'
                    f'已设置舍弃{column_list(omitted)}，因此剪贴板应为{r2-r1+1:,}行 × {expected_width:,}列；'
                    f'实际为{height:,}行 × {width:,}列。请核对网页隐藏列与设置是否一致，表A1未修改',can_split=width==expected_width,
                    kind='columns')
            end=f'{column_name(first_col+width-1)}{first_row+height-1}'
            raise CopyIntegrityError(f'复制结果与指定区域大小不一致：请求{expected_range}（{r2-r1+1:,}行 × {c2-c1+1:,}列），'
                             f'剪贴板识别为{origin}:{end}（{height:,}行 × {width:,}列）。请确认网页工作表、选区以及隐藏列／筛选状态；不能用空白列补齐未明确舍弃的缺失数据，表A1未修改',can_split=width==c2-c1+1,
                             kind='columns')
    grid['clipboard_width']=width
    grid['omitted_columns']=omitted
    grid['width']=logical_width
    if omitted:
        if _text_keeps_omitted(grid,shape,omitted,first_col):
            # 纯文本已经按原列位带着隐藏列：列位与完整区域一一对应，
            # 再按“舍弃列”插空白会整行右移（筛选就会读到隔壁列而匹配不到行）。
            grid['text_includes_omitted']=True
            if retain_rows:
                rows=[]
                for row in grid['rows']:
                    row=list(row)+['']*(logical_width-len(row))
                    for column_number in omitted:
                        row[column_number-first_col]=''
                    rows.append(row)
                grid['rows']=rows
            header=list(grid['header'])+['']*(logical_width-len(grid['header']))
            for column_number in omitted:
                header[column_number-first_col]=''
            grid['header']=header[:logical_width]
            # 隐藏列不会出现在粘贴内容里，非空计数要把它们扣掉，否则回读核对会差这些格。
            counts=grid.get('column_counts') or []
            dropped=sum(counts[column_number-first_col] for column_number in omitted
                        if 0<=column_number-first_col<len(counts))
            grid['nonempty_count']=max(0,grid['nonempty_count']-dropped)
        else:
            grid['header']=expand_values(grid['header'],first_col,logical_width,omitted)
            if retain_rows:
                grid['rows']=[expand_values(row,first_col,logical_width,omitted) for row in grid['rows']]
    return grid


def paste_to_a1(root, config, clip, origin, mode, expected_range=None, layout=None,
                paste_origin=None, filter_spec=None, progress=None, progress_fraction=None):
    grid=inspect_copy(clip,config,origin,mode,retain_rows=filter_spec is not None,expected_range=expected_range)
    return paste_batches_to_a1(root,config,iter([(clip,grid)]),origin,grid['row_count'],grid['width'],layout=layout,
                               paste_origin=paste_origin,filter_spec=filter_spec,progress=progress,
                               progress_fraction=progress_fraction)


def apply_reference_format(excel, sheet, reference, first_col, first_row, total_rows, width):
    """Apply layout from the reference via Excel automation: merged ranges.

    Row heights, column widths, default heights and the frozen pane are all
    written into the saved file by apply_column_layout() instead: the COM
    layer (WPS/Excel) snaps widths to pixels, its freeze-pane assignment is
    unreliable headless, and paste-derived row heights must be normalized
    away for every row 表C does not define."""
    stats = dict(widths=0, heights=0, merges=0, panes=False, sheet=None, partial=False)
    if not reference:
        return stats
    stats['sheet'] = reference['sheet']
    stats['partial'] = bool(reference.get('partial'))
    end_col, end_row = first_col + width - 1, first_row + total_rows - 1
    for c1, r1, c2, r2 in reference['merges']:
        left, top = max(c1, first_col), max(r1, first_row)
        right, bottom = min(c2, end_col), min(r2, end_row)
        if right < left or bottom < top or (left == right and top == bottom):
            continue
        ref = f'{column_name(left)}{top}:{column_name(right)}{bottom}'
        sheet.Range(ref).UnMerge()
        sheet.Range(ref).Merge()
        stats['merges'] += 1
    return stats


def apply_column_layout(path, sheet_name, reference, first_col, end_col, first_row, end_row):
    """Write the exact layout from the reference directly into the saved xlsx
    (after Excel/WPS has saved it): column widths, default row/column
    heights, row heights and the frozen pane.

    Every row of the copied range is normalized against 表C: rows with an
    explicit height get it written; all other rows lose any height the paste
    produced and fall back to the default row height. This guarantees paste
    styles can never leave the header too tall."""
    stats = dict(widths=0, heights=0, default_row=False, default_col=False, panes=False)
    if not reference:
        return stats
    data = path.read_bytes()
    archive = zipfile.ZipFile(io.BytesIO(data))
    try:
        names = archive.namelist()
        wb = xml(archive.read('xl/workbook.xml'))
        sheets = wb.find(tag('sheets'))
        target = next((s for s in sheets if s.get('name') == sheet_name), None)
        if target is None:
            raise ValueError('表A1中找不到指定工作表，无法写入列宽')
        rels = xml(archive.read('xl/_rels/workbook.xml.rels'))
        relation = next((r for r in rels if r.get('Id') == target.get('{' + REL + '}id')), None)
        if relation is None or relation.get('TargetMode') == 'External':
            raise ValueError('表A1工作表关系无效，无法写入列宽')
        link = relation.get('Target')
        name = posixpath.normpath(link.lstrip('/') if link.startswith('/') else posixpath.join('xl', link))
        if name not in names:
            raise ValueError('表A1工作表路径无效，无法写入列宽')
        tree = xml(archive.read(name))
        fmt = tree.find(tag('sheetFormatPr'))
        if fmt is None:
            fmt = etree.Element(tag('sheetFormatPr'))
            tree.insert(0, fmt)
        normalize_rows = reference.get('default_row_height') is not None
        if reference.get('default_row_height'):
            fmt.set('defaultRowHeight', repr(reference['default_row_height']))
            fmt.set('customHeight', '1')
            stats['default_row'] = True
        if reference.get('default_col_width'):
            fmt.set('defaultColWidth', repr(reference['default_col_width']))
            stats['default_col'] = True
        old = tree.find(tag('cols'))
        if old is not None:
            tree.remove(old)
        cols = etree.Element(tag('cols'))
        data_node = tree.find(tag('sheetData'))
        tree.insert(list(tree).index(data_node) if data_node is not None else 0, cols)
        for col in range(first_col, end_col + 1):
            width = reference['column_widths'].get(col)
            if width is None:
                continue
            entry = etree.SubElement(cols, tag('col'))
            entry.set('min', str(col))
            entry.set('max', str(col))
            entry.set('width', repr(width))
            entry.set('customWidth', '1')
            stats['widths'] += 1
        # 行高归一化：表C有显式行高则写入；否则清除粘贴行高，回落默认行高。
        row_heights = reference['row_heights']
        if data_node is not None and (row_heights or normalize_rows):
            existing = {}
            for row_el in data_node.findall(tag('row')):
                try:
                    number = int(row_el.get('r'))
                except (TypeError, ValueError):
                    continue
                existing[number] = row_el
                if not first_row <= number <= end_row:
                    continue
                height = row_heights.get(number)
                if height is not None:
                    row_el.set('ht', repr(height))
                    row_el.set('customHeight', '1')
                    stats['heights'] += 1
                elif normalize_rows:
                    row_el.attrib.pop('ht', None)
                    row_el.attrib.pop('customHeight', None)
            # 表C有显式行高但表A1中没有该行元素（空行）时，补一个只含行高的行元素。
            for number in sorted(row_heights):
                if not first_row <= number <= end_row or number in existing:
                    continue
                row_el = etree.Element(tag('row'))
                row_el.set('r', str(number))
                row_el.set('ht', repr(row_heights[number]))
                row_el.set('customHeight', '1')
                before = next((r for r in data_node if int(r.get('r')) > number), None)
                if before is None:
                    data_node.append(row_el)
                else:
                    data_node.insert(data_node.index(before), row_el)
                stats['heights'] += 1
        pane = reference['pane']
        if pane:
            x, y, top_left = pane
            if 0 < x <= end_col and 0 < y <= end_row:
                views = tree.find(tag('sheetViews'))
                if views is None:
                    views = etree.Element(tag('sheetViews'))
                    tree.insert(0, views)
                view = views.find(tag('sheetView'))
                if view is None:
                    view = etree.SubElement(views, tag('sheetView'))
                    view.set('workbookViewId', '0')
                old_pane = view.find(tag('pane'))
                if old_pane is not None:
                    view.remove(old_pane)
                pane_el = etree.Element(tag('pane'))
                pane_el.set('xSplit', str(x))
                pane_el.set('ySplit', str(y))
                pane_el.set('topLeftCell', top_left)
                pane_el.set('activePane', 'bottomRight')
                pane_el.set('state', 'frozen')
                view.insert(0, pane_el)
                for selection in view.findall(tag('selection')):
                    view.remove(selection)
                selection = etree.Element(tag('selection'))
                selection.set('pane', 'bottomRight')
                selection.set('activeCell', top_left)
                selection.set('sqref', top_left)
                view.append(selection)
                stats['panes'] = True
        replacements = {name: etree.tostring(tree, encoding='UTF-8', xml_declaration=True)}
        fd, tmp = tempfile.mkstemp(prefix='.a1_columns_', suffix='.xlsx', dir=path.parent)
        os.close(fd)
        try:
            with zipfile.ZipFile(tmp, 'w') as output:
                output.comment = archive.comment
                for info in archive.infolist():
                    output.writestr(info, replacements.get(info.filename, archive.read(info.filename)))
            os.replace(tmp, path)
        finally:
            if os.path.exists(tmp):
                os.unlink(tmp)
    finally:
        archive.close()
    return stats


def combine_layout(file_reference, layout):
    """表C（本地格式表）是权威格式来源，在线页面读取只补齐表C没有的部分。

    file_reference/layout 均为 dict(column_widths, row_heights, merges, ...)，
    键为绝对坐标。表C 只提供格式、不提供数据。"""
    if layout is None:
        return file_reference
    if file_reference is None:
        combined = dict(layout)
        combined['sheet'] = '在线页面'
        combined['partial'] = bool(layout.get('partial'))
        return combined
    combined = dict(file_reference)
    for key in ('column_widths', 'row_heights'):
        merged = dict(layout[key])
        merged.update(file_reference[key])
        combined[key] = merged
    combined['merges'] = file_reference['merges'] if file_reference['merges'] else layout['merges']
    combined['sheet'] = '表C' + ('+在线页面' if layout else '')
    combined['partial'] = False
    return combined


def _offset_reference(reference, dcol, drow):
    """平移表C格式坐标：源区域粘贴到偏移位置时，行高/列宽/合并随目标移动。"""
    if not reference or (dcol == 0 and drow == 0):
        return reference
    result = dict(reference)
    result['merges'] = [(m[0]+dcol, m[1]+drow, m[2]+dcol, m[3]+drow)
                        for m in reference.get('merges', ())
                        if m[2]+dcol <= 16384 and m[3]+drow <= 1048576]
    result['row_heights'] = {k+drow: v for k, v in reference.get('row_heights', {}).items()
                             if k+drow <= 1048576}
    result['column_widths'] = {k+dcol: v for k, v in reference.get('column_widths', {}).items()
                               if k+dcol <= 16384}
    result['column_width_points'] = {k+dcol: v for k, v in (reference.get('column_width_points') or {}).items()
                                     if k+dcol <= 16384}
    result['pane'] = None  # 偏移粘贴不沿用原冻结窗格位置
    return result


def paste_batches_to_a1(root, config, batches, origin, total_rows, width, layout=None,
                        paste_origin=None, filter_spec=None, progress=None, progress_fraction=None):
    """Consume one verified clipboard at a time; commit only after every batch.

    layout: live page layout from page_layout.extract(); 表C (format
    reference file) wins wherever configured, the page fills the gaps.

    paste_origin: 表A1自定义粘贴起点（默认与源区域同位置）。
    filter_spec: {'column': n, 'value': s, 'name': 'BW'} 时只粘贴该列等于value的行，
    连续堆叠到粘贴起点；此时不应用表C结构格式（行高/合并/冻结），列宽与样式由粘贴携带。
    progress: 每批粘贴/筛选完成后回调，向网页报告进度，避免长时间无响应。
    progress_fraction: 每批完成后回调已处理行比例（0~1），用于进度条。
    """
    first_col,first_row=address(origin)
    check_omitted_target(config)
    omitted=omitted_in_range(config,first_col,width)
    clipboard_width=width-len(omitted)
    if clipboard_width<1: raise ValueError('指定区域需保留至少一列数据')
    if filter_spec and filter_spec['column'] > width:
        raise ValueError(f'筛选列 {filter_spec["name"]} 不在复制区域内')
    if filter_spec and filter_spec['column'] in omitted:
        # 隐藏列在表A1按空白保留、剪贴板里也没有这些列，拿它筛选永远匹配不到行。
        raise ValueError(f'筛选列 {filter_spec["name"]} 正好是被隐藏/舍弃的列：'
                         f'跳过它之后这一列没有内容，按它筛选永远匹配不到行。'
                         f'请先在网页取消隐藏该列（或从“舍弃的隐藏列”里去掉它），'
                         f'或改用别的筛选列后重试；表A1未修改')
    paste_col,paste_row=address(paste_origin or origin)
    if paste_col+width-1>16384 or paste_row+total_rows-1>1048576:
        raise ValueError('粘贴起点导致区域超出Excel行列限制')
    check_cells(total_rows,width,config['source']['max_cells'],'完整复制范围')
    path=source_path(root,config)
    _check_writable_target(path)
    file_reference=load_format_reference(root,config)
    reference=combine_layout(file_reference,layout)
    if filter_spec:
        reference=None  # 筛选结果行已重排，表C绝对坐标不再对应
    elif paste_origin:
        reference=_offset_reference(reference, paste_col-first_col, paste_row-first_row)
    original=path.read_bytes();digest=hashlib.sha256(original).digest()
    try:
        import pythoncom
        import win32com.client
        import win32clipboard
    except ImportError as exc:
        raise ValueError('缺少Excel自动化组件，请重新运行安装依赖.bat') from exc
    fd,tmp=tempfile.mkstemp(prefix='.a1_formatted_',suffix='.xlsx',dir=path.parent)
    os.close(fd);Path(tmp).write_bytes(original)
    excel=book=scratch=None
    pythoncom.CoInitialize()
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
        excel.Visible=False
        excel.DisplayAlerts=False
        excel.EnableEvents=False
        excel.AutomationSecurity=3
        excel.AskToUpdateLinks=False
        book=excel.Workbooks.Open(str(Path(tmp).resolve()),UpdateLinks=0,ReadOnly=False)
        try: sheet=book.Worksheets(config['source']['sheet_name'])
        except Exception as exc: raise ValueError('表A1目标工作表不存在') from exc
        if sheet.ProtectContents:raise ValueError('表A1工作表受保护，请先解除保护或换用中转副本')
        check_cells(sheet.UsedRange.Rows.Count,sheet.UsedRange.Columns.Count,config['source']['max_cells'],'表A1旧使用范围')
        end_row=paste_row+total_rows-1;end_col=paste_col+width-1
        used_before=sheet.UsedRange
        last_used_col=used_before.Column+used_before.Columns.Count-1
        last_used_row=used_before.Row+used_before.Rows.Count-1
        if filter_spec:
            # 筛选模式：行数按匹配结果动态决定，先清空粘贴列下方旧数据避免残留。
            tail_bottom=max(paste_row,last_used_row)
            tail=sheet.Range(sheet.Cells(paste_row,paste_col),sheet.Cells(tail_bottom,paste_col+width-1))
            tail.UnMerge()
            tail.Clear()
        elif paste_origin is None and first_col==1 and first_row==1 and end_col>=last_used_col:
            # 从A1开始且覆盖旧表整宽：完整替换并清空旧尾部（原有行为）。
            used_before.UnMerge()
            used_before.Clear()
        else:
            # 随意区域或自定义起点：只替换目标区域，其余内容保留。
            region=sheet.Range(sheet.Cells(paste_row,paste_col),sheet.Cells(end_row,end_col))
            region.UnMerge()
            region.Clear()
        destination=sheet
        if omitted:
            # Insert columns only in a disposable workbook. Inserting in A1
            # could rewrite formulas/names on its other worksheets.
            scratch=excel.Workbooks.Add(-4167)  # xlWBATWorksheet
            sheet=scratch.Worksheets(1)
        sheet.Activate()
        source_next=first_row
        paste_next=paste_row
        batch_count=0
        total_nonempty=0
        total_kept=0
        filter_seen={}   # 筛选列实际出现过的值（只统计该列，用于“没有匹配行”时说明原因）
        filter_rows=0
        value_columns={}  # 第一批里“筛选值真正出现在哪一列”，用于定位整行错位
        scanned=False
        for clip,grid in batches:
            rows=grid['row_count']
            if (grid['first_column']!=first_col or grid['first_row']!=source_next or
                    grid['width']!=width or grid.get('clipboard_width',grid['width'])!=clipboard_width or
                    tuple(grid.get('omitted_columns',()))!=omitted or
                    rows<1 or source_next+rows>first_row+total_rows):
                raise ValueError('分批数据不连续或超出完整区域，表A1未修改')
            if filter_spec:
                idx=filter_spec['column']-1
                want=_filter_text(filter_spec['value'])
                matched=[]
                scan=not scanned
                for i,row in enumerate(grid['rows']):
                    filter_rows+=1
                    if idx<len(row):
                        value=_filter_text(row[idx])
                        if value:
                            filter_seen[value]=filter_seen.get(value,0)+1
                            if value==want: matched.append(i)
                    if scan:
                        # 只在第一批扫一遍“这一列到底在哪”：整行错位时筛选会一行都匹配不到，
                        # 有了这一列的位置就能直接告诉用户该填哪一列/该取消哪个隐藏列。
                        for j in range(min(width,len(row))):
                            if j!=idx and _filter_text(row[j])==want:
                                value_columns[j]=value_columns.get(j,0)+1
                scanned=True
                expected_nonempty=sum(cell!='' for i in matched for cell in grid['rows'][i])
            else:
                matched=None
            source_next+=rows
            if filter_spec and not matched:
                continue  # 本批没有匹配行，跳过粘贴
            start=f'{column_name(paste_col)}{paste_next}'
            if clip.get('blank'):
                # 整批都是空白单元格：目标区域在开始时已经清空，不需要再把剪贴板
                # 里的占位文本粘进去；直接计入进度继续下一批。
                batch_count+=1
                paste_next+=rows
                if progress_fraction is not None:
                    progress_fraction((source_next-first_row)/total_rows)
                if progress:
                    progress(f'已按空白写入第{batch_count}批：累计 {paste_next-paste_row:,}/{total_rows:,} 行。')
                continue
            sheet.Range(start).Select()
            if win32clipboard.GetClipboardSequenceNumber()!=clip['sequence']:
                raise ValueError('复制后剪贴板被其他程序改动，请重试；表A1未修改')
            sheet.Paste(Destination=sheet.Range(start))
            if win32clipboard.GetClipboardSequenceNumber()!=clip['sequence']:
                raise ValueError('粘贴期间剪贴板发生变化，已取消保存')
            area=sheet.Range(sheet.Cells(paste_next,paste_col),sheet.Cells(paste_next+rows-1,paste_col+clipboard_width-1))
            try: formulas=area.SpecialCells(-4123)
            except Exception: formulas=None
            if formulas is not None:
                raise ValueError('复制内容包含公式，为避免引入外部引用，本次未保存；请将源显示结果作为值复制')
            if filter_spec:
                area.UnMerge()
                # 连续的非匹配行合并成段；再每30段用Union合并为一次删除。
                # 逐行删除会有上千次WPS/Excel COM调用，表现为“复制到一半卡死”。
                matched_set=set(matched)
                runs=[];run_start=None
                for i in range(rows):
                    if i in matched_set:
                        if run_start is not None:
                            runs.append((run_start,i-1));run_start=None
                    elif run_start is None:
                        run_start=i
                if run_start is not None:
                    runs.append((run_start,rows-1))
                ranges=[sheet.Range(sheet.Cells(paste_next+a,paste_col),
                                    sheet.Cells(paste_next+b,paste_col+clipboard_width-1))
                        for a,b in runs]
                # 必须从下往上分批删：一次Delete会把被删行下面的所有行整体上移，
                # 若先删上面的段，后面各段的绝对行号就全错位了——非匹配行删不干净，
                # 还会误删匹配行（表现为“非空格数比期望多”）。
                for offset in range((len(ranges)-1)//30*30,-1,-30) if ranges else ():
                    chunk=ranges[offset:offset+30]
                    if len(chunk)==1:
                        chunk[0].Delete(-4162)
                    else:
                        excel.Union(*chunk).Delete(-4162)
                kept=sheet.Range(sheet.Cells(paste_next,paste_col),
                                 sheet.Cells(paste_next+len(matched)-1,paste_col+clipboard_width-1))
                # WPS 删除几千行是异步的：刚删完立刻数，可能只数到删了一半的结果。
                # 这里最多再等约6秒复核，只有一直对不上才算真不一致。
                actual=int(excel.WorksheetFunction.CountA(kept))
                for _attempt in range(20):
                    if actual>=expected_nonempty: break
                    time.sleep(0.3)
                    actual=int(excel.WorksheetFunction.CountA(kept))
                if actual<expected_nonempty:
                    # 粘贴出来的内容比纯文本描述的还少：可能丢内容，必须停下。
                    raise ValueError(_filter_count_mismatch(
                        excel,sheet,start,paste_next,paste_col,len(matched),clipboard_width,
                        expected_nonempty,actual,grid,matched,first_col,omitted)+'，表A1未修改')
                if actual>expected_nonempty and progress:
                    # 纯文本剪贴板只是交叉核对，真正粘进表A1的是富文本：纯文本会省略
                    # 部分空白单元格/单元格内的空白，多出来的格说明纯文本描述不全，
                    # 不会因此丢数据，按实际粘贴内容继续。
                    progress(f'第{batch_count+1}批：纯文本比实际粘贴少描述'
                             f'{actual-expected_nonempty:,}格（纯文本会省略部分空白内容），'
                             f'已按实际粘贴内容继续。')
                total_nonempty+=actual
                total_kept+=len(matched)
                paste_next+=len(matched)
            else:
                if int(excel.WorksheetFunction.CountA(area))!=grid['nonempty_count']:
                    raise ValueError(f'Excel粘贴后的非空单元格数量与本批复制内容不一致（{start}起），表A1未修改')
                total_nonempty+=grid['nonempty_count']
                paste_next+=rows
            batch_count+=1
            if progress_fraction is not None:
                progress_fraction((source_next-first_row)/total_rows)
            if progress:
                if filter_spec:
                    progress(f'已粘贴并筛选第{batch_count}批：累计 {total_kept:,} 行（{filter_spec["name"]}={filter_spec["value"]}）。')
                else:
                    progress(f'已粘贴第{batch_count}批：累计 {paste_next-paste_row:,}/{total_rows:,} 行。')
        if source_next!=first_row+total_rows:
            raise ValueError('分批复制尚未完成，表A1未修改')
        if filter_spec:
            if not total_kept:
                sample='、'.join(f'“{value}”' for value in
                                 sorted(filter_seen,key=lambda v:-filter_seen[v])[:4])
                seen=f'这一列实际出现：{sample}' if sample else '这一列在这批区域里全是空白'
                message=(f'筛选 {filter_spec["name"]}={filter_spec["value"]} 没有匹配行'
                         f'（共看了{filter_rows:,}行，{seen}）。请核对筛选值的写法、'
                         f'筛选列字母（第{filter_spec["column"]}列）是否填对，'
                         f'以及该列是否被隐藏（隐藏列会按空白跳过）')
                if value_columns:
                    hits=sorted(value_columns.items(),key=lambda item:-item[1])[:3]
                    where='、'.join(f'{column_name(first_col+column_index)}'
                                    f'（第{first_col+column_index}列，{count:,}行）'
                                    for column_index,count in hits)
                    offset=hits[0][0]-(filter_spec['column']-first_col)
                    message+=(f'。不过“{filter_spec["value"]}”其实出现在 {where}，'
                              f'看起来整行错开了{offset:+d}列：请在网页上取消隐藏这些列'
                              f'（或把隐藏列按正确列字母填进“舍弃的隐藏列”）后重试')
                raise ValueError(message+'；表A1未修改')
            end_row=paste_row+total_kept-1
        elif not batch_count:
            # 空白单元格／空白行不算错误：只要确实粘贴过批次就继续保存。
            raise ValueError('没有可粘贴的批次，表A1未修改')
        if omitted:
            # 舍弃列（手工填写或自动识别的隐藏列）在临时工作表里按原列位置插入空白，
            # 自定义粘贴起点时按同一个偏移量平移，保证列位与表A一致。
            shift = paste_col - first_col
            for col in omitted:
                sheet.Columns(col + shift).Insert()
            full=sheet.Range(sheet.Cells(paste_row,paste_col),sheet.Cells(end_row,end_col))
            for col in omitted:
                empty=sheet.Range(sheet.Cells(paste_row,col + shift),sheet.Cells(end_row,col + shift))
                if int(excel.WorksheetFunction.CountA(empty)):
                    raise ValueError('舍弃列恢复位置后不是空白，已取消保存表A1')
            if int(excel.WorksheetFunction.CountA(full))!=total_nonempty:
                raise ValueError('恢复原列位置后单元格数量不一致，表A1未修改')
            full.Copy(Destination=destination.Range(paste_origin or origin))
            _copy_dimensions(sheet,destination,paste_col,paste_row,total_rows,width)
            final=destination.Range(destination.Cells(paste_row,paste_col),destination.Cells(end_row,end_col))
            if int(excel.WorksheetFunction.CountA(final))!=total_nonempty:
                raise ValueError('写回表A1后的单元格数量不一致，表A1未修改')
            scratch.Close(SaveChanges=False);scratch=None
            sheet=destination
        formats=apply_reference_format(excel,sheet,reference,paste_col,paste_row,total_rows,width)
        used=sheet.UsedRange
        if not filter_spec and (used.Row+used.Rows.Count-1>end_row or used.Column+used.Columns.Count-1>end_col):
            # Old formatting can keep UsedRange large, so validate actual content instead.
            found=sheet.Cells.Find(What='*',LookIn=-4123,SearchOrder=1,SearchDirection=2)
            if found is not None and found.Row>end_row: raise ValueError('粘贴结果行数超出预期，已取消保存')
            found=sheet.Cells.Find(What='*',LookIn=-4123,SearchOrder=2,SearchDirection=2)
            if found is not None and found.Column>end_col: raise ValueError('粘贴结果列数超出预期，已取消保存')
        book.Save();book.Close(SaveChanges=False);book=None
        if reference is not None:
            # 精确列宽、默认行高/列宽、行高归一化、冻结窗格直接写入已保存的文件，
            # 绕开COM的像素取整、无头冻结失效与粘贴行高污染。
            column_stats=apply_column_layout(Path(tmp),config['source']['sheet_name'],
                                             reference,paste_col,end_col,paste_row,end_row)
            formats['widths']=column_stats['widths']
            formats['heights']=column_stats['heights']
            formats['panes']=column_stats['panes']
        if hashlib.sha256(path.read_bytes()).digest()!=digest:
            raise ValueError('表A1同时被其他程序修改，已取消替换')
        folder=path.parent/(path.stem+'_备份')
        try:
            folder.mkdir(exist_ok=True)
        except OSError as exc:
            raise ValueError('表A1备份文件夹无法创建（可能只读、没有权限或磁盘已满），已取消保存；表A1未修改') from exc
        backup=folder/(path.stem+'_'+dt.datetime.now().strftime('%Y%m%d_%H%M%S_%f')+'.xlsx')
        try:
            with backup.open('xb') as f:f.write(original)
        except OSError as exc:
            raise ValueError('表A1备份无法写入（可能磁盘空间不足或没有权限），已取消保存；表A1未修改') from exc
        if hashlib.sha256(path.read_bytes()).digest()!=digest:raise ValueError('表A1发生变化，已取消替换')
        _replace_locked(tmp,path)
        result=dict(rows=total_kept if filter_spec else total_rows,columns=width,
                    start_cell=paste_origin or origin,batches=batch_count,backup=str(backup),
                    omitted_columns=list(omitted),data_columns=clipboard_width,
                    format_mode='Excel分批粘贴全部（筛选）' if filter_spec else 'Excel分批粘贴全部',
                    formats=formats if reference else None)
        if filter_spec:
            result.update(filtered=True,filter_column=filter_spec['name'],
                          filter_value=filter_spec['value'],matched_rows=total_kept)
        return result
    finally:
        if scratch is not None:
            try:scratch.Close(SaveChanges=False)
            except Exception:pass
        if book is not None:
            try:book.Close(SaveChanges=False)
            except Exception:pass
        if excel is not None:
            try:excel.Quit()
            except Exception:pass
        pythoncom.CoUninitialize()
        if os.path.exists(tmp):os.unlink(tmp)


def _filter_count_mismatch(excel, sheet, start, paste_next, paste_col, matched_rows, clipboard_width,
                           expected_nonempty, actual, grid, matched, first_col, omitted):
    """筛选粘贴后非空格数对不上时，给出差多少、差在哪几列，便于一次定位。"""
    message=(f'筛选粘贴后非空单元格数量不一致（{start}起）：本批{matched_rows:,}行，'
             f'期望{expected_nonempty:,}格有内容，实际{actual:,}格'
             f'（差{actual-expected_nonempty:+,}）')
    try:
        expected_columns=[0]*clipboard_width
        for index in matched:
            position=0
            for offset,cell in enumerate(grid['rows'][index]):
                if (first_col+offset) in omitted:
                    continue
                if position<clipboard_width and cell!='':
                    expected_columns[position]+=1
                position+=1
        parts=[]
        for column_offset in range(clipboard_width):
            got=int(excel.WorksheetFunction.CountA(
                sheet.Range(sheet.Cells(paste_next,paste_col+column_offset),
                            sheet.Cells(paste_next+matched_rows-1,paste_col+column_offset))))
            if got!=expected_columns[column_offset]:
                parts.append(f'{column_name(paste_col+column_offset)}列期望{expected_columns[column_offset]:,}格/实际{got:,}格')
                if len(parts)>=4: break
        if parts:
            message+='；不一致的列：'+'、'.join(parts)
    except Exception:
        pass
    return message


def _copy_dimensions(source, destination, first_col, first_row, rows, width):
    """Range.Copy keeps cell styles; transfer grouped row heights too. Column
    widths are never copied from the scratch sheet (it only has defaults);
    they come from the existing 表A1 or the optional format reference."""
    pending=[(first_row,first_row+rows-1)]
    while pending:
        start,end=pending.pop()
        span=f'{start}:{end}'
        height=source.Rows(span).RowHeight
        if height is not None:
            destination.Rows(span).RowHeight=height
        elif start==end:
            destination.Rows(span).RowHeight=source.Rows(span).Height
        else:
            middle=(start+end)//2
            pending.extend(((start,middle),(middle+1,end)))
