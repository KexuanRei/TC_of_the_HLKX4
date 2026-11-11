"""按自定义批大小读取表D，按主键合并复制连续差异行，再继续下一批。"""
import json
from pathlib import Path
import time
import urllib.parse

from copy_range import column_name
from update_excel import address
from excel_clipboard import inspect_copy
from clipboard_html import CopyIntegrityError
from upload_verify import snapshot, differences, mismatch
from sync_engine import parse_pk_column, data_start_row, sync_batch_rows, build_index, pk_key, MAX_BATCH_ROWS
from watch import install as install_watch

# 同一处不一致持续这么久才判定失败（WPS 粘贴后异步渲染）。
MISMATCH_SETTLE_SECONDS = 8.0
READBACK_TIMEOUT_SECONDS = 30.0


def read_online_grid(copier, page, config, ref, label):
    """把在线表读成二维网格（rows 为按表头行排序的值列表，header 为表头）。"""
    cfg = config['source']
    if ref:
        left, right = ref.split(':')
        c1, r1 = address(left)
        c2, r2 = address(right)
        if c2 < c1 or r2 < r1:
            raise ValueError(f'{label}扫描区域无效')
        batch_rows = min(max(10, int(cfg.get('batch_rows') or 1000)), MAX_BATCH_ROWS)
        width = c2 - c1 + 1
        rows, header = [], None
        cursor = r1
        while cursor <= r2:
            count = min(batch_rows, r2 - cursor + 1)
            while True:
                end = cursor + count - 1
                bref = f'{column_name(c1)}{cursor}:{column_name(c2)}{end}'
                copier.update(message=f'正在扫描{label} {bref}…')
                try:
                    clip = copier._capture(page, bref)
                    grid = inspect_copy(clip, config, f'{column_name(c1)}{cursor}', 'selection',
                                        expected_range=bref, batch=True, strict_rows=True)
                except CopyIntegrityError as exc:
                    if exc.can_split and count > 1:
                        count = max(1, count // 2)
                        continue
                    raise
                if header is None:
                    header = grid['header']
                rows.extend(grid['rows'])
                cursor += count
                break
        return dict(rows=rows, header=header, width=width,
                    first_column=c1, first_row=r1, row_count=len(rows))
    # 整表一次读取（适合小表；大表请配置扫描区域）。
    copier.update(message=f'正在扫描{label}整张工作表…')
    clip = copier._capture(page, None)
    grid = inspect_copy(clip, config, 'A1', 'all')
    return dict(rows=grid['rows'], header=grid['header'], width=grid['width'],
                first_column=grid['first_column'], first_row=grid['first_row'],
                row_count=grid['row_count'])


def _verify_rows(copier, page, ref, expected, max_cells, label):
    """粘贴后持续读回直到稳定；不一致才报错（WPS 异步渲染）。"""
    origin = ref.split(':')[0]
    start = time.monotonic()
    last = None
    first_seen = None
    samples = 0
    while True:
        now = time.monotonic()
        copier.update_upload(message=f'正在回读校验表A {ref}（仅检查，不重复粘贴）…')
        actual = _read_snapshot_split(copier, page, ref, max_cells)
        found = differences(expected, actual, origin)
        if not found:
            return actual
        diff = mismatch(expected, actual, origin) if found else None
        if diff != last:
            last, first_seen, samples = diff, now, 1
        else:
            samples += 1
        stable = now - first_seen if first_seen is not None else 0.0
        if (samples >= 2 and stable >= MISMATCH_SETTLE_SECONDS) or now - start >= READBACK_TIMEOUT_SECONDS:
            raise ValueError(f'表A {label} 回读校验失败：{diff}')
        page.wait_for_timeout(min(400 + 400 * samples, 1500))


def _open_sheet(copier, config, which):
    from online_format import OnlineFormatter
    section, label = ('source', '表D') if which == 'd' else ('upload', '表A')
    opener = copier._page_d if which == 'd' else copier._page_a
    page = opener(config[section]['share_url'], reload=True)
    page.bring_to_front()
    install_watch(page)
    url = urllib.parse.urlparse(page.url)
    if url.scheme != 'https' or url.hostname not in {'www.kdocs.cn', 'kdocs.cn'}:
        raise ValueError(f'请先打开{label}并登录，完成后再同步')
    copier.context.grant_permissions(['clipboard-read', 'clipboard-write'],
                                     origin=url.scheme + '://' + url.netloc)
    formatter = OnlineFormatter(page, lambda ref: copier._select(page, ref),
                                progress=lambda text: copier.update(message=text.replace('表D', label)))
    formatter.select_sheet(config[section]['sheet_name'])
    return page


def _bounds(copier, page, config):
    ref = config['source'].get('scan_range')
    if ref:
        left, right = ref.split(':')
        c1, r1 = address(left)
        c2, r2 = address(right)
    else:
        # 未指定范围时，仅用整表复制识别边界；实际比较仍逐行重新读取。
        grid = read_online_grid(copier, page, config, None, '表D（识别范围）')
        c1, r1 = grid['first_column'], grid['first_row']
        c2, r2 = c1 + grid['width'] - 1, r1 + grid['row_count'] - 1
    if c1 < 1 or c2 < c1 or r2 < r1:
        raise ValueError('表D扫描范围为空或无效')
    start = max(r1, data_start_row(config))
    return c1, start, c2, r2


def _read_snapshot_split(copier, page, ref, max_cells):
    """读取 ref 范围的快照；一次复制不完整时对半缩小批次重读并合并。

    与 snapshot() 的区别：遇到“纯文本中段丢行”这类可缩小批次的错误时自动拆分重读，
    而不是让整轮同步中断。返回的 rows 按行号升序拼接，merges 为各子块的并集。
    """
    left, right = ref.split(':')
    c1, r1 = address(left)
    c2, r2 = address(right)
    todo = [(r1, r2)]
    chunks = []
    while todo:
        lo, hi = todo.pop()
        sub = f'{column_name(c1)}{lo}:{column_name(c2)}{hi}'
        try:
            result = snapshot(copier._capture(page, sub), sub, max_cells, strict_rows=True)
        except CopyIntegrityError as exc:
            if exc.can_split and hi > lo:
                mid = (lo + hi) // 2
                todo.append((mid + 1, hi))
                todo.append((lo, mid))
                continue
            raise
        chunks.append((lo, result))
    chunks.sort(key=lambda item: item[0])
    rows, merges = [], set()
    unverified = 0
    for _lo, result in chunks:
        rows.extend(result['rows'])
        merges.update((left, top + _lo - r1, right, bottom + _lo - r1)
                      for left, top, right, bottom in result['merges'])
        unverified += int(result.get('unverified_rows') or 0)
    return dict(rows=rows, row_count=len(rows), width=c2 - c1 + 1,
                merges=merges, unverified_rows=unverified)


def _read_block(copier, page, ref, max_cells):
    result = _read_snapshot_split(copier, page, ref, max_cells)
    if any(bottom != top for _left, top, _right, bottom in result['merges']):
        raise ValueError(f'{ref} 含跨行合并单元格，无法安全按主键分批复制')
    return result


def _record(copier, **values):
    """发现差异后立刻持久化主键和游标；不记录整行内容。"""
    path = Path(copier.root) / 'sync_progress.json'
    temp = path.with_suffix('.tmp')
    temp.write_text(json.dumps(values, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


def _groups(items, consecutive):
    """把有序条目拆成可连续读取或粘贴的矩形组。"""
    group = []
    for item in items:
        if group and not consecutive(group[-1], item):
            yield group
            group = []
        group.append(item)
    if group:
        yield group


def _row_batches(start, end, batch_rows, only_rows=None):
    """把要处理的行拆成 (first,last) 区间；only_rows 为 None 时按批全表扫描。"""
    if only_rows is None:
        cursor = start
        while cursor <= end:
            yield cursor, min(end, cursor + batch_rows - 1)
            cursor += batch_rows
        return
    rows = sorted(r for r in set(only_rows) if start <= r <= end)
    if not rows:
        return
    first = prev = rows[0]
    for r in rows[1:]:
        if r <= prev + 1 and r - first < batch_rows:
            prev = r
            continue
        yield first, prev
        first = prev = r
    yield first, prev


def run_sync(copier, config, only_rows=None):
    """读一批、按主键比对并写完这一批，再继续；全表结束后由调度器等待。

    only_rows：要同步的指定行号列表（用于“改动即同步”的即时路径）；为 None 时全表扫描。
    """
    pk_name, pk_number = parse_pk_column(config)
    batch_rows = sync_batch_rows(config)
    data_start = data_start_row(config)
    max_cells = config['source']['max_cells']
    page_d = _open_sheet(copier, config, 'd')
    c1, start, c2, end = _bounds(copier, page_d, config)
    if not c1 <= pk_number <= c2:
        raise ValueError(f'表D扫描区域必须包含主键列 {pk_name}')
    page_a = _open_sheet(copier, config, 'a')
    pk_offset, width = pk_number - c1, c2 - c1 + 1

    def ref(first, last):
        return f'{column_name(c1)}{first}:{column_name(c2)}{last}'

    def target_map():
        scan_range = config['upload'].get('scan_range') or None
        if scan_range:
            left, right = scan_range.split(':')
            ac1, ar1 = address(left)
            ac2, ar2 = address(right)
            if not ac1 <= pk_number <= ac2:
                raise ValueError(f'表A扫描区域必须包含主键列 {pk_name}')
            # 索引只需要主键，避免其他隐藏列、格式和宽表复制限制影响索引。
            scan_range = f'{pk_name}{ar1}:{pk_name}{ar2}'
        grid = read_online_grid(copier, page_a, config, scan_range, '表A主键')
        first = grid['first_column']
        if not first <= pk_number < first + grid['width']:
            raise ValueError(f'表A扫描区域必须包含主键列 {pk_name}')
        index = build_index(grid, pk_name, pk_number, data_start,
                            config['upload'].get('scan_range') or '', label='表A')
        return {key: row_no for key, (_row, row_no) in index['rows'].items()}

    # 即时路径复用已建立的表A主键索引（表A行位置只被覆盖、不增删，因此索引稳定；
    # 若表A被人为改动导致错位，下方 moved 检测会重建）。
    cache_key = (config['upload'].get('share_url'), config['upload'].get('sheet_name'),
                 config['upload'].get('scan_range'), pk_name, data_start)
    a_map = None
    if only_rows is not None and getattr(copier, '_target_index_url', None) == cache_key:
        a_map = copier._target_index
    if a_map is None:
        a_map = target_map()
        copier._target_index = a_map
        copier._target_index_url = cache_key
    seen = {}
    changed = scanned = missing = blank = paste_batches = 0
    for batch_start, batch_end in _row_batches(start, end, batch_rows, only_rows):
        ref_d = ref(batch_start, batch_end)
        copier.update(message=f'正在读取表D第{batch_start}～{batch_end}行（每批{batch_rows}行，已更新{changed}行）…',
                      scan_row=batch_start, scan_end=end, batch_end=batch_end)
        wanted = _read_block(copier, page_d, ref_d, max_cells)
        scanned += wanted['row_count']
        entries = []
        for offset, row in enumerate(wanted['rows']):
            source_row = batch_start + offset
            key = pk_key(row[pk_offset])
            if key is None:
                blank += 1
                continue
            if key in seen:
                raise ValueError(f'表D主键列 {pk_name} 存在重复值“{key}”（第{seen[key]}行与第{source_row}行），已停止本轮')
            seen[key] = source_row
            entries.append(dict(pk=key, source=source_row, row=row))

        # 按表A实际行号合并读取，避免每一行都切换浏览器和复制剪贴板。
        for attempt in range(2):
            matched = [dict(entry, target=a_map[entry['pk']]) for entry in entries if entry['pk'] in a_map]
            actual_rows = {}
            moved = False
            for group in _groups(sorted(matched, key=lambda x: x['target']),
                                 lambda a, b: b['target'] == a['target'] + 1):
                actual = _read_block(copier, page_a, ref(group[0]['target'], group[-1]['target']), max_cells)
                for entry, row in zip(group, actual['rows']):
                    if pk_key(row[pk_offset]) != entry['pk']:
                        moved = True
                        break
                    actual_rows[entry['pk']] = row
                if moved:
                    break
            if not moved:
                break
            if attempt:
                raise ValueError('表A行位置持续变化，已停止本轮，请勿在同步时排序或插入行')
            a_map = target_map()
            copier._target_index = a_map
            copier._target_index_url = cache_key

        for entry in entries:
            if entry['pk'] not in a_map:
                missing += 1
                _record(copier, pk=entry['pk'], source_row=entry['source'], target_row=None,
                        status='missing', next_source_row=entry['source'] + 1)

        pending = []
        for entry in matched:
            expected = dict(rows=[entry['row']], row_count=1, width=width)
            actual = dict(rows=[actual_rows[entry['pk']]], row_count=1, width=width)
            if differences(expected, actual, f'{column_name(c1)}{entry["target"]}', limit=1):
                pending.append(entry)

        # 只有两表中的行均连续且顺序一致才能直接整块粘贴；不跨过无变化行。
        write_groups = list(_groups(pending, lambda a, b: b['source'] == a['source'] + 1
                                    and b['target'] == a['target'] + 1))
        write_groups.reverse()
        while write_groups:
            group = write_groups.pop()
            first, last = group[0], group[-1]
            source_ref, target_ref = ref(first['source'], last['source']), ref(first['target'], last['target'])
            progress = dict(pk=first['pk'], primary_keys=[e['pk'] for e in group],
                            source_row=first['source'], source_end=last['source'],
                            target_row=first['target'], target_end=last['target'])
            _record(copier, **progress, status='pending', next_source_row=first['source'])
            copier.update_upload(message=f'正在批量同步{len(group)}行：表D {source_ref} → 表A {target_ref}…')
            # 写入前再次确认目标主键，避免按已失效的行号覆盖其他记录。
            target = _read_block(copier, page_a, target_ref, max_cells)
            if [pk_key(row[pk_offset]) for row in target['rows']] != progress['primary_keys']:
                raise ValueError('表A行位置在写入前发生变化，已停止本轮，下轮重新按主键定位')
            expected = dict(rows=[e['row'] for e in group], row_count=len(group), width=width)
            # 必须从一次完整复制直接粘贴。拆批读取的合并快照只在内存中完整，
            # 系统剪贴板实际上仅保留最后一块，不能把它当作整组粘贴。
            try:
                fresh = snapshot(copier._capture(page_d, source_ref), source_ref,
                                 max_cells, strict_rows=True)
            except CopyIntegrityError as exc:
                if exc.can_split and len(group) > 1:
                    middle = len(group) // 2
                    write_groups.append(group[middle:])
                    write_groups.append(group[:middle])
                    copier.update_upload(message=f'{source_ref} 复制不完整，拆成较小块分别复制、粘贴和校验…')
                    continue
                raise
            if any(bottom != top for _left, top, _right, bottom in fresh['merges']):
                raise ValueError(f'{source_ref} 含跨行合并单元格，已停止本轮')
            if differences(expected, fresh, source_ref.split(':')[0], limit=1):
                raise ValueError(f'表D {source_ref} 在复制前发生变化，已停止本轮，下轮重试')
            page_a.bring_to_front()
            copier._select(page_a, f'{column_name(c1)}{first["target"]}')
            page_a.keyboard.press('Control+v')
            page_a.wait_for_timeout(500)
            _verify_rows(copier, page_a, target_ref, fresh, max_cells, target_ref)
            changed += len(group)
            paste_batches += 1
            _record(copier, **progress, status='verified', next_source_row=last['source'] + 1)
        copier.update(message=f'已完成表D第{batch_start}～{batch_end}行，累计扫描{scanned}行、更新{changed}行；继续下一批…',
                      scan_row=batch_end, scan_end=end, batch_end=batch_end)
        page_d.bring_to_front()
        copier._select(page_d, ref_d)
    page_d.bring_to_front()
    result = dict(baseline=False, scanned=scanned, modified=changed, added=0, deleted=0,
                  rows=changed, writes=changed, deletes=0, missing=missing, blank=blank,
                  batch_rows=batch_rows, paste_batches=paste_batches, targeted=only_rows is not None)
    result['message'] = (f'{"即时同步" if only_rows is not None else "本轮扫描"}完成：处理{scanned}行，'
                         f'更新{changed}行（{paste_batches}次粘贴）；表A无对应主键{missing}行，空主键{blank}行。'
                         + ('继续监听表D变化。' if only_rows is not None
                            else '定时任务启用时，30秒后从表D起始位置重新扫描。'))
    copier.update(message=result['message'])
    return result
