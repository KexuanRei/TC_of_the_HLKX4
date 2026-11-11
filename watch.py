"""监听表D页面的编辑动作，收集被改动的行号，供“改动即同步”使用。

金山文档没有对外提供变更通知接口，这里在专用 Edge 页里注入一段轻量监听脚本：
在 input/beforeinput/paste 等编辑事件发生时读取“名称框”（单元格地址框）里
显示的当前活动单元格/选区地址，把对应行号记入脏行集合。同步程序定期 drain 这些
行号，只读取这些行并按主键写回表A。

要点：金山文档的表格本体通常跑在 iframe 里，因此注入与 drain 都必须覆盖所有
frame，而不能只作用于主 frame。
"""

INSTALL_JS = r"""
(function () {
  if (window.__dsh_watch_version === 2) return;
  if (window.__dsh_watch_installed) return; // 旧脚本由重启专用Edge清除。
  window.__dsh_watch_installed = true;
  window.__dsh_watch_version = 2;
  var dirty = {};  // row number -> true

  function rowOf(addr) {
    var m = /^\$?[A-Za-z]{1,3}\$?(\d+)/.exec((addr || '').trim());
    return m ? parseInt(m[1], 10) : 0;
  }

  // 名称框：值形如 A1 / A1:B2，且不在单元格/对话框内部的可见输入框。
  function nameBoxOf(el) {
    if (!el || el.tagName !== 'INPUT') return null;
    var t = (el.type || 'text').toLowerCase();
    if (t === 'password' || t === 'hidden') return null;
    var v = (el.value || '').toUpperCase();
    if (!/^\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/.test(v)) return null;
    if (el.closest('[role="dialog"],[aria-modal="true"],[role="gridcell"],td,th')) return null;
    return v;
  }

  function activeAddress() {
    var inputs = document.querySelectorAll('input:not([type="password"]):not([type="hidden"])');
    for (var i = 0; i < inputs.length; i++) {
      var v = nameBoxOf(inputs[i]);
      if (v) return v;
    }
    return '';
  }

  function mark(ev) {
    // 忽略名称框自身输入（程序选区或用户手输地址），只认单元格编辑。
    var target = ev && ev.target;
    if (target && target.tagName === 'INPUT') {
      // 地址框被清空或输入到一半时，也不能当成单元格编辑。
      var hint = [target.id, target.className, target.getAttribute('name'),
                  target.getAttribute('aria-label'), target.title].join(' ');
      for (var p = target.parentElement, n = 0; p && n < 3; p = p.parentElement, n++) {
        hint += ' ' + p.className;
      }
      if (nameBoxOf(target) || /名称框|名称盒|单元格地址|name[-_\s]?box|range[-_\s]?(address|box)|(?:^|\s)edit-box(?:\s|$)/i.test(hint)) return;
    }
    var addr = activeAddress();
    if (!addr) return;
    var parts = addr.split(':');
    var top = rowOf(parts[0]);
    var bottom = rowOf(parts[1] || parts[0]);
    if (top > 0 && bottom >= top) {
      for (var r = top; r <= bottom; r++) dirty[r] = true;
    }
  }

  ['input', 'beforeinput', 'paste', 'cut', 'change'].forEach(function (ev) {
    document.addEventListener(ev, mark, true);
  });
  // keyup 会把 Ctrl+C、Enter、方向键以及程序选区误报为修改，形成自触发循环。
  // 删除和快捷键编辑从 keydown 记录，避免 Enter 提交后记录到下一行。
  document.addEventListener('keydown', function (ev) {
    var key = (ev.key || '').toLowerCase();
    if (ev.altKey) return;
    if (ev.ctrlKey || ev.metaKey) {
      if (['v', 'x', 'z', 'y'].indexOf(key) >= 0) mark(ev);
      return;
    }
    if (key === 'backspace' || key === 'delete' || key.length === 1) mark(ev);
  }, true);
  window.__dsh_drain_dirty = function () {
    var rows = Object.keys(dirty).map(Number).filter(function (n) { return n > 0; });
    dirty = {};
    return rows;
  };
  window.__dsh_watch_probe = function () {
    return {installed: true, namebox: activeAddress(), dirty: Object.keys(dirty).length};
  };
})();
"""


def _frame_targets(page):
    """页面所有 frame（含主 frame）。kdocs 表格本体通常在 iframe 里。"""
    try:
        return list(page.frames)
    except Exception:
        return [page]


def install(page):
    """把监听脚本注入页面所有 frame（幂等）。"""
    for target in _frame_targets(page):
        try:
            target.evaluate(INSTALL_JS)
        except Exception:
            pass


def drain_dirty_rows(page):
    """读取并清空各 frame 里累积的脏行号，返回升序整数列表。"""
    out = set()
    for target in _frame_targets(page):
        try:
            rows = target.evaluate('window.__dsh_drain_dirty ? window.__dsh_drain_dirty() : []')
        except Exception:
            continue
        if isinstance(rows, list):
            for value in rows:
                if isinstance(value, (int, float)) and 1 <= int(value) <= 1048576:
                    out.add(int(value))
    return sorted(out)


def probe(page):
    """诊断：返回各 frame 的监听状态与名称框识别结果。"""
    result = []
    for target in _frame_targets(page):
        try:
            data = target.evaluate('window.__dsh_watch_probe ? window.__dsh_watch_probe() : null')
        except Exception as exc:
            data = {'error': type(exc).__name__}
        result.append(data)
    return result
