import copy
import json
import os
import shutil
import subprocess
import unittest
from unittest.mock import patch

import sync_runner
import watch
from clipboard_html import CopyIntegrityError
import test_sync_flow
from update_excel import address


class SyncRegressionTests(unittest.TestCase):
    def setUp(self):
        test_sync_flow.SyncFlowTests.setUp(self)
        timeout = patch.object(sync_runner, 'READBACK_TIMEOUT_SECONDS', 0)
        timeout.start()
        self.addCleanup(timeout.stop)

    run_sync = test_sync_flow.SyncFlowTests.run_sync

    def prepare(self, count=6, batch=6):
        self.copier.tables['d'] = [['id', 'value']] + [[f'key{i}', f'new{i}'] for i in range(count)]
        self.copier.tables['a'] = [['id', 'value']] + [[f'key{i}', f'old{i}'] for i in range(count)]
        self.config['source'].update(scan_range=f'A1:B{count+1}', batch_rows=1000, max_cells=1000000)
        self.config['upload']['scan_range'] = f'A1:B{count+1}'
        self.config['sync']['batch_rows'] = batch

    def limit_copy(self, cap, silent=False):
        original = self.copier._capture
        def capture(page, ref):
            c1, r1 = address(ref.split(':')[0])
            c2, r2 = address(ref.split(':')[1])
            if r2 - r1 + 1 > cap:
                if silent:
                    from copy_range import column_name
                    return original(page, f'{column_name(c1)}{r1}:{column_name(c2)}{r1+cap-1}')
                raise CopyIntegrityError('incomplete', can_split=True, kind='rows')
            return original(page, ref)
        self.copier._capture = capture

    def test_split_reads_never_paste_only_last_clipboard_chunk(self):
        self.prepare()
        self.limit_copy(2)
        result = self.run_sync()
        self.assertEqual(result['modified'], 6)
        self.assertEqual(self.copier.tables['a'], self.copier.tables['d'])
        self.assertTrue(all(len(rows) <= 2 for rows in self.copier.paste_blocks))
        self.assertEqual(self.run_sync()['modified'], 0)

    def test_silent_short_clipboard_is_split_not_padded_with_blank_data(self):
        self.prepare()
        self.limit_copy(2, silent=True)
        self.assertEqual(self.run_sync()['modified'], 6)
        self.assertEqual(self.copier.tables['a'], self.copier.tables['d'])
        self.assertEqual(self.run_sync()['modified'], 0)

    def test_target_index_reads_only_primary_key_and_passes_1000_rows(self):
        self.prepare(2050, 1000)
        self.copier.tables['a'] = copy.deepcopy(self.copier.tables['d'])
        self.copier.tables['a'][-1][1] = 'outdated'
        result = self.run_sync()
        self.assertEqual((result['scanned'], result['modified'], result['missing']), (2050, 1, 0))
        reads = [e[2] for e in self.copier.events if e[:2] == ('read', 'a')]
        self.assertEqual(reads[:3], ['A1:A1000', 'A1001:A2000', 'A2001:A2051'])
        self.assertEqual(self.copier.tables['a'][-1], self.copier.tables['d'][-1])

    def test_targeted_batches_obey_size_limit(self):
        batches = list(sync_runner._row_batches(1, 20, 3, [2,3,4,5,6,7,8,12]))
        self.assertEqual(batches, [(2,4), (5,7), (8,8), (12,12)])

    def test_cached_index_invalidated_when_sheet_changes(self):
        self.prepare(2, 2)
        self.run_sync()
        self.copier.events.clear()
        self.config['upload']['sheet_name'] = 'new sheet'
        sync_runner.run_sync(self.copier, self.config, only_rows=[2])
        self.assertIn(('read', 'a', 'A1:A3'), self.copier.events)


class IndexRegressionTests(unittest.TestCase):
    def test_full_a2_sized_index_reaches_last_range(self):
        copier = test_sync_flow._ShrinkCopier(fail_when_rows_gt=1000)
        grid = sync_runner.read_online_grid(copier, None,
            {'source': {'batch_rows':1000, 'max_cells':1000000}}, 'D5:D48306', '表A')
        self.assertEqual(grid['row_count'], 48302)
        self.assertEqual(grid['rows'][-1], ['row48306'])
        self.assertIn('D48005:D48306', copier.messages[-1])


class WatchEventRegressionTests(unittest.TestCase):
    def test_copy_and_selection_do_not_trigger_sync_but_edits_do(self):
        node = os.environ.get('NODE_BINARY') or shutil.which('node')
        if not node:
            self.skipTest('Node.js is required for actual watcher event tests')
        script = r'''
const vm = require('node:vm'), assert = require('node:assert/strict');
let listeners = {};
const box = {tagName:'INPUT', type:'text', value:'D7', className:'edit-box',
    getAttribute(){return '';}, closest(){return null;}};
const cell = {tagName:'TEXTAREA'};
const window = {};
const document = {querySelectorAll(){return [box];},
    addEventListener(type, fn){(listeners[type] ||= []).push(fn);}};
vm.runInNewContext(JSON.parse(process.argv[2]), {window,document});
function event(type, props={}) {for (const fn of listeners[type] || []) fn({target:cell,...props});}
function dirty() {return Array.from(window.__dsh_drain_dirty());}
for (const key of ['Control','c','Enter','ArrowDown','Shift']) {
    event('keydown', {key,ctrlKey:key==='c'}); event('keyup', {key});
}
assert.deepEqual(dirty(), []);
box.value=''; event('input', {target:box}); box.value='D7';
event('keyup', {key:'Enter'}); assert.deepEqual(dirty(), []);
for (const key of ['Delete','Backspace','a']) {
    event('keydown', {key}); assert.deepEqual(dirty(), [7]);
}
for (const key of ['v','x','z','y']) {
    event('keydown', {key,ctrlKey:true}); assert.deepEqual(dirty(), [7]);
}
box.value='A20:B22'; event('paste'); assert.deepEqual(dirty(), [20,21,22]);
event('input'); assert.deepEqual(dirty(), [20,21,22]);
assert.deepEqual(dirty(), []);
'''
        result = subprocess.run([node, '-e', script, '--', 'watch', json.dumps(watch.INSTALL_JS)],
                                capture_output=True, text=True)
        self.assertEqual(result.returncode, 0, result.stderr)
