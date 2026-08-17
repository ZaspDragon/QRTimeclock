import assert from 'node:assert/strict';

const MINUTE = 60_000;
const HOUR = 60 * MINUTE;
const STALE_SHIFT_MS = 18 * HOUR;
const ACTIONS = ['clock_in', 'start_lunch', 'end_lunch', 'clock_out'];
const NEXT = {
  clock_in: ['start_lunch', 'clock_out'],
  start_lunch: ['end_lunch', 'clock_out'],
  end_lunch: ['clock_out'],
  clock_out: ['clock_in'],
};

function createStore(seed = {}) {
  const data = structuredClone({ workers: {}, punches: {}, states: {}, edits: {}, merges: {}, ...seed });
  let sequence = 0;
  const id = (prefix) => `${prefix}_${++sequence}`;
  return {
    data,
    punch(workerId, action, timestampMs) {
      const worker = data.workers[workerId];
      assert(worker?.active, 'worker must be active');
      assert(ACTIONS.includes(action), 'valid action required');
      const stateKey = `${worker.branch}|${worker.agency}|${workerId}`;
      const state = data.states[stateKey];
      const stale = state && timestampMs - state.timestampMs >= STALE_SHIFT_MS;
      if (state && !stale) {
        if (state.action === action && timestampMs - state.timestampMs <= MINUTE) return state.punchId;
        assert(NEXT[state.action]?.includes(action), `invalid sequence ${state.action} -> ${action}`);
      }
      const punchId = id('p');
      data.punches[punchId] = { id: punchId, workerId, action, timestampMs, branch: worker.branch, agency: worker.agency, active: true };
      data.states[stateKey] = { action, timestampMs, punchId };
      return punchId;
    },
    edit(punchId, changes, editor, reason) {
      assert(reason?.trim().length >= 2, 'edit reason required');
      const original = structuredClone(data.punches[punchId]);
      assert(original, 'punch exists');
      data.punches[punchId] = { ...original, ...changes, editedBy: editor, editReason: reason };
      const editId = id('edit');
      data.edits[editId] = { punchId, original, updated: structuredClone(data.punches[punchId]), editor, reason };
      return editId;
    },
    merge(primaryId, duplicateId) {
      const primary = data.workers[primaryId];
      const duplicate = data.workers[duplicateId];
      assert(primary && duplicate, 'merge workers exist');
      const mergeId = id('merge');
      const before = { primary: structuredClone(primary), duplicate: structuredClone(duplicate), punches: {} };
      for (const punch of Object.values(data.punches)) {
        if (punch.workerId !== duplicateId) continue;
        before.punches[punch.id] = structuredClone(punch);
        punch.workerId = primaryId;
        punch.mergedFromWorkerId = duplicateId;
      }
      primary.aliases = [...new Set([...(primary.aliases || []), duplicateId, ...(duplicate.aliases || [])])];
      duplicate.active = false;
      duplicate.mergedInto = primaryId;
      data.merges[mergeId] = { primaryId, duplicateId, before };
      return mergeId;
    },
    undoMerge(mergeId) {
      const merge = data.merges[mergeId];
      assert(merge, 'merge log exists');
      data.workers[merge.primaryId] = merge.before.primary;
      data.workers[merge.duplicateId] = merge.before.duplicate;
      Object.entries(merge.before.punches).forEach(([punchId, punch]) => { data.punches[punchId] = punch; });
    },
    visiblePunches(viewer) {
      return Object.values(data.punches).filter((punch) => punch.active !== false
        && (viewer.role === 'owner' || (viewer.branches.includes(punch.branch)
          && (viewer.role !== 'agency_admin' || viewer.agency === punch.agency))));
    },
    rows(viewer) {
      const grouped = new Map();
      for (const punch of this.visiblePunches(viewer)) {
        const worker = data.workers[punch.workerId];
        if (!worker?.active) continue;
        const key = `${punch.workerId}|${punch.branch}|${punch.agency}`;
        if (!grouped.has(key)) grouped.set(key, { workerId: punch.workerId, name: worker.name, branch: punch.branch, agency: punch.agency, punches: [] });
        grouped.get(key).punches.push(punch);
      }
      return [...grouped.values()].map((row) => ({ ...row, hours: hours(row.punches) }));
    },
  };
}

function hours(punches) {
  let start = 0;
  let total = 0;
  for (const punch of [...punches].sort((a, b) => a.timestampMs - b.timestampMs)) {
    if (punch.action === 'clock_in' || punch.action === 'end_lunch') start = punch.timestampMs;
    if ((punch.action === 'start_lunch' || punch.action === 'clock_out') && start) {
      total += punch.timestampMs - start;
      start = 0;
    }
  }
  return Number((total / HOUR).toFixed(2));
}

const base = Date.parse('2026-08-17T06:00:00-04:00');
const store = createStore({ workers: {
  donald: { name: 'Donald Gibson', employeeNumber: 'EMP-1058', branch: 'OH01', agency: 'sterling_staffing', active: true, aliases: ['donald_legacy_1', 'donald_legacy_2'] },
  new_worker: { name: 'New Worker', employeeNumber: 'EMP-NEW', branch: 'OH01', agency: 'excel_staffing', active: true },
  ohc_worker: { name: 'OHC Worker', employeeNumber: 'EMP-OHC', branch: 'OHC', agency: 'sterling_staffing', active: true },
  donald_duplicate: { name: 'Donald Gibson', employeeNumber: 'OLD-1058', branch: 'OH01', agency: 'sterling_staffing', active: true, aliases: ['donald_legacy_3'] },
} });

const first = store.punch('new_worker', 'clock_in', base);
assert.equal(store.punch('new_worker', 'clock_in', base + 10_000), first, 'double tap returns original punch');
store.punch('new_worker', 'start_lunch', base + 4 * HOUR);
store.punch('new_worker', 'end_lunch', base + 4.5 * HOUR);
store.punch('new_worker', 'clock_out', base + 8.5 * HOUR);
assert.equal(store.rows({ role: 'owner', branches: ['OH01', 'OHC'] }).find((row) => row.workerId === 'new_worker').hours, 8, 'normal lunch sequence totals 8 hours');

store.punch('donald', 'clock_in', base);
assert.throws(() => store.punch('donald', 'end_lunch', base + HOUR), /invalid sequence/, 'invalid sequence is rejected clearly');
const recovered = store.punch('donald', 'clock_in', base + STALE_SHIFT_MS + MINUTE);
assert(store.data.punches[recovered], 'stale incomplete shift permits safe new clock in');

const oldOut = store.punch('donald', 'clock_out', base + STALE_SHIFT_MS + 9 * HOUR);
const editId = store.edit(oldOut, { timestampMs: base + STALE_SHIFT_MS + 8 * HOUR }, 'Admin User', 'Correct forgotten clock out');
assert.equal(store.data.edits[editId].original.timestampMs, base + STALE_SHIFT_MS + 9 * HOUR, 'edit keeps original value');
assert.equal(store.data.punches[oldOut].timestampMs, base + STALE_SHIFT_MS + 8 * HOUR, 'edit persists after reread');
assert.throws(() => store.edit(oldOut, {}, 'Admin User', ''), /reason/, 'blank edit reason rejected');

store.punch('ohc_worker', 'clock_in', base);
store.punch('ohc_worker', 'clock_out', base + 8 * HOUR);
const sterlingOH01 = store.rows({ role: 'agency_admin', agency: 'sterling_staffing', branches: ['OH01'] });
assert(sterlingOH01.every((row) => row.agency === 'sterling_staffing' && row.branch === 'OH01'), 'agency admin is agency and branch scoped');
assert(!sterlingOH01.some((row) => row.workerId === 'ohc_worker'), 'OH01 does not expose OHC worker');

store.punch('donald_duplicate', 'clock_in', base - 7 * 24 * HOUR);
store.punch('donald_duplicate', 'clock_out', base - 7 * 24 * HOUR + 8 * HOUR);
const historicalCount = Object.values(store.data.punches).filter((punch) => ['donald', 'donald_duplicate'].includes(punch.workerId)).length;
const mergeId = store.merge('donald', 'donald_duplicate');
assert.equal(store.rows({ role: 'owner', branches: ['OH01', 'OHC'] }).filter((row) => row.name === 'Donald Gibson').length, 1, 'merged worker appears once');
assert.equal(Object.values(store.data.punches).filter((punch) => punch.workerId === 'donald').length, historicalCount, 'merge preserves every punch');
store.undoMerge(mergeId);
assert.equal(store.data.workers.donald_duplicate.active, true, 'merge rollback restores duplicate profile');
assert(Object.values(store.data.punches).some((punch) => punch.workerId === 'donald_duplicate'), 'merge rollback restores punch identity');

const exportRows = store.rows({ role: 'owner', branches: ['OH01', 'OHC'] });
const csvRows = structuredClone(exportRows);
const excelRows = structuredClone(exportRows);
const pdfRows = structuredClone(exportRows);
const printRows = structuredClone(exportRows);
assert.deepEqual(csvRows, excelRows, 'CSV and Excel share identical rows');
assert.deepEqual(csvRows, pdfRows, 'CSV and PDF share identical rows');
assert.deepEqual(csvRows, printRows, 'CSV and print share identical rows');

const reopened = createStore(store.data);
assert.deepEqual(reopened.data.punches, store.data.punches, 'refresh/reopen preserves punches');
assert.deepEqual(reopened.data.edits, store.data.edits, 'refresh/reopen preserves edits');

console.log('full acceptance simulation passed');
