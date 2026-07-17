import assert from 'node:assert/strict';

const HOUR = 60 * 60 * 1000;
const DAY = 24 * HOUR;
const DUPLICATE_WINDOW_MS = 60 * 1000;
const DISPLAY_TOLERANCE_MS = 60 * 1000;

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeWorkerNumber(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function monday(value) {
  const date = new Date(value);
  date.setHours(0, 0, 0, 0);
  const day = date.getDay();
  date.setDate(date.getDate() - (day === 0 ? 6 : day - 1));
  return date;
}

function dateKey(ms) {
  return new Date(ms).toISOString().slice(0, 10);
}

function timestampMs(row) {
  if (Number.isFinite(Number(row.timestampMs))) return Number(row.timestampMs);
  if (row.timestamp instanceof Date) return row.timestamp.getTime();
  if (row.createdAt instanceof Date) return row.createdAt.getTime();
  return 0;
}

function branchMatch(row, selectedBranch) {
  if (row.siteId) return row.siteId === selectedBranch;
  if (row.assignedSiteId) return row.assignedSiteId === selectedBranch;
  if (row.branchId) return row.branchId === selectedBranch;
  if (Array.isArray(row.siteIds) && row.siteIds.length) return row.siteIds.includes(selectedBranch);
  return true;
}

function normalizePunch(row) {
  const ms = timestampMs(row);
  return {
    ...row,
    timestampMs: ms,
    dateKey: row.dateKey || dateKey(ms),
    weekKey: row.weekKey || dateKey(monday(ms)),
    nameKey: row.nameKey || normalizeName(row.name || row.employeeName || ''),
  };
}

function punchIdentity(row) {
  if (row.employeeId) return `employee:${row.employeeId}`;
  if (row.employeeNumber) return `number:${normalizeWorkerNumber(row.employeeNumber)}`;
  return `name:${normalizeName(row.nameKey || row.name || '')}`;
}

function dedupePunches(rows) {
  const buckets = new Map();
  for (const row of rows.map(normalizePunch)) {
    const key = [
      punchIdentity(row),
      row.action,
      row.companyId || 'chadwell',
      row.siteId || row.assignedSiteId || row.branchId || '',
      row.agencyId || '',
      row.active === false || row.status === 'deleted' ? 'deleted' : 'active',
    ].join('|');
    if (!buckets.has(key)) buckets.set(key, []);
    buckets.get(key).push(row);
  }
  const groups = [];
  const deduped = [];
  for (const [key, bucket] of buckets.entries()) {
    const sorted = bucket.sort((a, b) => a.timestampMs - b.timestampMs);
    let group = [];
    let anchor = 0;
    const flush = () => {
      if (!group.length) return;
      if (group.length > 1) {
        groups.push(group);
        deduped.push(group.slice().sort((a, b) => completeness(b) - completeness(a))[0]);
      } else {
        deduped.push(group[0]);
      }
      group = [];
      anchor = 0;
    };
    for (const row of sorted) {
      if (!group.length || row.timestampMs - anchor <= DISPLAY_TOLERANCE_MS) {
        if (!group.length) anchor = row.timestampMs;
        group.push(row);
      } else {
        flush();
        group.push(row);
        anchor = row.timestampMs;
      }
    }
    flush();
  }
  return { deduped, groups };
}

function completeness(row) {
  return ['id', 'employeeId', 'employeeNumber', 'name', 'weekKey', 'dateKey', 'siteId', 'source']
    .filter((field) => row[field]).length;
}

async function loadWeeklySimulation({ weekStart, optimized, compatible, timesheets }) {
  const results = await Promise.allSettled([optimized(), compatible(), timesheets()]);
  const optimizedRows = results[0].status === 'fulfilled' ? results[0].value : [];
  const compatibleRows = results[1].status === 'fulfilled' ? results[1].value : [];
  const saved = results[2].status === 'fulfilled' ? results[2].value : [];
  return {
    punches: dedupePunches([...optimizedRows, ...compatibleRows]).deduped,
    timesheets: saved,
    metadata: {
      optimizedPunchQueryFailed: results[0].status === 'rejected',
      compatibleRowsLoaded: compatibleRows.length,
      timesheetsLoaded: saved.length,
      allHistoricalQueriesFailed: results.every((result) => result.status === 'rejected'),
    },
    weekKey: dateKey(monday(weekStart)),
  };
}

function createPunchStore() {
  const guards = new Map();
  const states = new Map();
  const punches = new Map();
  return {
    save(payload) {
      const window = Math.floor(payload.timestampMs / DUPLICATE_WINDOW_MS);
      const guardKey = [payload.companyId, payload.siteId, payload.employeeId, payload.action, window].join('|');
      const stateKey = [payload.companyId, payload.siteId, payload.employeeId].join('|');
      if (guards.has(guardKey)) return { status: 'duplicate', punch: guards.get(guardKey) };
      const state = states.get(stateKey);
      if (state?.lastAction === payload.action) {
        if (payload.timestampMs - state.lastPunchAtMs <= DUPLICATE_WINDOW_MS) {
          return { status: 'duplicate', punch: state.lastPunch };
        }
        return { status: 'invalid_sequence', punch: state.lastPunch };
      }
      const punch = { id: guardKey, ...payload };
      guards.set(guardKey, punch);
      states.set(stateKey, { lastAction: payload.action, lastPunchAtMs: payload.timestampMs, lastPunch: punch });
      punches.set(guardKey, punch);
      return { status: 'created', punch };
    },
    count: () => punches.size,
  };
}

function resolveEmployee({ typedName, employeeId = '', employeeNumber = '', employees }) {
  if (employeeId) return employees.find((employee) => employee.employeeId === employeeId || employee.id === employeeId) || null;
  if (employeeNumber) return employees.find((employee) => normalizeWorkerNumber(employee.employeeNumber) === normalizeWorkerNumber(employeeNumber)) || null;
  const matches = employees.filter((employee) => normalizeName(employee.name) === normalizeName(typedName));
  return matches.length === 1 ? matches[0] : null;
}

const baseMonday = monday('2026-07-13T12:00:00Z').getTime();
const monthsAgo = monday('2026-03-02T12:00:00Z').getTime();
const employee = { id: 'emp_1', employeeId: 'emp_1', employeeNumber: '100', name: 'Jordan Smith' };

const current = await loadWeeklySimulation({
  weekStart: baseMonday,
  optimized: async () => [{ ...employee, action: 'clock_in', timestampMs: baseMonday + HOUR, companyId: 'chadwell', siteId: 'OH01' }],
  compatible: async () => [],
  timesheets: async () => [],
});
assert.equal(current.punches.length, 1, '1 current week punches display');

const lastWeek = await loadWeeklySimulation({
  weekStart: baseMonday - 7 * DAY,
  optimized: async () => [],
  compatible: async () => [{ ...employee, action: 'clock_in', timestampMs: baseMonday - 7 * DAY + HOUR, companyId: 'chadwell', siteId: 'OH01' }],
  timesheets: async () => [],
});
assert.equal(lastWeek.punches.length, 1, '2 last week punches display');

const oldWeek = await loadWeeklySimulation({
  weekStart: monthsAgo,
  optimized: async () => [],
  compatible: async () => [{ ...employee, action: 'clock_out', timestampMs: monthsAgo + 8 * HOUR, companyId: 'chadwell', siteId: 'OH01' }],
  timesheets: async () => [],
});
assert.equal(oldWeek.punches[0].weekKey, dateKey(monday(monthsAgo)), '3 months-old week displays');

assert.equal(normalizePunch({ timestampMs: baseMonday, weekKey: dateKey(monday(baseMonday)) }).weekKey, dateKey(monday(baseMonday)), '4 weekKey and timestampMs preserved');
assert.equal(normalizePunch({ timestampMs: baseMonday }).weekKey, dateKey(monday(baseMonday)), '5 timestampMs without weekKey derives week');
assert.equal(normalizePunch({ timestamp: new Date(baseMonday) }).timestampMs, baseMonday, '6 timestamp without timestampMs derives ms');

const fallback = await loadWeeklySimulation({
  weekStart: baseMonday,
  optimized: async () => { throw new Error('missing index'); },
  compatible: async () => [{ ...employee, action: 'clock_in', timestampMs: baseMonday + HOUR, companyId: 'chadwell', siteId: 'OH01' }],
  timesheets: async () => [],
});
assert.equal(fallback.punches.length, 1, '7 optimized failure does not block fallback');
assert.equal(fallback.metadata.optimizedPunchQueryFailed, true, '7 optimized failure is reported');

const savedOnly = await loadWeeklySimulation({
  weekStart: baseMonday,
  optimized: async () => { throw new Error('permission denied'); },
  compatible: async () => { throw new Error('missing legacy index'); },
  timesheets: async () => [{ id: 'ts1', weekKey: dateKey(monday(baseMonday)), status: 'signed', managerSignedBy: 'Manager' }],
});
assert.equal(savedOnly.timesheets.length, 1, '8 saved timesheet displays when punch queries fail');

let appliedWeek = '';
function applyIfLatest(result, latestWeekKey) {
  if (result.weekKey !== latestWeekKey) return false;
  appliedWeek = result.weekKey;
  return true;
}
assert.equal(applyIfLatest({ weekKey: '2026-07-06' }, '2026-07-13'), false, '9 stale week cannot apply');
assert.equal(applyIfLatest({ weekKey: '2026-07-13' }, '2026-07-13'), true, '9 latest week can apply');
assert.equal(appliedWeek, '2026-07-13', '9 latest response wins');

const empty = await loadWeeklySimulation({
  weekStart: baseMonday,
  optimized: async () => [],
  compatible: async () => [],
  timesheets: async () => [],
});
assert.equal(empty.punches.length + empty.timesheets.length, 0, '10 empty week stays empty');

const store = createPunchStore();
const punch = { companyId: 'chadwell', siteId: 'OH01', employeeId: 'emp_1', action: 'clock_in', timestampMs: baseMonday + HOUR };
assert.equal(store.save(punch).status, 'created', '11 first clock in creates');
assert.equal(store.save(punch).status, 'duplicate', '11 double-click creates one');
assert.equal(store.count(), 1, '11 one punch stored');

const simultaneous = createPunchStore();
assert.deepEqual([simultaneous.save(punch).status, simultaneous.save(punch).status], ['created', 'duplicate'], '12 simultaneous clock in creates one');
assert.equal(simultaneous.count(), 1, '12 one simultaneous punch stored');

const retry = createPunchStore();
retry.save(punch);
assert.equal(retry.save(punch).status, 'duplicate', '13 retry returns original result');

const tabs = createPunchStore();
assert.deepEqual([tabs.save(punch).status, tabs.save({ ...punch }).status], ['created', 'duplicate'], '14 two tabs converge on one punch');
assert.equal(tabs.save({ ...punch, timestampMs: punch.timestampMs + 30_000 }).status, 'duplicate', '15 same action within 60 seconds rejected');
assert.equal(tabs.save({ ...punch, action: 'start_lunch', timestampMs: punch.timestampMs + 5 * HOUR }).status, 'created', '16 different action allowed');

const sameNameEmployees = [
  { id: 'emp_a', employeeId: 'emp_a', employeeNumber: '1', name: 'Alex Lee' },
  { id: 'emp_b', employeeId: 'emp_b', employeeNumber: '2', name: 'Alex Lee' },
];
assert.equal(resolveEmployee({ typedName: 'Alex Lee', employees: sameNameEmployees }), null, '17 same names do not merge without identity');
assert.equal(resolveEmployee({ typedName: ' alex   lee ', employees: [sameNameEmployees[0]] })?.id, 'emp_a', '18 formatting does not create another employee when unique');

const duplicateDisplay = dedupePunches([
  { id: 'p1', employeeId: 'emp_1', action: 'clock_in', timestampMs: baseMonday + HOUR, companyId: 'chadwell', siteId: 'OH01' },
  { id: 'p2', employeeId: 'emp_1', employeeNumber: '100', name: 'Jordan Smith', action: 'clock_in', timestampMs: baseMonday + HOUR + 20_000, companyId: 'chadwell', siteId: 'OH01' },
]);
assert.equal(duplicateDisplay.deduped.length, 1, '19 historical duplicates display once');
assert.equal(duplicateDisplay.groups.length, 1, '19 duplicate group remains diagnosable');

const manual = normalizePunch({ id: 'manual_1', source: 'manual_manager', action: 'clock_out', timestampMs: baseMonday + 9 * HOUR });
assert.equal(manual.source, 'manual_manager', '20 manual manager punches remain marked manual');

assert.equal(branchMatch({ siteId: 'OHC' }, 'OH01'), false, 'different branch is excluded');
assert.equal(branchMatch({ assignedSiteId: 'OH01' }, 'OH01'), true, 'assignedSiteId compatibility branch is included');
assert.equal(branchMatch({ timestampMs: baseMonday }, 'OH01'), true, 'branchless legacy record can use selected branch fallback');

console.log('payroll critical regression passed');
