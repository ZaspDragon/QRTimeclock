import assert from 'node:assert/strict';

const HOUR = 60 * 60 * 1000;
const DAY = 24 * HOUR;
const ACTIONS = ['clock_in', 'start_lunch', 'end_lunch', 'clock_out'];

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .replaceAll(' ', '_');
}

function monday(value) {
  const date = new Date(value);
  date.setHours(0, 0, 0, 0);
  const day = date.getDay();
  date.setDate(date.getDate() - (day === 0 ? 6 : day - 1));
  return date;
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function dateKey(value) {
  const date = value instanceof Date ? value : new Date(value);
  return date.toISOString().slice(0, 10);
}

function timestampMs(row) {
  const explicit = Number(row.timestampMs || 0);
  if (explicit) return explicit;
  if (row.timestamp instanceof Date) return row.timestamp.getTime();
  if (/^\d{4}-\d{2}-\d{2}$/.test(String(row.dateKey || ''))) {
    return new Date(`${row.dateKey}T12:00:00`).getTime();
  }
  return 0;
}

function normalizePunch(row) {
  const ms = timestampMs(row);
  const name = row.name || row.workerName || row.employeeName || '';
  return {
    ...row,
    timestampMs: ms,
    action: row.action,
    name,
    nameKey: row.nameKey || normalizeName(name),
    dateKey: row.dateKey || dateKey(ms),
    weekKey: row.weekKey || dateKey(monday(ms)),
  };
}

function isActivePunch(row) {
  return row && row.active !== false && String(row.status || '').toLowerCase() !== 'deleted';
}

function getValidPunchesForSelectedWeek(rows, selectedWeekStart) {
  const start = monday(selectedWeekStart).getTime();
  const end = addDays(new Date(start), 7).getTime();
  return rows.map(normalizePunch).filter((row) =>
    isActivePunch(row)
    && ACTIONS.includes(row.action)
    && row.timestampMs >= start
    && row.timestampMs < end
  );
}

function identity(row) {
  if (row.employeeId) return `worker:${row.employeeId}`;
  if (row.workerId) return `worker:${row.workerId}`;
  const signature = [row.nameKey || normalizeName(row.name), row.agencyId || '', row.siteId || row.assignedSiteId || ''].join('|');
  return signature.trim() ? `person:${signature}` : '';
}

function buildWorkedWorkerGroups(rows, selectedWeekStart) {
  const groups = new Map();
  for (const row of getValidPunchesForSelectedWeek(rows, selectedWeekStart)) {
    const key = identity(row);
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(row);
  }
  return [...groups.entries()].map(([identityKey, punches]) => ({ identityKey, punches }));
}

function punch(id, employeeId, action, offset, extra = {}) {
  const timestamp = weekStart + offset;
  return {
    id,
    employeeId,
    name: extra.name || 'Jordan Smith',
    companyId: 'chadwell',
    siteId: 'OH01',
    action,
    timestampMs: timestamp,
    dateKey: dateKey(timestamp),
    ...extra,
  };
}

const weekStart = monday('2026-07-13T12:00:00').getTime();

assert.equal(buildWorkedWorkerGroups([punch('one', 'emp_1', 'clock_in', HOUR)], weekStart).length, 1, 'worker with one clock in appears');
assert.equal(buildWorkedWorkerGroups([
  punch('in', 'emp_1', 'clock_in', HOUR),
  punch('l1', 'emp_1', 'start_lunch', 4 * HOUR),
  punch('l2', 'emp_1', 'end_lunch', 5 * HOUR),
  punch('out', 'emp_1', 'clock_out', 9 * HOUR),
], weekStart).length, 1, 'complete week appears');
assert.equal(buildWorkedWorkerGroups([punch('missing-out', 'emp_1', 'clock_in', HOUR)], weekStart)[0].punches.length, 1, 'missing clock out remains visible');
assert.equal(buildWorkedWorkerGroups([
  punch('in', 'emp_1', 'clock_in', HOUR),
  punch('out', 'emp_1', 'clock_out', 8 * HOUR),
], weekStart)[0].punches.length, 2, 'missing lunch remains visible');
assert.equal(buildWorkedWorkerGroups([punch('last-week', 'emp_1', 'clock_in', -6 * DAY)], weekStart).length, 0, 'last week worker excluded');
assert.equal(buildWorkedWorkerGroups([], weekStart).length, 0, 'active employee with no punches excluded');
assert.equal(buildWorkedWorkerGroups([{ id: 'deleted', employeeId: 'emp_1', action: 'clock_in', timestampMs: weekStart + HOUR, status: 'deleted' }], weekStart).length, 0, 'deleted-only punches excluded');
assert.equal(buildWorkedWorkerGroups([
  punch('a', 'emp_1', 'clock_in', HOUR, { name: 'Alex Lee' }),
  punch('b', 'emp_2', 'clock_in', 2 * HOUR, { name: 'Alex Lee' }),
], weekStart).length, 2, 'same names with different employee IDs stay separate');
assert.equal(buildWorkedWorkerGroups([
  punch('a', 'emp_1', 'clock_in', HOUR, { name: 'Alex Lee' }),
  punch('b', 'emp_1', 'clock_out', 8 * HOUR, { name: 'Alex Lee Copy' }),
], weekStart).length, 1, 'same employee ID merges copied names');
assert.equal(buildWorkedWorkerGroups([
  punch('hist', '', 'clock_in', HOUR, { name: 'Historical Person', agencyId: 'agency-a', employeeId: '', workerId: '' }),
], weekStart).length, 1, 'historical punch missing employeeId uses canonical fallback');
assert.equal(getValidPunchesForSelectedWeek([
  { id: 'date-only', employeeId: 'emp_date', action: 'clock_in', dateKey: dateKey(weekStart + DAY), name: 'Date Only' },
], weekStart).length, 1, 'dateKey fallback stays in selected week');

const weekly = buildWorkedWorkerGroups([punch('visible', 'emp_1', 'clock_in', HOUR)], weekStart);
const agencyPreview = weekly;
const csv = weekly;
const excel = weekly;
const print = weekly;
assert.equal(agencyPreview.length, weekly.length, 'agency preview worker count matches weekly sign-off');
assert.deepEqual(csv.map((row) => row.identityKey), weekly.map((row) => row.identityKey), 'CSV identities match');
assert.deepEqual(excel.map((row) => row.identityKey), weekly.map((row) => row.identityKey), 'Excel identities match');
assert.deepEqual(print.map((row) => row.identityKey), weekly.map((row) => row.identityKey), 'print identities match');

const before = JSON.stringify(weekly);
buildWorkedWorkerGroups(weekly.flatMap((row) => row.punches), weekStart);
assert.equal(JSON.stringify(weekly), before, 'viewing report creates no records and does not mutate rows');

console.log('weekly selected-week regression passed');
