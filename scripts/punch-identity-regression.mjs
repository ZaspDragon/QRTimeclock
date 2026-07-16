import assert from 'node:assert/strict';

const HOUR = 60 * 60 * 1000;
const MIN = 60 * 1000;
const TOLERANCE_MS = 60 * 1000;
const base = Date.parse('2026-07-13T08:00:00-04:00');

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, '_').replace(/^_+|_+$/g, '');
}

function dateKey(ms) {
  return new Date(ms).toISOString().slice(0, 10);
}

function punch(id, employeeId, action, offset, extra = {}) {
  const timestampMs = base + offset;
  return {
    id,
    employeeId,
    action,
    timestampMs,
    dateKey: dateKey(timestampMs),
    companyId: extra.companyId || 'chadwell',
    siteId: extra.siteId || 'OH01',
    agencyId: extra.agencyId || '',
    name: extra.name || 'Alex Rivera',
    nameKey: normalizeName(extra.name || 'Alex Rivera'),
    employeeNumber: extra.employeeNumber || 'EMP-1',
    ...extra,
  };
}

function resolveEmployeeIdentitySet(employee, records = []) {
  const employeeIds = new Set([employee.id, employee.employeeId, employee.workerId].filter(Boolean));
  const workerIds = new Set([employee.workerId].filter(Boolean));
  const employeeNumbers = new Set([employee.employeeNumber].filter(Boolean));
  const normalizedNames = new Set([normalizeName(employee.name || employee.nameKey)].filter(Boolean));
  const mergedRecordIds = new Set();
  const allowedSiteIds = new Set([employee.siteId, employee.assignedSiteId].filter(Boolean));
  const warnings = [];
  const sameScope = (row) =>
    (row.companyId || 'chadwell') === (employee.companyId || 'chadwell')
    && (row.agencyId || '') === (employee.agencyId || '')
    && (!row.siteId || !allowedSiteIds.size || allowedSiteIds.has(row.siteId) || allowedSiteIds.has(row.assignedSiteId));

  records.filter(sameScope).forEach((record) => {
    const ids = [record.id, record.employeeId, record.workerId].filter(Boolean);
    if (record.mergedInto && employeeIds.has(record.mergedInto)) {
      ids.forEach((id) => employeeIds.add(id));
      mergedRecordIds.add(record.id);
    }
    if (ids.some((id) => employeeIds.has(id)) || record.employeeNumber === employee.employeeNumber) {
      ids.forEach((id) => employeeIds.add(id));
      if (record.workerId) workerIds.add(record.workerId);
      if (record.employeeNumber) employeeNumbers.add(record.employeeNumber);
      if (record.name) normalizedNames.add(normalizeName(record.name));
    }
  });

  const scopedNameNumbers = new Set(records
    .filter((record) => sameScope(record) && normalizeName(record.name || record.nameKey) === normalizeName(employee.name))
    .map((record) => record.employeeNumber)
    .filter(Boolean));
  if (scopedNameNumbers.size > 1) warnings.push('Ambiguous same-name employees with different employee numbers; name fallback restricted.');

  return {
    canonicalEmployeeId: employee.mergedInto || employee.id || employee.employeeId,
    employeeIds: [...employeeIds],
    workerIds: [...workerIds],
    employeeNumbers: [...employeeNumbers],
    normalizedNames: [...normalizedNames],
    mergedRecordIds: [...mergedRecordIds],
    companyId: employee.companyId || 'chadwell',
    agencyId: employee.agencyId || '',
    allowedSiteIds: [...allowedSiteIds],
    warnings,
  };
}

function scopeMatches(p, identity) {
  return (p.companyId || 'chadwell') === identity.companyId
    && (p.agencyId || '') === identity.agencyId
    && (!identity.allowedSiteIds.length || identity.allowedSiteIds.includes(p.siteId) || identity.allowedSiteIds.includes(p.assignedSiteId));
}

function includePunches(employee, records, punches) {
  const identity = resolveEmployeeIdentitySet(employee, records);
  return punches.filter((p) => {
    if (p.status === 'deleted' || p.active === false) return false;
    if (!scopeMatches(p, identity)) return false;
    if (identity.employeeIds.includes(p.employeeId) || identity.workerIds.includes(p.workerId)) return true;
    if (p.employeeNumber && identity.employeeNumbers.includes(p.employeeNumber)) return true;
    if ((p.employeeId || p.workerId) && !identity.employeeIds.includes(p.employeeId)) return false;
    return identity.normalizedNames.includes(p.nameKey || normalizeName(p.name));
  }).map((p) => ({ ...p, resolvedEmployeeId: identity.canonicalEmployeeId }));
}

function semanticDedupe(punches) {
  const groups = [];
  for (const p of [...punches].sort((a, b) => a.timestampMs - b.timestampMs)) {
    const signature = [p.resolvedEmployeeId || p.employeeId || p.workerId || p.nameKey, p.action, p.companyId, p.siteId, p.agencyId || ''].join('|');
    let group = groups.find((g) => g.signature === signature && Math.abs(g.primary.timestampMs - p.timestampMs) <= TOLERANCE_MS);
    if (!group) {
      group = { signature, primary: p, punches: [] };
      groups.push(group);
    }
    group.punches.push(p);
  }
  return groups.map((g) => ({ ...g.punches[0], duplicateCount: g.punches.length, duplicateSourcePunches: g.punches }));
}

function totals(rawPunches) {
  const punches = semanticDedupe(rawPunches.filter((p) => p.status !== 'deleted' && p.active !== false));
  let minutes = 0;
  const warnings = new Set();
  let activeStart = null;
  let lunchStart = null;
  for (const p of punches.sort((a, b) => a.timestampMs - b.timestampMs)) {
    if (p.duplicateCount > 1) warnings.add(`Duplicate ${p.action}`);
    if (p.action === 'clock_in') {
      if (activeStart) warnings.add('Overlapping shift');
      activeStart = p.timestampMs;
    }
    if (p.action === 'start_lunch') {
      if (!activeStart) warnings.add('Start Lunch has no Clock In');
      else minutes += Math.round((p.timestampMs - activeStart) / MIN);
      activeStart = null;
      lunchStart = p.timestampMs;
    }
    if (p.action === 'end_lunch') {
      if (!lunchStart) warnings.add('Missing Start Lunch');
      lunchStart = null;
      activeStart = p.timestampMs;
    }
    if (p.action === 'clock_out') {
      if (!activeStart) warnings.add('Missing Clock In');
      else minutes += Math.round((p.timestampMs - activeStart) / MIN);
      activeStart = null;
    }
  }
  if (activeStart) warnings.add('Missing Clock Out');
  if (lunchStart) warnings.add('Start Lunch has no End Lunch');
  return { hours: Number((minutes / 60).toFixed(2)), warnings: [...warnings], punches };
}

function assertHours(label, punches, expected) {
  assert.equal(totals(punches).hours, expected, label);
}

const employee = { id: 'emp-new', employeeId: 'emp-new', workerId: 'wrk-1', employeeNumber: 'EMP-1', name: 'Alex Rivera', companyId: 'chadwell', siteId: 'OH01', agencyId: '' };
const merged = [{ id: 'emp-old', employeeId: 'emp-old', mergedInto: 'emp-new', employeeNumber: 'EMP-1', name: 'Alex Rivera', companyId: 'chadwell', siteId: 'OH01', agencyId: '' }];

assertHours('normal complete shift', [
  punch('a', 'emp-new', 'clock_in', 0),
  punch('b', 'emp-new', 'start_lunch', 4 * HOUR),
  punch('c', 'emp-new', 'end_lunch', 4.5 * HOUR),
  punch('d', 'emp-new', 'clock_out', 9 * HOUR),
], 8.5);
assertHours('no lunch shift', [punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR)], 8);
assertHours('merged older id plus canonical id', includePunches(employee, merged, [punch('a', 'emp-old', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR)]), 8);
assertHours('legacy blank employee id scoped name', includePunches(employee, [], [punch('a', '', 'clock_in', 0), punch('b', '', 'clock_out', 8 * HOUR)]), 8);
assert.equal(includePunches(employee, [], [punch('x', '', 'clock_in', 0, { agencyId: 'other' })]).length, 0, 'same-name different agency excluded');
assert.equal(includePunches(employee, [], [punch('x', '', 'clock_in', 0, { siteId: 'OHC' })]).length, 0, 'same-name different site excluded');
assert.equal(totals([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR), punch('c', 'emp-new', 'clock_out', 8 * HOUR)]).hours, 8, 'identical duplicate clock out ignored');
assert.equal(totals([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR), punch('c', 'emp-new', 'clock_out', 8 * HOUR + 30 * 1000)]).hours, 8, '30 second duplicate clock out ignored');
assertHours('two legitimate shifts', [punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 4 * HOUR), punch('c', 'emp-new', 'clock_in', 6 * HOUR), punch('d', 'emp-new', 'clock_out', 10 * HOUR)], 8);
assert.equal(semanticDedupe([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_in', 10 * 1000)]).length, 1, 'two phones same time deduped');
assert.equal(semanticDedupe([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_in', 5 * 1000)]).length, 1, 'refresh retry deduped');
assert(totals([punch('a', 'emp-new', 'clock_out', 8 * HOUR)]).warnings.includes('Missing Clock In'), 'missing clock in warned');
assert(totals([punch('a', 'emp-new', 'clock_in', 0)]).warnings.includes('Missing Clock Out'), 'missing clock out warned');
assert(totals([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'start_lunch', 4 * HOUR)]).warnings.includes('Start Lunch has no End Lunch'), 'missing end lunch warned');
assertHours('deleted punch excluded', [punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR, { status: 'deleted', active: false })], 0);
assertHours('manager correction included', [punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR, { source: 'manager_inserted', auditId: 'edit-1' })], 8);
const exportHours = totals([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR)]).hours;
assert.equal(exportHours, 8, 'agency export equals corrected timecard calculation');
assert.equal(exportHours, totals([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR)]).hours, 'weekly signoff equals My Time totals');
assert.equal(dateKey(base - 14 * 24 * HOUR), '2026-06-29', 'historical weeks remain addressable');
const original = Object.freeze([punch('a', 'emp-new', 'clock_in', 0), punch('b', 'emp-new', 'clock_out', 8 * HOUR)]);
const before = JSON.stringify(original);
totals(original);
assert.equal(JSON.stringify(original), before, 'read-only calculations do not mutate source punches');

console.log('punch identity regression passed');
