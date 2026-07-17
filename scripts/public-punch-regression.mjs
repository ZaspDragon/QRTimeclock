import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';

const ACTIONS = ['clock_in', 'start_lunch', 'end_lunch', 'clock_out'];
const DUPLICATE_WINDOW_MS = 60 * 1000;

function activeEmployee(row) {
  const siteOk = ['OH01', 'OHC'].some((site) =>
    row.siteId === site
    || row.assignedSiteId === site
    || (Array.isArray(row.siteIds) && row.siteIds.includes(site))
  );
  return row.companyId === 'chadwell'
    && siteOk
    && row.active !== false
    && !['inactive', 'removed', 'terminated', 'disabled', 'archived', 'merged'].includes(String(row.status || '').toLowerCase())
    && (row.status === 'active' || row.active === true);
}

function hash(value) {
  return createHash('sha256').update(value).digest('hex').slice(0, 40);
}

function guardKey(payload) {
  const window = Math.floor(payload.timestampMs / DUPLICATE_WINDOW_MS);
  return `public_${hash([payload.companyId, payload.siteId, payload.agencyId || '', payload.employeeId, payload.action, window].join('|'))}`;
}

function stateKey(payload) {
  return `public_${hash([payload.companyId, payload.siteId, payload.agencyId || '', payload.employeeId].join('|'))}`;
}

function validatePublicPunch(payload, employees) {
  assert.equal(payload.companyId, 'chadwell', 'companyId must be chadwell');
  assert(['OH01', 'OHC'].includes(payload.siteId), 'siteId must be OH01 or OHC');
  assert.equal(payload.source, 'public_qr', 'source must be public_qr');
  assert(ACTIONS.includes(payload.action), 'action must be valid');
  assert(activeEmployee(employees.get(payload.employeeId) || {}), 'employee must exist and be active');
  assert.equal(payload.duplicateGuardKey, guardKey(payload), 'punch id and duplicateGuardKey match app key');
  assert.equal(payload.workerStateKey, stateKey(payload), 'workerStateKey matches app key');
}

function createStore(employees) {
  const punches = new Map();
  const guards = new Map();
  const states = new Map();
  return {
    save(payload) {
      payload.duplicateGuardKey = guardKey(payload);
      payload.idempotencyKey = payload.duplicateGuardKey;
      payload.workerStateKey = stateKey(payload);
      validatePublicPunch(payload, employees);
      if (guards.has(payload.duplicateGuardKey)) return { status: 'duplicate' };
      const state = states.get(payload.workerStateKey);
      if (state?.lastAction === payload.action) return { status: 'invalid_repeated_action' };
      punches.set(payload.duplicateGuardKey, payload);
      guards.set(payload.duplicateGuardKey, { punchId: payload.duplicateGuardKey, duplicateGuardKey: payload.duplicateGuardKey });
      states.set(payload.workerStateKey, {
        workerStateKey: payload.workerStateKey,
        employeeId: payload.employeeId,
        lastAction: payload.action,
        lastPunchAtMs: payload.timestampMs,
        lastPunchId: payload.duplicateGuardKey,
      });
      return { status: 'created' };
    },
    counts: () => ({ punches: punches.size, guards: guards.size, states: states.size }),
  };
}

function payload(employeeId, action, siteId = 'OH01', timestampMs = Date.parse('2026-07-17T08:00:00-04:00')) {
  return {
    companyId: 'chadwell',
    siteId,
    source: 'public_qr',
    name: 'Jordan Smith',
    nameKey: 'jordan_smith',
    action,
    timestampMs,
    dateKey: '2026-07-17',
    weekKey: '2026-07-13',
    employeeId,
    employeeNumber: '100',
    agencyId: '',
    assignedSiteId: siteId,
    siteIds: [siteId],
    qrSlug: siteId.toLowerCase(),
  };
}

const employees = new Map([
  ['emp_oh01', { companyId: 'chadwell', siteId: 'OH01', status: 'active' }],
  ['emp_ohc', { companyId: 'chadwell', assignedSiteId: 'OHC', active: true }],
  ['emp_sites', { companyId: 'chadwell', siteIds: ['OH01'], active: true }],
  ['emp_inactive', { companyId: 'chadwell', siteId: 'OH01', status: 'inactive', active: false }],
  ['emp_stale_active', { companyId: 'chadwell', siteId: 'OH01', status: 'inactive', active: true }],
]);

assert(activeEmployee(employees.get('emp_oh01')), 'status active employee is active');
assert(activeEmployee(employees.get('emp_ohc')), 'active true assignedSiteId employee is active');
assert(activeEmployee(employees.get('emp_sites')), 'active true siteIds employee is active');

let store = createStore(employees);
assert.equal(store.save(payload('emp_oh01', 'clock_in')).status, 'created', 'anonymous OH01 employee can clock in');
assert.deepEqual(store.counts(), { punches: 1, guards: 1, states: 1 }, 'matching punch, guard, and state created atomically');
assert.equal(store.save(payload('emp_oh01', 'clock_in')).status, 'duplicate', 'duplicate request does not create a second punch');
assert.deepEqual(store.counts(), { punches: 1, guards: 1, states: 1 }, 'duplicate count remains one');
assert.equal(store.save(payload('emp_oh01', 'start_lunch', 'OH01', Date.parse('2026-07-17T12:00:00-04:00'))).status, 'created', 'start lunch after clock in updates existing state');
assert.equal(store.save(payload('emp_oh01', 'end_lunch', 'OH01', Date.parse('2026-07-17T12:30:00-04:00'))).status, 'created', 'end lunch after start lunch');
assert.equal(store.save(payload('emp_oh01', 'clock_out', 'OH01', Date.parse('2026-07-17T17:00:00-04:00'))).status, 'created', 'clock out after end lunch');

store = createStore(employees);
assert.equal(store.save(payload('emp_ohc', 'clock_in', 'OHC')).status, 'created', 'anonymous OHC employee can clock in');

assert.throws(() => createStore(employees).save(payload('emp_inactive', 'clock_in')), /active/, 'inactive employee punch fails');
assert.throws(() => createStore(employees).save(payload('emp_stale_active', 'clock_in')), /active/, 'inactive status with stale active true fails');
assert.throws(() => createStore(employees).save(payload('missing', 'clock_in')), /active/, 'unknown employee punch fails');
assert.throws(() => createStore(employees).save({ ...payload('emp_oh01', 'clock_in'), companyId: 'other' }), /companyId/, 'invalid company fails');
assert.throws(() => createStore(employees).save(payload('emp_oh01', 'clock_in', 'BAD')), /siteId/, 'invalid site fails');
assert.throws(() => createStore(employees).save(payload('emp_oh01', 'bad_action')), /action/, 'invalid action fails');

const rules = readFileSync(new URL('../firestore.rules', import.meta.url), 'utf8');
assert(!/allow\s+read\s*,\s*write\s*:\s*if\s+true/.test(rules), 'rules do not use allow read, write: if true');
assert(/match \/punchGuards\/\{guardId\}[\s\S]*allow list: if false/.test(rules), 'public cannot list punchGuards');
assert(/match \/punchStates\/\{stateId\}[\s\S]*allow list: if false/.test(rules), 'public cannot list punchStates');
assert(/allow update:\s*if publicPunchStateUpdate/.test(rules), 'existing punchState can be updated by public transaction rules');
assert(/allow delete:\s*if false/.test(rules), 'public deletes remain denied');

console.log('public punch regression passed');
