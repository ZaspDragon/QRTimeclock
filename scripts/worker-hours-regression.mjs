import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import vm from 'node:vm';

// Execute the production lookup module with a small DOM/fetch adapter. There
// are deliberately no Firestore or storage write capabilities in this context.
const read = (path) => readFileSync(new URL(`../${path}`, import.meta.url), 'utf8');
const nodes = new Map();
class Element {
  constructor(id = '') {
    this.id = id;
    this.value = '';
    this.textContent = '';
    this.disabled = false;
    this.children = [];
    this.dataset = {};
    this.classList = { add() {}, remove() {}, toggle() {} };
  }
  append(...items) { this.children.push(...items); }
  replaceChildren(...items) { this.children = items; }
  closest(selector) {
    return selector.split(',').some((part) => part.trim() === `#${this.id}`
      || (part.trim() === '.worker-action-btn' && this.dataset.action)
      || (part.trim() === '.worker-range-quick' && this.dataset.range)) ? this : null;
  }
}
const el = (id) => {
  if (!nodes.has(id)) nodes.set(id, new Element(id));
  return nodes.get(id);
};
const listeners = new Map();
const timers = new Map();
let nextTimer = 0;
let fetchImpl;
const requests = [];
const context = vm.createContext({
  console, Element, URLSearchParams, AbortController, TypeError,
  location: { search: '' },
  document: {
    getElementById: el,
    createElement: () => new Element(),
    createTextNode: (text) => text,
    addEventListener(type, handler) {
      if (!listeners.has(type)) listeners.set(type, []);
      listeners.get(type).push(handler);
    },
  },
  window: {
    setTimeout(fn) { timers.set(++nextTimer, fn); return nextTimer; },
    clearTimeout(id) { timers.delete(id); },
  },
  fetch: async (url, options) => {
    requests.push({ url, ...options });
    return fetchImpl(url, options);
  },
});
vm.runInContext(read('name-only-time-lookup.js'), context);
// Run the unchanged production punch dispatcher beside the lookup listener.
vm.runInContext(read('punch-writer-lock.js').replaceAll('export ', ''), context);
vm.runInContext("globalThis.punchActions = []; registerPunchWriter('test-writer', 100, async action => punchActions.push(action));", context);
const evaluate = (code) => vm.runInContext(code, context);
async function dispatch(type, target) {
  const event = {
    target, stopped: false, prevented: false,
    preventDefault() { this.prevented = true; },
    stopPropagation() {},
    stopImmediatePropagation() { this.stopped = true; },
  };
  for (const listener of listeners.get(type) || []) {
    await listener(event);
    if (event.stopped) break;
  }
  return event;
}
el('workerNameInput').value = 'Test Worker';
el('workerBranchSelect').value = 'OH01';
el('workerAgencySelect').value = 'sterling_staffing';
el('workerLookupStatus').textContent = 'Ready to punch';
el('workerStatusMessage').textContent = 'Clock ready';
const good = (punches = []) => ({ ok: true, status: 200, json: async () => ({ worker: { name: 'Test Worker' }, punches }) });

fetchImpl = async () => good();
await evaluate("handleLookup('week')");
assert.equal(el('workerWeekHoursValue').textContent, '0.00', 'successful empty history really is zero');
assert.equal(requests.at(-1).method, 'POST');
assert.equal(requests.at(-1).cache, 'no-store');
assert.equal(JSON.parse(requests.at(-1).body).agencyId, 'sterling_staffing');

for (const response of [
  { ok: false, status: 404, json: async () => { throw new Error('HTML 404'); } },
  { ok: false, status: 409, json: async () => ({ error: 'Duplicate profiles need review.' }) },
  { ok: true, status: 200, json: async () => ({}) },
]) {
  fetchImpl = async () => response;
  await evaluate("handleLookup('week')");
  assert.equal(el('workerWeekHoursValue').textContent, '—', 'failure must not display zero hours');
  assert(!el('workerTimeRangeStatus').textContent.includes('Total Hours: 0'));
  assert.equal(el('workerViewTimeBtn').disabled, false);
}
fetchImpl = async () => { throw new TypeError('Failed to fetch'); };
await evaluate("handleLookup('week')");
assert.match(el('workerTimeRangeStatus').textContent, /Could not connect/);

// A slow or failed hours service must never disable or intercept punch buttons.
fetchImpl = async (_url, { signal }) => new Promise((_resolve, reject) => {
  signal.addEventListener('abort', () => reject(Object.assign(new Error('aborted'), { name: 'AbortError' })));
});
const slowLookup = evaluate("handleLookup('week')");
assert.equal(el('workerViewTimeBtn').disabled, true);
for (const action of ['clock_in', 'start_lunch', 'end_lunch', 'clock_out']) {
  const button = el(action);
  button.dataset.action = action;
  assert.equal(button.disabled, false);
  await dispatch('click', button);
}
assert.equal(JSON.stringify(context.punchActions), JSON.stringify(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']));
for (const fn of [...timers.values()]) fn();
await slowLookup;
assert.match(el('workerTimeRangeStatus').textContent, /timed out/);
assert.equal(el('workerViewTimeBtn').disabled, false);
assert.equal(el('workerLookupStatus').textContent, 'Ready to punch');
assert.equal(el('workerStatusMessage').textContent, 'Clock ready');

// Ignore an old response after the user changes worker or branch.
let resolveOld;
fetchImpl = async () => new Promise((resolve) => { resolveOld = resolve; });
const oldLookup = evaluate("handleLookup('week')");
el('workerNameInput').value = 'Next Worker';
await dispatch('input', el('workerNameInput'));
resolveOld(good());
await oldLookup;
assert.equal(el('workerWeekHoursValue').textContent, '—');
assert.match(el('workerTimeRangeStatus').textContent, /refresh/);

fetchImpl = async () => good();
el('workerTimeFromInput').value = '2026-08-03';
el('workerTimeToInput').value = '2026-08-09';
const before = requests.length;
await dispatch('click', el('workerTimeLookupBtn'));
await new Promise((resolve) => setImmediate(resolve));
assert.equal(requests.length, before + 1, 'past-time button requests exactly once');
assert.equal(JSON.parse(requests.at(-1).body).name, 'Next Worker', 'click event never becomes the employee');
const quickRange = evaluate("applyQuickRange('last_2_weeks')");
assert.equal(Math.round((quickRange.toMs - quickRange.fromMs + 1) / 86400000), 14);
const requestCount = requests.length;
await evaluate("handleLookup('custom', { fromMs: 0, toMs: 40 * 86400000 })");
assert.equal(requests.length, requestCount, 'invalid range never reaches the server');
assert.match(el('workerTimeRangeStatus').textContent, /31 days/);

// Use actual summary code: two 35-hour weeks must have zero overtime.
context.fixture = ['2026-08-03', '2026-08-10'].flatMap((monday) => Array.from({ length: 5 }, (_, day) => {
  const start = Date.parse(`${monday}T08:00:00Z`) + day * 86400000;
  return [
    { action: 'clock_in', timestampMs: start, dateKey: new Date(start).toISOString().slice(0, 10) },
    { action: 'clock_out', timestampMs: start + 7 * 3600000, dateKey: new Date(start).toISOString().slice(0, 10) },
  ];
})).flat();
const summary = evaluate('summarizePunches(fixture)');
assert.equal(summary.totalMinutes, 70 * 60);
assert.equal(summary.regularMinutes, 70 * 60);
assert.equal(summary.overtimeMinutes, 0);
assert(!/localStorage|sessionStorage/.test(read('name-only-time-lookup.js')));
console.log('worker hours regression passed: success, failures, timeout, punch isolation, stale responses, past dates, weekly totals');

// Execute the real HTTP function against a read-only Firestore adapter.
let data = { employees: [], punches: [] };
let queryFailure = false;
function collection(name) {
  function makeQuery(field, value, pageSize = 400, cursor = null) {
    return {
      limit(size) { return makeQuery(field, value, size, cursor); },
      startAfter(doc) { return makeQuery(field, value, pageSize, doc.id); },
      async get() {
        if (queryFailure) throw new Error('database unavailable');
        const all = data[name].filter((row) => row[field] === value).sort((a, b) => a.id.localeCompare(b.id));
        const start = cursor ? all.findIndex((row) => row.id === cursor) + 1 : 0;
        return { docs: all.slice(start, start + pageSize).map((row) => ({ id: row.id, data: () => ({ ...row }) })) };
      },
    };
  }
  return {
    where(field, op, value) { assert.equal(op, '=='); return makeQuery(field, value); },
    doc(id) { return { get: async () => ({ exists: data[name].some((row) => row.id === id), id, data: () => data[name].find((row) => row.id === id) }) }; },
  };
}
const backend = vm.createContext({
  console: { error() {}, warn() {} }, exports: {},
  require(name) {
    if (name === 'firebase-functions/v2/https') return { onRequest: (_options, handler) => handler };
    if (name === 'firebase-admin/app') return { initializeApp() {} };
    if (name === 'firebase-admin/firestore') return { getFirestore: () => ({ collection }) };
    throw new Error(`Unexpected dependency ${name}`);
  },
});
vm.runInContext(read('functions/index.js'), backend);
const worker = { id: 'worker1', name: 'Test Worker', nameKey: 'test_worker', active: true, companyId: 'chadwell', siteId: 'OH01', agencyId: 'sterling_staffing', linkedWorkerIds: ['legacy1'] };
data.employees = [worker];
const fromMs = Date.parse('2026-08-03T00:00:00-04:00');
const toMs = Date.parse('2026-08-09T23:59:59.999-04:00');
const punch = (id, extra = {}) => ({ id, employeeId: 'worker1', companyId: 'chadwell', siteId: 'OH01', agencyId: 'sterling_staffing', action: 'clock_in', timestampMs: fromMs + 8 * 3600000, ...extra });
data.punches = [
  ...Array.from({ length: 405 }, (_, n) => punch(`a${String(n).padStart(4, '0')}`, { timestampMs: fromMs - 86400000 })),
  punch('z-current'),
  punch('z-legacy', { employeeId: 'legacy1', action: 'clock_out', timestampMs: fromMs + 16 * 3600000 }),
  punch('z-wrong-agency', { agencyId: 'excel_staffing' }),
  punch('z-wrong-site', { siteId: 'OHC' }),
  punch('z-deleted', { status: 'deleted' }),
  punch('z-inactive', { active: false }),
];
let ip = 0;
async function call(extra = {}) {
  const res = { code: 200, set() {}, status(code) { this.code = code; return this; }, json(payload) { this.payload = payload; return this; }, send(value) { this.payload = value; return this; } };
  await backend.exports.publicWorkerTimeLookup({ method: 'POST', headers: { origin: 'https://zaspdragon.github.io' }, ip: `test-${++ip}`, body: { name: worker.name, siteId: 'OH01', agencyId: 'sterling_staffing', fromMs, toMs, ...extra } }, res);
  return res;
}
let result = await call();
assert.equal(result.code, 200);
assert.equal(result.payload.punches.length, 2, 'find recent and linked history beyond first 400 records; exclude other scopes/deleted');
assert.equal(Object.keys(result.payload.punches[0]).sort().join(','), 'action,dateKey,timestampMs');
data.punches[405].timestampMs += 3600000;
result = await call();
assert.equal(result.payload.punches[0].timestampMs, fromMs + 9 * 3600000, 'fresh lookup reflects correction');
queryFailure = true;
result = await call();
assert.equal(result.code, 500, 'database failures are not successful empty history');
queryFailure = false;
data.punches = Array.from({ length: 401 }, (_, n) => punch(`p${n}`));
assert.equal((await call()).code, 422, 'too many results never silently truncate totals');
data.punches = [];
data.employees.push({ ...worker, id: 'unrelated', linkedWorkerIds: [] });
assert.equal((await call()).code, 409, 'ambiguous identities are not combined');
assert.equal((await call({ agencyId: 'invalid' })).code, 400);
assert.equal((await call({ toMs: fromMs + 40 * 86400000 })).code, 400);
console.log('worker hours backend regression passed: pagination, linked IDs, scope, corrections, errors, output limits, ambiguity');
