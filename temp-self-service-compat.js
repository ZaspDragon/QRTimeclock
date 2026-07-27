// Temp worker self-service compatibility layer.
// Isolated from core punch, manager, export, and timesheet logic.
// Reads existing data only; the only write is a new pending missedPunchRequest.

import { firebaseConfig } from './firebase-config.js';
import { getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  getFirestore,
  collection,
  query,
  where,
  getDocs,
  addDoc,
  serverTimestamp,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps()[0] || initializeApp(firebaseConfig);
const db = getFirestore(app);
const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
let resolvedWorker = null;

const el = (id) => document.getElementById(id);
const normalizeName = (value) => String(value || '')
  .trim()
  .toLowerCase()
  .replace(/[^a-z0-9]+/g, ' ')
  .trim();
const selectedSite = () => VALID_SITES.has(el('workerBranchSelect')?.value)
  ? el('workerBranchSelect').value
  : 'OH01';
const isActive = (worker) => worker?.active === true || String(worker?.status || '').toLowerCase() === 'active';
const workerId = (worker) => String(worker?.employeeId || worker?.id || '').trim();
const timestampMs = (row) => Number(
  row?.timestampMs
  || row?.timestamp?.toMillis?.()
  || (row?.timestamp?.seconds ? row.timestamp.seconds * 1000 : 0)
  || 0
);
const actionLabel = (action) => ({
  clock_in: 'Clock In',
  start_lunch: 'Start Lunch',
  end_lunch: 'End Lunch',
  clock_out: 'Clock Out',
}[action] || action || '-');

function setStatus(message, error = false) {
  const status = el('workerLookupStatus');
  if (status) {
    status.textContent = message;
    status.style.borderColor = error ? 'rgba(255,90,90,.55)' : 'rgba(43,213,118,.4)';
  }
}

function setRangeStatus(message) {
  const status = el('workerTimeRangeStatus');
  if (status) status.textContent = message;
}

function uniqueById(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const key = row.id || `${row.employeeId || ''}|${row.workerId || ''}|${row.timestampMs || ''}|${row.action || ''}`;
    if (!map.has(key)) map.set(key, row);
  });
  return [...map.values()];
}

async function safeQuery(collectionName, constraints) {
  try {
    const snapshot = await getDocs(query(collection(db, collectionName), ...constraints));
    return snapshot.docs.map((docSnapshot) => ({ id: docSnapshot.id, ...docSnapshot.data() }));
  } catch (error) {
    console.warn(`[temp self-service] ${collectionName} compatibility query skipped`, error?.code || '', error?.message || error);
    return [];
  }
}

async function resolveWorkerByTypedName() {
  const typed = String(el('workerNameInput')?.value || '').trim();
  const nameKey = normalizeName(typed);
  const siteId = selectedSite();
  if (nameKey.length < 2) throw new Error('Enter your first and last name.');

  const queryResults = await Promise.all([
    safeQuery('employees', [
      where('companyId', '==', COMPANY_ID),
      where('siteId', '==', siteId),
      where('nameKey', '==', nameKey),
    ]),
    safeQuery('employees', [
      where('companyId', '==', COMPANY_ID),
      where('assignedSiteId', '==', siteId),
      where('nameKey', '==', nameKey),
    ]),
    safeQuery('employees', [
      where('companyId', '==', COMPANY_ID),
      where('nameKey', '==', nameKey),
    ]),
    safeQuery('employees', [where('nameKey', '==', nameKey)]),
  ]);

  const matches = uniqueById(queryResults.flat()).filter((worker) => {
    const workerName = normalizeName(worker.name || worker.nameKey || '');
    const companyMatches = !worker.companyId || worker.companyId === COMPANY_ID;
    const workerSites = [worker.siteId, worker.assignedSiteId, ...(Array.isArray(worker.siteIds) ? worker.siteIds : [])]
      .filter(Boolean);
    const siteMatches = workerSites.length === 0 || workerSites.includes(siteId);
    return workerName === nameKey && companyMatches && siteMatches;
  });

  if (!matches.length) throw new Error('No existing worker record was found for that exact name and branch.');

  const activeMatches = matches.filter(isActive);
  const candidates = activeMatches.length ? activeMatches : matches;
  if (candidates.length > 1) {
    const numbered = candidates.filter((worker) => String(worker.employeeNumber || '').trim());
    if (numbered.length === 1) {
      resolvedWorker = numbered[0];
    } else {
      throw new Error('More than one worker has that name. Ask a manager to merge or rename the duplicate profiles.');
    }
  } else {
    resolvedWorker = candidates[0];
  }

  setStatus(`Found ${resolvedWorker.name || typed}. Time lookup is ready.`);
  return resolvedWorker;
}

function readRange(mode) {
  const now = new Date();
  const start = new Date(now);
  start.setHours(0, 0, 0, 0);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
  const end = new Date(start);
  end.setDate(end.getDate() + 6);
  end.setHours(23, 59, 59, 999);

  if (mode === 'more') {
    const fromValue = el('workerTimeFromInput')?.value;
    const toValue = el('workerTimeToInput')?.value;
    if (fromValue && toValue) {
      const from = new Date(`${fromValue}T00:00:00`);
      const to = new Date(`${toValue}T23:59:59.999`);
      if (Number.isFinite(from.getTime()) && Number.isFinite(to.getTime()) && from <= to) {
        return { fromMs: from.getTime(), toMs: to.getTime() };
      }
    }
    start.setDate(start.getDate() - 7);
  }
  return { fromMs: start.getTime(), toMs: end.getTime() };
}

async function loadWorkerPunches(worker, range) {
  const idCandidates = [...new Set([
    worker.id,
    worker.employeeId,
    worker.workerId,
  ].map((value) => String(value || '').trim()).filter(Boolean))];
  const nameKey = normalizeName(worker.name || worker.nameKey || '');
  const siteId = selectedSite();
  const jobs = [];

  idCandidates.forEach((id) => {
    jobs.push(safeQuery('punches', [
      where('companyId', '==', COMPANY_ID),
      where('siteId', '==', siteId),
      where('employeeId', '==', id),
    ]));
    jobs.push(safeQuery('punches', [
      where('companyId', '==', COMPANY_ID),
      where('siteId', '==', siteId),
      where('workerId', '==', id),
    ]));
  });
  jobs.push(safeQuery('punches', [
    where('companyId', '==', COMPANY_ID),
    where('siteId', '==', siteId),
    where('nameKey', '==', nameKey),
  ]));
  jobs.push(safeQuery('punches', [
    where('companyId', '==', COMPANY_ID),
    where('nameKey', '==', nameKey),
  ]));

  const rows = uniqueById((await Promise.all(jobs)).flat())
    .map((row) => ({ ...row, timestampMs: timestampMs(row) }))
    .filter((row) => row.timestampMs >= range.fromMs && row.timestampMs <= range.toMs)
    .filter((row) => row.status !== 'deleted' && row.active !== false)
    .filter((row) => {
      const rowIds = [row.employeeId, row.workerId].map((value) => String(value || '').trim()).filter(Boolean);
      const idMatch = rowIds.some((id) => idCandidates.includes(id));
      const legacyNameMatch = !rowIds.length && normalizeName(row.nameKey || row.name || '') === nameKey;
      return idMatch || legacyNameMatch;
    })
    .sort((left, right) => left.timestampMs - right.timestampMs);

  return rows;
}

function summarizePunches(rows) {
  const byDate = new Map();
  rows.forEach((row) => {
    const dateKey = row.dateKey || new Date(row.timestampMs).toISOString().slice(0, 10);
    if (!byDate.has(dateKey)) byDate.set(dateKey, []);
    byDate.get(dateKey).push(row);
  });

  let totalMinutes = 0;
  const days = [];
  [...byDate.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([dateKey, punches]) => {
    let activeStart = null;
    let minutes = 0;
    const actions = {};
    punches.forEach((punch) => {
      if (!(punch.action in actions)) actions[punch.action] = punch.timestampMs;
      if (punch.action === 'clock_in' || punch.action === 'end_lunch') activeStart = punch.timestampMs;
      if ((punch.action === 'start_lunch' || punch.action === 'clock_out') && activeStart) {
        minutes += Math.max(0, Math.round((punch.timestampMs - activeStart) / 60000));
        activeStart = null;
      }
    });
    totalMinutes += minutes;
    days.push({ dateKey, minutes, actions });
  });
  return { days, totalMinutes };
}

function renderSummary(rows) {
  const summary = summarizePunches(rows);
  const hours = summary.totalMinutes / 60;
  if (el('workerWeekHoursValue')) el('workerWeekHoursValue').textContent = hours.toFixed(2);
  if (el('workerRegularHoursValue')) el('workerRegularHoursValue').textContent = Math.min(hours, 40).toFixed(2);
  if (el('workerOvertimeHoursValue')) el('workerOvertimeHoursValue').textContent = Math.max(0, hours - 40).toFixed(2);
  if (el('workerDaysWorkedValue')) el('workerDaysWorkedValue').textContent = String(summary.days.length);

  const results = el('workerTimeRangeResults');
  if (results) {
    results.innerHTML = summary.days.length
      ? summary.days.slice().reverse().map((day) => `
        <article class="time-result-card">
          <div class="time-result-head">
            <strong>${new Date(`${day.dateKey}T12:00:00`).toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' })}</strong>
            <span>${(day.minutes / 60).toFixed(2)} hrs</span>
          </div>
          <div class="time-result-grid">
            <span>Clock In<strong>${day.actions.clock_in ? new Date(day.actions.clock_in).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '-'}</strong></span>
            <span>Start Lunch<strong>${day.actions.start_lunch ? new Date(day.actions.start_lunch).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '-'}</strong></span>
            <span>End Lunch<strong>${day.actions.end_lunch ? new Date(day.actions.end_lunch).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '-'}</strong></span>
            <span>Clock Out<strong>${day.actions.clock_out ? new Date(day.actions.clock_out).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '-'}</strong></span>
          </div>
        </article>`).join('')
      : '<div class="empty-state">No punches found for this worker in the selected range.</div>';
  }
  setRangeStatus(`Total Hours: ${hours.toFixed(2)} from ${summary.days.length} day(s).`);
}

async function handleTimeLookup(mode) {
  el('workerMyTimePanel')?.classList.remove('hidden');
  el('workerFixPanel')?.classList.add('hidden');
  el('workerTimeRangeControls')?.classList.toggle('hidden', mode !== 'more');
  setRangeStatus('Looking up saved punches...');
  try {
    const worker = await resolveWorkerByTypedName();
    const rows = await loadWorkerPunches(worker, readRange(mode));
    renderSummary(rows);
  } catch (error) {
    setRangeStatus(error.message || 'Time lookup failed.');
    setStatus(error.message || 'Time lookup failed.', true);
  }
}

async function handleFixSubmit(event) {
  event.preventDefault();
  event.stopImmediatePropagation();
  try {
    const worker = await resolveWorkerByTypedName();
    const action = String(el('workerFixActionInput')?.value || '');
    const dateValue = String(el('workerFixDateInput')?.value || '');
    const timeValue = String(el('workerFixTimeInput')?.value || '');
    const reason = String(el('workerFixReasonInput')?.value || '').trim();
    if (!VALID_ACTIONS.has(action) || !dateValue || !timeValue || reason.length < 2) {
      throw new Error('Choose the punch, date, time, and enter a reason.');
    }
    const requestedDate = new Date(`${dateValue}T${timeValue}:00`);
    if (!Number.isFinite(requestedDate.getTime())) throw new Error('Choose a valid requested date and time.');

    await addDoc(collection(db, 'missedPunchRequests'), {
      companyId: COMPANY_ID,
      siteId: selectedSite(),
      assignedSiteId: selectedSite(),
      agencyId: worker.agencyId || '',
      employeeId: workerId(worker),
      workerId: worker.workerId || worker.id || workerId(worker),
      employeeNumber: worker.employeeNumber || '',
      name: worker.name || String(el('workerNameInput')?.value || '').trim(),
      nameKey: normalizeName(worker.name || el('workerNameInput')?.value || ''),
      requestedAction: action,
      requestedTimestampMs: requestedDate.getTime(),
      requestedDateKey: dateValue,
      requestedTime: timeValue,
      reason,
      status: 'pending',
      source: 'public_worker',
      createdAt: serverTimestamp(),
    });
    setStatus(`Time fix request submitted for ${actionLabel(action)} on ${dateValue}.`);
    if (el('workerFixReasonInput')) el('workerFixReasonInput').value = '';
  } catch (error) {
    const inactiveMessage = resolvedWorker && !isActive(resolvedWorker)
      ? ' This worker is inactive; a manager may need to reactivate the profile before the request can be accepted.'
      : '';
    setStatus(`${error.message || 'Could not submit the time fix request.'}${inactiveMessage}`, true);
  }
}

function interceptClick(buttonId, mode) {
  el(buttonId)?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopImmediatePropagation();
    handleTimeLookup(mode);
  }, true);
}

function install() {
  interceptClick('workerViewTimeBtn', 'week');
  interceptClick('workerViewMoreTimeBtn', 'more');
  el('workerRequestFixBtn')?.addEventListener('click', (event) => {
    event.preventDefault();
    event.stopImmediatePropagation();
    el('workerFixPanel')?.classList.remove('hidden');
    el('workerMyTimePanel')?.classList.add('hidden');
    resolveWorkerByTypedName().catch((error) => setStatus(error.message, true));
  }, true);
  el('workerFixForm')?.addEventListener('submit', handleFixSubmit, true);
  el('workerBranchSelect')?.addEventListener('change', () => { resolvedWorker = null; });
  el('workerNameInput')?.addEventListener('input', () => { resolvedWorker = null; });
  console.info('[QRTimeclock] Temp self-service compatibility layer installed.');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', install, { once: true });
} else {
  install();
}
