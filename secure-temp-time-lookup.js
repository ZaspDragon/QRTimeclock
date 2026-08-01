// Public temp time lookup through the sanitized Firebase HTTPS endpoint.
// This module intentionally does not read the punches collection from the browser.
// A one-time compatibility fallback remains only until the endpoint is deployed;
// once Firestore rules are secured, that legacy browser query is denied automatically.

const LOOKUP_ENDPOINT = 'https://us-central1-qrtimeclock-42764.cloudfunctions.net/publicWorkerTimeLookup';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_AGENCIES = new Set([
  'sterling_staffing',
  'excel_staffing',
  'lifestyle_staffing',
]);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
let lookupBusy = false;

const el = (id) => document.getElementById(id);

function selectedSite() {
  const value = String(el('workerBranchSelect')?.value || '').trim().toUpperCase();
  return VALID_SITES.has(value) ? value : 'OH01';
}

function selectedAgency() {
  const value = String(el('workerAgencySelect')?.value || '').trim();
  return VALID_AGENCIES.has(value) ? value : '';
}

function enteredName() {
  return String(el('workerNameInput')?.value || '').trim().replace(/\s+/g, ' ');
}

function setLookupStatus(message, error = false) {
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

function setLookupButtonsDisabled(disabled) {
  ['workerViewTimeBtn', 'workerViewMoreTimeBtn'].forEach((id) => {
    const button = el(id);
    if (button) button.disabled = disabled;
  });
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
    const fromValue = String(el('workerTimeFromInput')?.value || '');
    const toValue = String(el('workerTimeToInput')?.value || '');
    if (fromValue && toValue) {
      const from = new Date(`${fromValue}T00:00:00`);
      const to = new Date(`${toValue}T23:59:59.999`);
      if (Number.isFinite(from.getTime()) && Number.isFinite(to.getTime()) && from <= to) {
        return { fromMs: from.getTime(), toMs: to.getTime() };
      }
      throw new Error('Choose a valid date range.');
    }
    start.setDate(start.getDate() - 7);
  }

  return { fromMs: start.getTime(), toMs: end.getTime() };
}

function summarizePunches(rows) {
  const byDate = new Map();
  rows.forEach((row) => {
    if (!VALID_ACTIONS.has(row.action) || !Number.isFinite(Number(row.timestampMs))) return;
    const dateKey = String(row.dateKey || new Date(Number(row.timestampMs)).toISOString().slice(0, 10));
    if (!byDate.has(dateKey)) byDate.set(dateKey, []);
    byDate.get(dateKey).push({ ...row, timestampMs: Number(row.timestampMs) });
  });

  let totalMinutes = 0;
  const days = [];
  [...byDate.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([dateKey, punches]) => {
    punches.sort((left, right) => left.timestampMs - right.timestampMs);
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

function formatTime(value) {
  return value
    ? new Date(value).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })
    : '-';
}

function appendTimeCell(grid, label, value) {
  const cell = document.createElement('span');
  cell.append(document.createTextNode(label));
  const strong = document.createElement('strong');
  strong.textContent = formatTime(value);
  cell.append(strong);
  grid.append(cell);
}

function renderSummary(rows) {
  const summary = summarizePunches(Array.isArray(rows) ? rows : []);
  const hours = summary.totalMinutes / 60;

  if (el('workerWeekHoursValue')) el('workerWeekHoursValue').textContent = hours.toFixed(2);
  if (el('workerRegularHoursValue')) el('workerRegularHoursValue').textContent = Math.min(hours, 40).toFixed(2);
  if (el('workerOvertimeHoursValue')) el('workerOvertimeHoursValue').textContent = Math.max(0, hours - 40).toFixed(2);
  if (el('workerDaysWorkedValue')) el('workerDaysWorkedValue').textContent = String(summary.days.length);

  const results = el('workerTimeRangeResults');
  if (results) {
    results.replaceChildren();
    if (!summary.days.length) {
      const empty = document.createElement('div');
      empty.className = 'empty-state';
      empty.textContent = 'No punches were found for this worker in the selected range.';
      results.append(empty);
    } else {
      summary.days.slice().reverse().forEach((day) => {
        const card = document.createElement('article');
        card.className = 'time-result-card';

        const head = document.createElement('div');
        head.className = 'time-result-head';
        const date = document.createElement('strong');
        date.textContent = new Date(`${day.dateKey}T12:00:00`).toLocaleDateString([], {
          weekday: 'short',
          month: 'short',
          day: 'numeric',
        });
        const total = document.createElement('span');
        total.textContent = `${(day.minutes / 60).toFixed(2)} hrs`;
        head.append(date, total);

        const grid = document.createElement('div');
        grid.className = 'time-result-grid';
        appendTimeCell(grid, 'Clock In', day.actions.clock_in);
        appendTimeCell(grid, 'Lunch Out', day.actions.start_lunch);
        appendTimeCell(grid, 'Lunch In', day.actions.end_lunch);
        appendTimeCell(grid, 'Clock Out', day.actions.clock_out);

        card.append(head, grid);
        results.append(card);
      });
    }
  }

  setRangeStatus(`Total Hours: ${hours.toFixed(2)} from ${summary.days.length} day(s).`);
}

function legacyFallbackError(message) {
  const error = new Error(message);
  error.allowLegacyFallback = true;
  return error;
}

async function requestSecureTime(mode) {
  const name = enteredName();
  const siteId = selectedSite();
  const agencyId = selectedAgency();
  if (name.length < 2) throw new Error('Enter your first and last name.');
  if (!agencyId) throw new Error('Choose your staffing agency before checking time.');

  const range = readRange(mode);
  let response;
  try {
    response = await fetch(LOOKUP_ENDPOINT, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      cache: 'no-store',
      body: JSON.stringify({ name, siteId, agencyId, ...range }),
    });
  } catch {
    throw legacyFallbackError('Secure lookup is not reachable yet.');
  }

  let payload = null;
  try {
    payload = await response.json();
  } catch {
    payload = null;
  }

  if (!response.ok) {
    if (!payload || typeof payload.error !== 'string') {
      throw legacyFallbackError('Secure lookup is not deployed yet.');
    }
    throw new Error(payload.error);
  }
  return payload;
}

async function handleLookup(mode, button) {
  if (lookupBusy) return;
  lookupBusy = true;
  let useLegacyFallback = false;
  setLookupButtonsDisabled(true);
  el('workerMyTimePanel')?.classList.remove('hidden');
  el('workerFixPanel')?.classList.add('hidden');
  el('workerTimeRangeControls')?.classList.toggle('hidden', mode !== 'more');
  setRangeStatus('Securely looking up saved punches...');

  try {
    const payload = await requestSecureTime(mode);
    const workerName = String(payload?.worker?.name || enteredName());
    if (el('workerNameValue')) el('workerNameValue').textContent = workerName;
    setLookupStatus(`Found ${workerName}. Only sanitized time records were returned.`);
    renderSummary(payload?.punches || []);
  } catch (error) {
    if (error?.allowLegacyFallback) {
      useLegacyFallback = true;
      setRangeStatus('Secure lookup is being deployed. Using the existing lookup temporarily...');
      setLookupStatus('Secure lookup is being deployed. Existing time lookup remains available.');
    } else {
      const message = error?.message || 'Secure time lookup failed.';
      setRangeStatus(message);
      setLookupStatus(message, true);
    }
  } finally {
    lookupBusy = false;
    setLookupButtonsDisabled(false);
  }

  if (useLegacyFallback && button) {
    button.dataset.secureLookupBypass = 'true';
    window.setTimeout(() => button.click(), 0);
  }
}

document.addEventListener('click', (event) => {
  const target = event.target instanceof Element ? event.target : event.target?.parentElement;
  const button = target?.closest?.('#workerViewTimeBtn, #workerViewMoreTimeBtn');
  if (!button) return;
  if (button.dataset.secureLookupBypass === 'true') {
    delete button.dataset.secureLookupBypass;
    return;
  }
  event.preventDefault();
  event.stopImmediatePropagation();
  handleLookup(button.id === 'workerViewMoreTimeBtn' ? 'more' : 'week', button);
}, true);

console.info('[QRTimeclock] Secure public time lookup installed.');
