// Exact-name public time lookup. Workers type their name; the branch and agency
// already selected on the clock page are used to safely locate their saved time.

const LOOKUP_ENDPOINT = 'https://us-central1-qrtimeclock-42764.cloudfunctions.net/publicWorkerTimeLookup';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_AGENCIES = new Set(['sterling_staffing', 'excel_staffing', 'lifestyle_staffing']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
let lookupBusy = false;
const LOOKUP_TIMEOUT_MS = 15000;
let lookupGeneration = 0;
let pendingController = null;

const element = (id) => document.getElementById(id);

function selectedSite() {
  const querySite = String(new URLSearchParams(location.search).get('site') || '').trim().toUpperCase();
  if (VALID_SITES.has(querySite)) return querySite;
  const value = String(element('workerBranchSelect')?.value || '').trim().toUpperCase();
  return VALID_SITES.has(value) ? value : 'OH01';
}

function selectedAgency() {
  // Use the visible selection. Lookup must never read or write punch-device caches.
  const value = String(element('workerAgencySelect')?.value || '').trim();
  return VALID_AGENCIES.has(value) ? value : '';
}

function enteredName() {
  return String(element('workerNameInput')?.value || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function clearSummary() {
  ['workerWeekHoursValue', 'workerRegularHoursValue', 'workerOvertimeHoursValue', 'workerDaysWorkedValue'].forEach((id) => {
    if (element(id)) element(id).textContent = '—';
  });
  element('workerTimeRangeResults')?.replaceChildren();
}

function setRangeStatus(message) {
  const status = element('workerTimeRangeStatus');
  if (status) status.textContent = message;
}

function setLookupButtonsDisabled(disabled) {
  ['workerViewTimeBtn', 'workerViewMoreTimeBtn', 'workerTimeLookupBtn'].forEach((id) => {
    const button = element(id);
    if (button) button.disabled = disabled;
  });
}

function mondayStart(date) {
  const start = new Date(date);
  start.setHours(0, 0, 0, 0);
  start.setDate(start.getDate() - ((start.getDay() + 6) % 7));
  return start;
}

function endOfDay(date) {
  const end = new Date(date);
  end.setHours(23, 59, 59, 999);
  return end;
}

function dateInputValue(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function applyQuickRange(rangeName) {
  const now = new Date();
  let from = mondayStart(now);
  let to = endOfDay(new Date(from.getFullYear(), from.getMonth(), from.getDate() + 6));

  if (rangeName === 'last_week') {
    from.setDate(from.getDate() - 7);
    to.setDate(to.getDate() - 7);
  } else if (rangeName === 'last_2_weeks') {
    from.setDate(from.getDate() - 7);
  } else if (rangeName === 'this_month') {
    from = new Date(now.getFullYear(), now.getMonth(), 1);
    to = endOfDay(new Date(now.getFullYear(), now.getMonth() + 1, 0));
  }

  if (element('workerTimeFromInput')) element('workerTimeFromInput').value = dateInputValue(from);
  if (element('workerTimeToInput')) element('workerTimeToInput').value = dateInputValue(to);
  return { fromMs: from.getTime(), toMs: to.getTime() };
}

function readRange(mode) {
  const now = new Date();
  const start = mondayStart(now);
  const end = endOfDay(new Date(start.getFullYear(), start.getMonth(), start.getDate() + 6));

  if (mode === 'custom' || mode === 'more') {
    const fromValue = String(element('workerTimeFromInput')?.value || '');
    const toValue = String(element('workerTimeToInput')?.value || '');
    if (fromValue && toValue) {
      const from = new Date(`${fromValue}T00:00:00`);
      const to = new Date(`${toValue}T23:59:59.999`);
      if (Number.isFinite(from.getTime()) && Number.isFinite(to.getTime()) && from <= to) {
        return { fromMs: from.getTime(), toMs: to.getTime() };
      }
      throw new Error('Choose a valid date range.');
    }
    if (mode === 'custom') throw new Error('Choose both a from date and a to date.');
    start.setDate(start.getDate() - 7);
  }

  return { fromMs: start.getTime(), toMs: end.getTime() };
}

function summarizePunches(rows) {
  const byDate = new Map();
  rows.forEach((row) => {
    const timestamp = Number(row.timestampMs);
    if (!VALID_ACTIONS.has(row.action) || !Number.isFinite(timestamp)) return;
    const dateKey = String(row.dateKey || new Date(timestamp).toISOString().slice(0, 10));
    if (!byDate.has(dateKey)) byDate.set(dateKey, []);
    byDate.get(dateKey).push({ ...row, timestampMs: timestamp });
  });

  let totalMinutes = 0;
  const weeklyMinutes = new Map();
  const days = [];
  [...byDate.entries()].sort(([left], [right]) => left.localeCompare(right)).forEach(([dateKey, punches]) => {
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

    const week = dateInputValue(mondayStart(new Date(`${dateKey}T12:00:00`)));
    weeklyMinutes.set(week, (weeklyMinutes.get(week) || 0) + minutes);
    totalMinutes += minutes;
    days.push({ dateKey, minutes, actions });
  });

  const regularMinutes = [...weeklyMinutes.values()].reduce((total, minutes) => total + Math.min(minutes, 40 * 60), 0);
  return { days, totalMinutes, regularMinutes, overtimeMinutes: totalMinutes - regularMinutes };
}

function formatTime(value) {
  return value ? new Date(value).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '-';
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

  if (element('workerWeekHoursValue')) element('workerWeekHoursValue').textContent = hours.toFixed(2);
  if (element('workerRegularHoursValue')) element('workerRegularHoursValue').textContent = (summary.regularMinutes / 60).toFixed(2);
  if (element('workerOvertimeHoursValue')) element('workerOvertimeHoursValue').textContent = (summary.overtimeMinutes / 60).toFixed(2);
  if (element('workerDaysWorkedValue')) element('workerDaysWorkedValue').textContent = String(summary.days.length);

  const results = element('workerTimeRangeResults');
  if (results) {
    results.replaceChildren();
    if (!summary.days.length) {
      const empty = document.createElement('div');
      empty.className = 'empty-state';
      empty.textContent = 'No punches were found for this name in the selected range.';
      results.append(empty);
    } else {
      summary.days.slice().reverse().forEach((day) => {
        const card = document.createElement('article');
        card.className = 'time-result-card';
        const head = document.createElement('div');
        head.className = 'time-result-head';
        const date = document.createElement('strong');
        date.textContent = new Date(`${day.dateKey}T12:00:00`).toLocaleDateString([], { weekday: 'short', month: 'short', day: 'numeric' });
        const total = document.createElement('span');
        total.textContent = `${(day.minutes / 60).toFixed(2)} hrs`;
        head.append(date, total);
        const grid = document.createElement('div');
        grid.className = 'time-result-grid';
        appendTimeCell(grid, 'Clock In', day.actions.clock_in);
        appendTimeCell(grid, 'Start Lunch', day.actions.start_lunch);
        appendTimeCell(grid, 'End Lunch', day.actions.end_lunch);
        appendTimeCell(grid, 'Clock Out', day.actions.clock_out);
        card.append(head, grid);
        results.append(card);
      });
    }
  }

  setRangeStatus(`Total Hours: ${hours.toFixed(2)} from ${summary.days.length} day(s).`);
}

async function requestTimeByName(mode, suppliedRange = null, signal = undefined) {
  const name = enteredName();
  if (name.length < 2) throw new Error('Enter the worker name.');
  const agencyId = selectedAgency();
  if (!agencyId) throw new Error('Choose the worker staffing agency before viewing time.');
  const range = suppliedRange || readRange(mode);
  // One extra hour accommodates a 31-day range crossing the autumn DST change.
  if (range.toMs - range.fromMs > 31 * 86400000 + 3600000) {
    throw new Error('Choose up to 31 days at a time.');
  }

  const response = await fetch(LOOKUP_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    cache: 'no-store',
    signal,
    body: JSON.stringify({ name, siteId: selectedSite(), agencyId, ...range }),
  });

  let payload = null;
  try { payload = await response.json(); } catch { payload = null; }
  if (!response.ok) {
    if (response.status === 404 && !payload?.error) {
      throw new Error('The hours service is unavailable. Ask your manager to enable time lookup. You can still clock in and out.');
    }
    throw new Error(payload?.error || 'Could not load your hours. Please try again. You can still clock in and out.');
  }
  if (!payload || !Array.isArray(payload.punches) || payload.punches.some((row) =>
    !row || !VALID_ACTIONS.has(row.action) || !Number.isFinite(row.timestampMs)
    || (row.dateKey && !/^\d{4}-\d{2}-\d{2}$/.test(row.dateKey)))) {
    throw new Error('The hours service returned an incomplete response. Please try again.');
  }
  return payload;
}

async function handleLookup(mode, suppliedRange = null) {
  if (lookupBusy) return;
  lookupBusy = true;
  const generation = ++lookupGeneration;
  const controller = new AbortController();
  pendingController = controller;
  const timeout = window.setTimeout(() => controller.abort(), LOOKUP_TIMEOUT_MS);
  setLookupButtonsDisabled(true);
  element('workerMyTimePanel')?.classList.remove('hidden');
  element('workerFixPanel')?.classList.add('hidden');
  element('workerTimeRangeControls')?.classList.toggle('hidden', mode === 'week');
  setRangeStatus('Looking up saved time by name...');
  clearSummary();

  try {
    const payload = await requestTimeByName(mode, suppliedRange, controller.signal);
    if (generation !== lookupGeneration) return;
    const workerName = String(payload?.worker?.name || enteredName());
    renderSummary(payload?.punches || []);
    setRangeStatus(`${workerName} · ${element('workerTimeRangeStatus')?.textContent || ''} Refreshed ${new Date().toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' })}.`);
  } catch (error) {
    if (generation !== lookupGeneration) return;
    const message = error?.name === 'AbortError'
      ? 'Loading hours timed out. Please try again. You can still clock in and out.'
      : error instanceof TypeError
        ? 'Could not connect to the hours service. Please try again. You can still clock in and out.'
        : error?.message || 'Could not load your hours. Please try again.';
    clearSummary();
    setRangeStatus(message);
  } finally {
    window.clearTimeout(timeout);
    if (generation === lookupGeneration) {
      pendingController = null;
      lookupBusy = false;
      setLookupButtonsDisabled(false);
    }
  }
}

document.addEventListener('click', (event) => {
  const target = event.target instanceof Element ? event.target : event.target?.parentElement;
  const mainButton = target?.closest?.('#workerViewTimeBtn, #workerViewMoreTimeBtn, #workerTimeLookupBtn');
  const quickButton = target?.closest?.('.worker-range-quick');
  if (!mainButton && !quickButton) return;

  event.preventDefault();
  event.stopImmediatePropagation();
  if (lookupBusy) return;

  if (quickButton) {
    handleLookup('custom', applyQuickRange(String(quickButton.dataset.range || 'this_week')));
  } else if (mainButton.id === 'workerViewTimeBtn') {
    handleLookup('week');
  } else if (mainButton.id === 'workerViewMoreTimeBtn') {
    handleLookup('more');
  } else {
    handleLookup('custom');
  }
}, true);

// A previous worker's in-flight response must not appear under a new selection.
function invalidateLookup(event) {
  if (!['workerNameInput', 'workerBranchSelect', 'workerAgencySelect', 'workerTimeFromInput', 'workerTimeToInput'].includes(event.target?.id)) return;
  lookupGeneration += 1;
  pendingController?.abort();
  pendingController = null;
  lookupBusy = false;
  setLookupButtonsDisabled(false);
  clearSummary();
  setRangeStatus('Select My Hours or Look Up Past Time to refresh.');
}
document.addEventListener('input', invalidateLookup);
document.addEventListener('change', invalidateLookup);

console.info('[QRTimeclock] Corrected exact-name time lookup installed.');
