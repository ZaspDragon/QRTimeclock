// Exact-name public time lookup. Agency selection remains required for punching,
// but workers only need to type their name when viewing saved time.

const LOOKUP_ENDPOINT = 'https://us-central1-qrtimeclock-42764.cloudfunctions.net/publicWorkerTimeLookupByName';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
let lookupBusy = false;

const element = (id) => document.getElementById(id);

function selectedSite() {
  const value = String(element('workerBranchSelect')?.value || '').trim().toUpperCase();
  return VALID_SITES.has(value) ? value : 'OH01';
}

function enteredName() {
  return String(element('workerNameInput')?.value || '')
    .trim()
    .replace(/\s+/g, ' ');
}

function setLookupStatus(message, isError = false) {
  const status = element('workerLookupStatus');
  if (!status) return;
  status.textContent = message;
  status.style.borderColor = isError
    ? 'rgba(255,90,90,.55)'
    : 'rgba(43,213,118,.4)';
}

function setRangeStatus(message) {
  const status = element('workerTimeRangeStatus');
  if (status) status.textContent = message;
}

function setLookupButtonsDisabled(disabled) {
  [
    'workerViewTimeBtn',
    'workerViewMoreTimeBtn',
    'workerTimeLookupBtn',
  ].forEach((id) => {
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
    from.setDate(from.getDate() - 14);
  } else if (rangeName === 'this_month') {
    from = new Date(now.getFullYear(), now.getMonth(), 1);
    to = endOfDay(new Date(now.getFullYear(), now.getMonth() + 1, 0));
  }

  const fromInput = element('workerTimeFromInput');
  const toInput = element('workerTimeToInput');
  if (fromInput) fromInput.value = dateInputValue(from);
  if (toInput) toInput.value = dateInputValue(to);
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
  const days = [];
  [...byDate.entries()]
    .sort(([left], [right]) => left.localeCompare(right))
    .forEach(([dateKey, punches]) => {
      punches.sort((left, right) => left.timestampMs - right.timestampMs);
      let activeStart = null;
      let minutes = 0;
      const actions = {};

      punches.forEach((punch) => {
        if (!(punch.action in actions)) actions[punch.action] = punch.timestampMs;
        if (punch.action === 'clock_in' || punch.action === 'end_lunch') {
          activeStart = punch.timestampMs;
        }
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

  if (element('workerWeekHoursValue')) element('workerWeekHoursValue').textContent = hours.toFixed(2);
  if (element('workerRegularHoursValue')) element('workerRegularHoursValue').textContent = Math.min(hours, 40).toFixed(2);
  if (element('workerOvertimeHoursValue')) element('workerOvertimeHoursValue').textContent = Math.max(0, hours - 40).toFixed(2);
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

async function requestTimeByName(mode, suppliedRange = null) {
  const name = enteredName();
  if (name.length < 2) throw new Error('Enter the worker name.');
  const range = suppliedRange || readRange(mode);

  const response = await fetch(LOOKUP_ENDPOINT, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    cache: 'no-store',
    body: JSON.stringify({
      name,
      siteId: selectedSite(),
      ...range,
    }),
  });

  let payload = null;
  try {
    payload = await response.json();
  } catch {
    payload = null;
  }

  if (!response.ok) {
    throw new Error(payload?.error || 'Time lookup failed.');
  }
  return payload;
}

async function handleLookup(mode, suppliedRange = null) {
  if (lookupBusy) return;
  lookupBusy = true;
  setLookupButtonsDisabled(true);
  element('workerMyTimePanel')?.classList.remove('hidden');
  element('workerFixPanel')?.classList.add('hidden');
  element('workerTimeRangeControls')?.classList.toggle('hidden', mode === 'week');
  setRangeStatus('Looking up saved time by name...');

  try {
    const payload = await requestTimeByName(mode, suppliedRange);
    const workerName = String(payload?.worker?.name || enteredName());
    if (element('workerNameValue')) element('workerNameValue').textContent = workerName;
    setLookupStatus(`Found saved time for ${workerName}.`);
    renderSummary(payload?.punches || []);
  } catch (error) {
    const message = error?.message || 'Time lookup failed.';
    setLookupStatus(message, true);
    setRangeStatus(message);
    renderSummary([]);
  } finally {
    lookupBusy = false;
    setLookupButtonsDisabled(false);
  }
}

document.addEventListener('click', (event) => {
  const target = event.target instanceof Element ? event.target : event.target?.parentElement;
  const mainButton = target?.closest?.('#workerViewTimeBtn, #workerViewMoreTimeBtn, #workerTimeLookupBtn');
  const quickButton = target?.closest?.('.worker-range-quick');
  if (!mainButton && !quickButton) return;

  event.preventDefault();
  event.stopImmediatePropagation();

  if (quickButton) {
    const range = applyQuickRange(String(quickButton.dataset.range || 'this_week'));
    handleLookup('custom', range);
    return;
  }

  if (mainButton.id === 'workerViewTimeBtn') {
    handleLookup('week');
  } else if (mainButton.id === 'workerViewMoreTimeBtn') {
    handleLookup('more');
  } else {
    handleLookup('custom');
  }
}, true);

console.info('[QRTimeclock] Name-only time lookup installed.');
