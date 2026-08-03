// Temp worker hour shortcuts.
// Uses the existing read-only worker time lookup so this feature never creates,
// edits, deletes, or rewrites punches.

function localDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function startOfWeek(date) {
  const result = new Date(date);
  result.setHours(0, 0, 0, 0);
  const day = result.getDay();
  result.setDate(result.getDate() - (day === 0 ? 6 : day - 1));
  return result;
}

function addDays(date, amount) {
  const result = new Date(date);
  result.setDate(result.getDate() + amount);
  return result;
}

function setRangeAndLookup(fromDate, toDate, label) {
  const fromInput = document.getElementById('workerTimeFromInput');
  const toInput = document.getElementById('workerTimeToInput');
  const lookupButton = document.getElementById('workerTimeLookupBtn');
  const controls = document.getElementById('workerTimeRangeControls');
  const panel = document.getElementById('workerMyTimePanel');
  const status = document.getElementById('workerTimeRangeStatus');

  if (!fromInput || !toInput || !lookupButton) return;

  panel?.classList.remove('hidden');
  controls?.classList.remove('hidden');
  fromInput.value = localDateValue(fromDate);
  toInput.value = localDateValue(toDate);
  if (status) status.textContent = `Loading ${label.toLowerCase()}...`;
  lookupButton.click();
}

function buildWorkerHourShortcuts() {
  const panel = document.getElementById('workerMyTimePanel');
  const controls = document.getElementById('workerTimeRangeControls');
  const mainButton = document.getElementById('workerViewTimeBtn');
  const moreButton = document.getElementById('workerViewMoreTimeBtn');
  if (!panel || !controls || !mainButton || panel.dataset.hourShortcutsReady === 'true') return;

  panel.dataset.hourShortcutsReady = 'true';
  mainButton.textContent = 'My Hours';
  if (moreButton) moreButton.textContent = 'Choose Dates';

  const heading = panel.querySelector('.card-head h3');
  const description = panel.querySelector('.card-head p');
  if (heading) heading.textContent = 'My Hours and Punches';
  if (description) {
    description.textContent = 'See today, last week, all recorded hours, or choose your own dates. This is read-only and cannot change a punch.';
  }

  const shortcuts = document.createElement('div');
  shortcuts.className = 'quick-range-actions worker-hours-shortcuts';
  shortcuts.setAttribute('aria-label', 'Worker hour ranges');

  const ranges = [
    {
      label: 'Today',
      getDates() {
        const today = new Date();
        return [today, today];
      },
    },
    {
      label: 'Last Week',
      getDates() {
        const lastMonday = addDays(startOfWeek(new Date()), -7);
        return [lastMonday, addDays(lastMonday, 6)];
      },
    },
    {
      label: 'Overall Hours',
      getDates() {
        // Safely includes every record created by this app while keeping the
        // existing employee identity and branch/agency filters in control.
        return [new Date(2020, 0, 1), new Date()];
      },
    },
  ];

  ranges.forEach((range, index) => {
    const button = document.createElement('button');
    button.type = 'button';
    button.className = index === 2 ? 'primary-btn' : 'ghost-btn';
    button.textContent = range.label;
    button.addEventListener('click', () => {
      const [fromDate, toDate] = range.getDates();
      setRangeAndLookup(fromDate, toDate, range.label);
    });
    shortcuts.appendChild(button);
  });

  controls.parentNode?.insertBefore(shortcuts, controls);

  mainButton.addEventListener('click', () => {
    window.setTimeout(() => {
      panel.classList.remove('hidden');
      controls.classList.remove('hidden');
    }, 0);
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', buildWorkerHourShortcuts, { once: true });
} else {
  buildWorkerHourShortcuts();
}

// The main module may finish rendering after this compatibility module.
window.setTimeout(buildWorkerHourShortcuts, 250);
window.setTimeout(buildWorkerHourShortcuts, 1000);
