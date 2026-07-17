const STORAGE_KEY = 'qrtimeclock.correctionSchedule.v1';

const DEFAULT_SCHEDULE = {
  clock_in: '07:00',
  start_lunch: '11:00',
  end_lunch: '11:30',
  clock_out: '15:30',
  graceMinutes: 10
};

const ACTION_LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Lunch Out',
  end_lunch: 'Lunch In',
  clock_out: 'Clock Out'
};

function loadSchedule() {
  try {
    return { ...DEFAULT_SCHEDULE, ...JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}') };
  } catch {
    return { ...DEFAULT_SCHEDULE };
  }
}

function saveSchedule(schedule) {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(schedule));
}

function normalizeAction(value) {
  const text = String(value || '').trim().toLowerCase().replace(/\s+/g, '_');
  if (['clock_in', 'clockin'].includes(text)) return 'clock_in';
  if (['lunch_out', 'start_lunch', 'lunchout'].includes(text)) return 'start_lunch';
  if (['lunch_in', 'end_lunch', 'lunchin'].includes(text)) return 'end_lunch';
  if (['clock_out', 'clockout'].includes(text)) return 'clock_out';
  return '';
}

function parseDisplayedTime(value) {
  const text = String(value || '').trim();
  const match = text.match(/(\d{1,2}):(\d{2})(?::\d{2})?\s*(AM|PM)?/i);
  if (!match) return null;
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const suffix = (match[3] || '').toUpperCase();
  if (suffix === 'PM' && hour !== 12) hour += 12;
  if (suffix === 'AM' && hour === 12) hour = 0;
  return hour * 60 + minute;
}

function timeToMinutes(value) {
  const [hour, minute] = String(value || '00:00').split(':').map(Number);
  return hour * 60 + minute;
}

function formatTime(value) {
  const [hourText, minute] = String(value).split(':');
  const hour = Number(hourText);
  const suffix = hour >= 12 ? 'PM' : 'AM';
  const displayHour = hour % 12 || 12;
  return `${displayHour}:${minute} ${suffix}`;
}

function readTodayPunches() {
  const body = document.getElementById('livePunchBody');
  if (!body) return [];
  return [...body.querySelectorAll('tr')].map((row) => {
    const cells = [...row.querySelectorAll('td')].map((cell) => cell.textContent.trim());
    if (cells.length < 3 || cells[0].toLowerCase().includes('no live data')) return null;
    return {
      timeText: cells[0],
      timeMinutes: parseDisplayedTime(cells[0]),
      name: cells[1],
      action: normalizeAction(cells[2])
    };
  }).filter((item) => item && item.name && item.action && item.timeMinutes !== null);
}

function analyzePunches(schedule) {
  const punches = readTodayPunches();
  const grouped = new Map();
  punches.forEach((punch) => {
    if (!grouped.has(punch.name)) grouped.set(punch.name, []);
    grouped.get(punch.name).push(punch);
  });

  const now = new Date();
  const nowMinutes = now.getHours() * 60 + now.getMinutes();
  const grace = Math.max(0, Number(schedule.graceMinutes) || 0);
  const issues = [];

  grouped.forEach((employeePunches, name) => {
    const byAction = new Map();
    employeePunches.forEach((punch) => {
      if (!byAction.has(punch.action)) byAction.set(punch.action, []);
      byAction.get(punch.action).push(punch);
    });

    Object.keys(ACTION_LABELS).forEach((action) => {
      const scheduledMinutes = timeToMinutes(schedule[action]);
      const actionPunches = byAction.get(action) || [];
      if (actionPunches.length > 1) {
        issues.push({ name, issue: `Duplicate ${ACTION_LABELS[action]}`, scheduled: schedule[action], actual: actionPunches.map((p) => p.timeText).join(', '), severity: 'warning' });
      }
      if (nowMinutes > scheduledMinutes + grace && actionPunches.length === 0) {
        issues.push({ name, issue: `Missing ${ACTION_LABELS[action]}`, scheduled: schedule[action], actual: '—', severity: 'missing' });
      }
      if (actionPunches.length > 0) {
        const actual = actionPunches[0];
        const difference = Math.abs(actual.timeMinutes - scheduledMinutes);
        if (difference > grace) {
          issues.push({ name, issue: `${ACTION_LABELS[action]} outside schedule`, scheduled: schedule[action], actual: actual.timeText, severity: 'warning' });
        }
      }
    });

    const ordered = employeePunches.slice().sort((a, b) => a.timeMinutes - b.timeMinutes);
    const expectedOrder = ['clock_in', 'start_lunch', 'end_lunch', 'clock_out'];
    let lastIndex = -1;
    for (const punch of ordered) {
      const currentIndex = expectedOrder.indexOf(punch.action);
      if (currentIndex < lastIndex) {
        issues.push({ name, issue: 'Punches out of order', scheduled: 'Clock In → Lunch Out → Lunch In → Clock Out', actual: ordered.map((p) => ACTION_LABELS[p.action]).join(' → '), severity: 'warning' });
        break;
      }
      lastIndex = currentIndex;
    }
  });

  return { issues, employeeCount: grouped.size };
}

function renderDashboard() {
  const schedule = loadSchedule();
  const result = analyzePunches(schedule);
  const count = document.getElementById('correctionIssueCount');
  const employeeCount = document.getElementById('correctionEmployeeCount');
  const status = document.getElementById('correctionStatus');
  const body = document.getElementById('correctionDashboardBody');
  if (!count || !employeeCount || !status || !body) return;

  count.textContent = String(result.issues.length);
  employeeCount.textContent = String(result.employeeCount);
  status.textContent = result.issues.length
    ? `${result.issues.length} possible correction issue${result.issues.length === 1 ? '' : 's'} found from today's live punch feed.`
    : result.employeeCount
      ? 'No correction issues found for employees currently shown in today’s live feed.'
      : 'No employee punches are currently available in the live feed.';

  if (!result.issues.length) {
    body.innerHTML = '<tr><td colspan="5">No corrections currently detected.</td></tr>';
    return;
  }

  body.innerHTML = result.issues.map((item) => `
    <tr>
      <td>${escapeHtml(item.name)}</td>
      <td><strong>${escapeHtml(item.issue)}</strong></td>
      <td>${escapeHtml(formatMaybeTime(item.scheduled))}</td>
      <td>${escapeHtml(item.actual)}</td>
      <td><button class="ghost-btn correction-open-editor" type="button" data-name="${escapeAttribute(item.name)}">Open Punch Editor</button></td>
    </tr>
  `).join('');

  body.querySelectorAll('.correction-open-editor').forEach((button) => {
    button.addEventListener('click', () => {
      const editButton = document.getElementById('editPunchesTabBtn');
      editButton?.click();
      const filter = document.getElementById('editFilterNameInput');
      if (filter) {
        filter.value = button.dataset.name || '';
        filter.dispatchEvent(new Event('input', { bubbles: true }));
        filter.focus();
      }
    });
  });
}

function formatMaybeTime(value) {
  return /^\d{2}:\d{2}$/.test(String(value)) ? formatTime(value) : value;
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
}

function escapeAttribute(value) {
  return escapeHtml(value).replace(/`/g, '&#96;');
}

function injectDashboard() {
  if (document.getElementById('correctionsTabBtn')) return;
  const tabBar = document.getElementById('tabBar');
  const managerTab = document.getElementById('managerTab');
  if (!tabBar || !managerTab) return;

  const tabButton = document.createElement('button');
  tabButton.id = 'correctionsTabBtn';
  tabButton.className = 'tab';
  tabButton.type = 'button';
  tabButton.textContent = 'Needs Corrections';
  tabBar.insertBefore(tabButton, document.getElementById('timesheetsTabBtn'));

  const panel = document.createElement('section');
  panel.id = 'correctionsTab';
  panel.className = 'tab-panel hidden';
  panel.innerHTML = `
    <div class="card">
      <div class="card-head split-head">
        <div>
          <h2>Needs Time Correction</h2>
          <p>Trial read-only checker. It reviews today’s existing live punches and never changes Firestore data.</p>
        </div>
        <button id="refreshCorrectionsBtn" class="primary-btn" type="button">Refresh Check</button>
      </div>
      <div class="grid-form" style="margin-bottom:18px;">
        <label><span>Clock In</span><input id="scheduleClockIn" type="time"></label>
        <label><span>Lunch Out</span><input id="scheduleLunchOut" type="time"></label>
        <label><span>Lunch In</span><input id="scheduleLunchIn" type="time"></label>
        <label><span>Clock Out</span><input id="scheduleClockOut" type="time"></label>
        <label><span>Grace minutes</span><input id="scheduleGrace" type="number" min="0" max="60"></label>
        <div class="form-actions"><button id="saveCorrectionScheduleBtn" class="secondary-btn" type="button">Save Trial Schedule</button></div>
      </div>
      <div class="stats-grid" style="margin-bottom:18px;">
        <div class="stat-card"><span>Employees checked</span><strong id="correctionEmployeeCount">0</strong></div>
        <div class="stat-card"><span>Possible issues</span><strong id="correctionIssueCount">0</strong></div>
      </div>
      <div id="correctionStatus" class="status-box">Waiting for today’s live punch feed.</div>
      <div class="mini-table-wrap tall" style="margin-top:18px;">
        <table>
          <thead><tr><th>Employee</th><th>Issue</th><th>Scheduled</th><th>Actual</th><th>Action</th></tr></thead>
          <tbody id="correctionDashboardBody"><tr><td colspan="5">No check has run yet.</td></tr></tbody>
        </table>
      </div>
    </div>`;
  managerTab.insertAdjacentElement('afterend', panel);

  const schedule = loadSchedule();
  document.getElementById('scheduleClockIn').value = schedule.clock_in;
  document.getElementById('scheduleLunchOut').value = schedule.start_lunch;
  document.getElementById('scheduleLunchIn').value = schedule.end_lunch;
  document.getElementById('scheduleClockOut').value = schedule.clock_out;
  document.getElementById('scheduleGrace').value = schedule.graceMinutes;

  tabButton.addEventListener('click', () => {
    document.querySelectorAll('.tab').forEach((button) => button.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach((item) => item.classList.add('hidden'));
    tabButton.classList.add('active');
    panel.classList.remove('hidden');
    renderDashboard();
  });

  document.getElementById('refreshCorrectionsBtn').addEventListener('click', renderDashboard);
  document.getElementById('saveCorrectionScheduleBtn').addEventListener('click', () => {
    saveSchedule({
      clock_in: document.getElementById('scheduleClockIn').value || DEFAULT_SCHEDULE.clock_in,
      start_lunch: document.getElementById('scheduleLunchOut').value || DEFAULT_SCHEDULE.start_lunch,
      end_lunch: document.getElementById('scheduleLunchIn').value || DEFAULT_SCHEDULE.end_lunch,
      clock_out: document.getElementById('scheduleClockOut').value || DEFAULT_SCHEDULE.clock_out,
      graceMinutes: Number(document.getElementById('scheduleGrace').value) || 0
    });
    renderDashboard();
  });

  const liveBody = document.getElementById('livePunchBody');
  if (liveBody) new MutationObserver(() => {
    if (!panel.classList.contains('hidden')) renderDashboard();
  }).observe(liveBody, { childList: true, subtree: true });
}

function start() {
  injectDashboard();
  if (!document.getElementById('correctionsTabBtn')) setTimeout(start, 500);
}

start();
