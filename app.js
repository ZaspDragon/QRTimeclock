const CONFIG = {
  webAppUrl: 'PASTE_YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL_HERE',
  timezoneLabel: 'America/New_York',
};

const state = {
  scanner: null,
  scanning: false,
  employee: null,
  managerLoggedIn: false,
  managerName: '',
  recentPunches: [],
};

const els = {};

document.addEventListener('DOMContentLoaded', init);

function init() {
  grabEls();
  bindEvents();
  setConnectionBadge();
  renderPosterQr();
  setDefaultManagerDate();
}

function grabEls() {
  Object.assign(els, {
    connectionBadge: document.getElementById('connectionBadge'),
    employeeSection: document.getElementById('employeeSection'),
    managerSection: document.getElementById('managerSection'),
    qrPosterSection: document.getElementById('qrPosterSection'),
    showEmployeeBtn: document.getElementById('showEmployeeBtn'),
    showManagerBtn: document.getElementById('showManagerBtn'),
    showQrPosterBtn: document.getElementById('showQrPosterBtn'),
    toggleScannerBtn: document.getElementById('toggleScannerBtn'),
    scannerWrap: document.getElementById('scannerWrap'),
    employeeCode: document.getElementById('employeeCode'),
    employeeName: document.getElementById('employeeName'),
    employeeStatus: document.getElementById('employeeStatus'),
    recentPunchesBody: document.getElementById('recentPunchesBody'),
    managerUsername: document.getElementById('managerUsername'),
    managerPassword: document.getElementById('managerPassword'),
    managerLoginBtn: document.getElementById('managerLoginBtn'),
    managerLoginStatus: document.getElementById('managerLoginStatus'),
    managerLoginBox: document.getElementById('managerLoginBox'),
    managerPanel: document.getElementById('managerPanel'),
    managerDate: document.getElementById('managerDate'),
    refreshManagerBtn: document.getElementById('refreshManagerBtn'),
    managerLogoutBtn: document.getElementById('managerLogoutBtn'),
    managerStatus: document.getElementById('managerStatus'),
    timesheetBody: document.getElementById('timesheetBody'),
    posterQr: document.getElementById('posterQr'),
    posterUrlText: document.getElementById('posterUrlText'),
    printPosterBtn: document.getElementById('printPosterBtn'),
    badgeCode: document.getElementById('badgeCode'),
    badgeName: document.getElementById('badgeName'),
    generateBadgeBtn: document.getElementById('generateBadgeBtn'),
    badgePreview: document.getElementById('badgePreview'),
  });
}

function bindEvents() {
  els.showEmployeeBtn.addEventListener('click', () => showSection('employee'));
  els.showManagerBtn.addEventListener('click', () => showSection('manager'));
  els.showQrPosterBtn.addEventListener('click', () => showSection('poster'));
  els.toggleScannerBtn.addEventListener('click', toggleScanner);
  els.employeeCode.addEventListener('change', handleEmployeeLookup);
  els.employeeCode.addEventListener('blur', handleEmployeeLookup);
  document.querySelectorAll('.action-btn').forEach(btn => {
    btn.addEventListener('click', () => submitPunch(btn.dataset.action));
  });
  els.managerLoginBtn.addEventListener('click', managerLogin);
  els.refreshManagerBtn.addEventListener('click', loadTimesheets);
  els.managerLogoutBtn.addEventListener('click', managerLogout);
  els.printPosterBtn.addEventListener('click', () => window.print());
  els.generateBadgeBtn.addEventListener('click', generateBadgeQr);
}

function showSection(section) {
  els.employeeSection.classList.toggle('hidden', section !== 'employee');
  els.managerSection.classList.toggle('hidden', section !== 'manager');
  els.qrPosterSection.classList.toggle('hidden', section !== 'poster');
}

function setConnectionBadge() {
  const connected = CONFIG.webAppUrl && !CONFIG.webAppUrl.includes('PASTE_');
  els.connectionBadge.textContent = connected ? 'Connected' : 'Add backend URL';
}

function setDefaultManagerDate() {
  const today = new Date();
  els.managerDate.value = toDateInput(today);
}

async function handleEmployeeLookup() {
  const employeeCode = els.employeeCode.value.trim();
  els.employeeName.value = '';
  state.employee = null;
  if (!employeeCode) return;

  setStatus(els.employeeStatus, 'Checking employee...', false);
  try {
    const res = await callApi('getEmployee', { employeeCode });
    if (!res.ok) throw new Error(res.message || 'Employee not found.');
    state.employee = res.employee;
    els.employeeName.value = res.employee.name || '';
    setStatus(els.employeeStatus, `Ready for ${res.employee.name}.`, true);
  } catch (err) {
    setStatus(els.employeeStatus, err.message, false);
  }
}

async function submitPunch(action) {
  const employeeCode = els.employeeCode.value.trim();
  if (!employeeCode) {
    setStatus(els.employeeStatus, 'Scan or enter an employee code first.', false);
    return;
  }

  if (!state.employee) {
    await handleEmployeeLookup();
    if (!state.employee) return;
  }

  const actionLabels = {
    clockIn: 'Clock In',
    startLunch: 'Start Lunch',
    endLunch: 'End Lunch',
    clockOut: 'Clock Out',
  };

  setStatus(els.employeeStatus, `${actionLabels[action]} sending...`, false);
  try {
    const res = await callApi('recordPunch', { employeeCode, action });
    if (!res.ok) throw new Error(res.message || 'Could not record punch.');

    const entry = {
      date: res.entry.workDate,
      employee: res.entry.employeeName,
      action: actionLabels[action],
      time: res.entry.localTime,
      status: 'Saved',
    };
    state.recentPunches.unshift(entry);
    state.recentPunches = state.recentPunches.slice(0, 10);
    renderRecentPunches();
    setStatus(els.employeeStatus, `${actionLabels[action]} saved for ${res.entry.employeeName} at ${res.entry.localTime}.`, true);
  } catch (err) {
    setStatus(els.employeeStatus, err.message, false);
  }
}

function renderRecentPunches() {
  if (!state.recentPunches.length) {
    els.recentPunchesBody.innerHTML = '<tr><td colspan="5" class="empty">No punches yet.</td></tr>';
    return;
  }

  els.recentPunchesBody.innerHTML = state.recentPunches.map(row => `
    <tr>
      <td>${escapeHtml(row.date)}</td>
      <td>${escapeHtml(row.employee)}</td>
      <td>${escapeHtml(row.action)}</td>
      <td>${escapeHtml(row.time)}</td>
      <td>${escapeHtml(row.status)}</td>
    </tr>
  `).join('');
}

async function managerLogin() {
  const username = els.managerUsername.value.trim();
  const password = els.managerPassword.value.trim();
  if (!username || !password) {
    setStatus(els.managerLoginStatus, 'Enter manager username and password.', false);
    return;
  }

  setStatus(els.managerLoginStatus, 'Checking manager...', false);
  try {
    const res = await callApi('managerLogin', { username, password });
    if (!res.ok) throw new Error(res.message || 'Login failed.');
    state.managerLoggedIn = true;
    state.managerName = res.manager.name || username;
    els.managerLoginBox.classList.add('hidden');
    els.managerPanel.classList.remove('hidden');
    setStatus(els.managerStatus, `Logged in as ${state.managerName}.`, true);
    loadTimesheets();
  } catch (err) {
    setStatus(els.managerLoginStatus, err.message, false);
  }
}

function managerLogout() {
  state.managerLoggedIn = false;
  state.managerName = '';
  els.managerLoginBox.classList.remove('hidden');
  els.managerPanel.classList.add('hidden');
  els.timesheetBody.innerHTML = '<tr><td colspan="8" class="empty">Log in to load timesheets.</td></tr>';
  els.managerUsername.value = '';
  els.managerPassword.value = '';
  setStatus(els.managerLoginStatus, 'Manager locked.', false);
}

async function loadTimesheets() {
  if (!state.managerLoggedIn) return;
  const workDate = els.managerDate.value;
  setStatus(els.managerStatus, 'Loading timesheets...', false);

  try {
    const res = await callApi('getTimesheets', { workDate });
    if (!res.ok) throw new Error(res.message || 'Could not load timesheets.');
    renderTimesheets(res.rows || []);
    setStatus(els.managerStatus, `Loaded ${res.rows.length} timesheet row(s).`, true);
  } catch (err) {
    setStatus(els.managerStatus, err.message, false);
  }
}

function renderTimesheets(rows) {
  if (!rows.length) {
    els.timesheetBody.innerHTML = '<tr><td colspan="8" class="empty">No timesheets for this date.</td></tr>';
    return;
  }

  els.timesheetBody.innerHTML = rows.map((row, i) => `
    <tr>
      <td>${escapeHtml(row.workDate || '')}</td>
      <td>${escapeHtml(row.employeeName || '')}</td>
      <td>${escapeHtml(row.clockIn || '')}</td>
      <td>${escapeHtml(row.lunchStart || '')}</td>
      <td>${escapeHtml(row.lunchEnd || '')}</td>
      <td>${escapeHtml(row.clockOut || '')}</td>
      <td>${escapeHtml(row.totalHours || '')}</td>
      <td>
        ${row.managerSigned
          ? `<span class="signed-pill">Signed by ${escapeHtml(row.managerSigned)}</span>`
          : `<button class="sign-btn" data-timesheet-id="${escapeHtml(row.timesheetId || '')}">Sign</button>`}
      </td>
    </tr>
  `).join('');

  els.timesheetBody.querySelectorAll('.sign-btn').forEach(btn => {
    btn.addEventListener('click', () => signTimesheet(btn.dataset.timesheetId));
  });
}

async function signTimesheet(timesheetId) {
  if (!timesheetId || !state.managerLoggedIn) return;
  setStatus(els.managerStatus, 'Signing timesheet...', false);

  try {
    const res = await callApi('signTimesheet', {
      timesheetId,
      managerName: state.managerName,
    });
    if (!res.ok) throw new Error(res.message || 'Could not sign timesheet.');
    setStatus(els.managerStatus, `Timesheet signed by ${state.managerName}.`, true);
    loadTimesheets();
  } catch (err) {
    setStatus(els.managerStatus, err.message, false);
  }
}

function renderPosterQr() {
  const url = window.location.href.split('#')[0];
  els.posterUrlText.textContent = url;
  if (window.QRCode) {
    els.posterQr.innerHTML = '';
    QRCode.toCanvas(url, { width: 230 }, (err, canvas) => {
      if (err) {
        els.posterQr.textContent = 'QR could not be generated.';
        return;
      }
      els.posterQr.appendChild(canvas);
    });
  }
}

function generateBadgeQr() {
  const code = els.badgeCode.value.trim();
  const name = els.badgeName.value.trim();
  if (!code || !name) {
    els.badgePreview.innerHTML = '<p class="tiny">Enter both employee code and employee name.</p>';
    return;
  }

  els.badgePreview.innerHTML = '';
  const wrap = document.createElement('div');
  wrap.className = 'badge-card';
  const title = document.createElement('h4');
  title.textContent = name;
  const sub = document.createElement('p');
  sub.textContent = code;
  sub.className = 'tiny';
  const qrHolder = document.createElement('div');
  wrap.append(title, sub, qrHolder);
  els.badgePreview.appendChild(wrap);

  if (window.QRCode) {
    QRCode.toCanvas(code, { width: 180 }, (err, canvas) => {
      if (err) {
        qrHolder.textContent = 'Badge QR could not be generated.';
        return;
      }
      qrHolder.appendChild(canvas);
    });
  }
}

async function toggleScanner() {
  if (!window.Html5Qrcode) {
    setStatus(els.employeeStatus, 'QR scanner library did not load.', false);
    return;
  }

  if (state.scanning) {
    await stopScanner();
    return;
  }

  els.scannerWrap.classList.remove('hidden');
  state.scanner = state.scanner || new Html5Qrcode('qrReader');
  try {
    await state.scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 220, height: 220 } },
      decodedText => {
        els.employeeCode.value = decodedText.trim();
        handleEmployeeLookup();
        stopScanner();
      }
    );
    state.scanning = true;
    els.toggleScannerBtn.textContent = 'Close QR Scanner';
  } catch (err) {
    setStatus(els.employeeStatus, `Scanner error: ${err.message || err}`, false);
    els.scannerWrap.classList.add('hidden');
  }
}

async function stopScanner() {
  try {
    if (state.scanner && state.scanning) {
      await state.scanner.stop();
      await state.scanner.clear();
    }
  } catch (_) {
    // ignore cleanup errors
  }
  state.scanning = false;
  els.toggleScannerBtn.textContent = 'Open QR Scanner';
  els.scannerWrap.classList.add('hidden');
}

async function callApi(action, payload = {}) {
  if (!CONFIG.webAppUrl || CONFIG.webAppUrl.includes('PASTE_')) {
    throw new Error('Add your Google Apps Script web app URL inside app.js first.');
  }

  const res = await fetch(CONFIG.webAppUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ action, ...payload }),
  });

  const text = await res.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch {
    throw new Error(text || 'Bad backend response.');
  }

  return data;
}

function setStatus(el, message, good) {
  el.textContent = message;
  el.style.borderColor = good ? '#cceedd' : '#cfe3fb';
  el.style.background = good ? '#f2fcf7' : '#f7fbff';
}

function toDateInput(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
