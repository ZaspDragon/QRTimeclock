import { firebaseConfig, appSettings } from './firebase-config.js';
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  getAuth,
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  sendPasswordResetEmail
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  getFirestore,
  collection,
  doc,
  getDoc,
  setDoc,
  addDoc,
  updateDoc,
  deleteDoc,
  serverTimestamp,
  query,
  where,
  orderBy,
  limit,
  onSnapshot,
  Timestamp
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const ACTIONS = ['clock_in', 'start_lunch', 'end_lunch', 'clock_out'];

const state = {
  me: null,
  profile: null,
  unsubscribers: [],
  selectedWeekStart: getMondayDate(new Date()),
  workerUnsub: null,
  allPunchRows: [],
  selectedWeekPunchRows: [],
  selectedWeekTimesheetDocs: {},
  userRows: [],
  auditRows: []
};

const els = {
  workerView: document.getElementById('workerView'),
  authView: document.getElementById('authView'),
  managerView: document.getElementById('managerView'),

  managerPortalBtn: document.getElementById('managerPortalBtn'),
  backToWorkerBtn: document.getElementById('backToWorkerBtn'),

  workerNameInput: document.getElementById('workerNameInput'),
  workerNameValue: document.getElementById('workerNameValue'),
  workerLastActionValue: document.getElementById('workerLastActionValue'),
  workerLastPunchValue: document.getElementById('workerLastPunchValue'),
  workerStatusValue: document.getElementById('workerStatusValue'),
  workerStatusMessage: document.getElementById('workerStatusMessage'),
  workerHistoryBody: document.getElementById('workerHistoryBody'),

  authCard: document.getElementById('authCard'),
  sessionChip: document.getElementById('sessionChip'),
  sessionName: document.getElementById('sessionName'),
  sessionRole: document.getElementById('sessionRole'),
  signOutBtn: document.getElementById('signOutBtn'),

  loginForm: document.getElementById('loginForm'),
  emailInput: document.getElementById('emailInput'),
  passwordInput: document.getElementById('passwordInput'),
  resetPasswordBtn: document.getElementById('resetPasswordBtn'),

  tabBar: document.getElementById('tabBar'),
  adminTabBtn: document.getElementById('adminTabBtn'),

  livePunchBody: document.getElementById('livePunchBody'),
  activeNowList: document.getElementById('activeNowList'),
  liveRangeFilter: document.getElementById('liveRangeFilter'),
  liveNameFilter: document.getElementById('liveNameFilter'),
  liveFlagFilter: document.getElementById('liveFlagFilter'),

  weekPicker: document.getElementById('weekPicker'),
  timesheetBody: document.getElementById('timesheetBody'),
  timesheetNameFilter: document.getElementById('timesheetNameFilter'),
  timesheetStatusFilter: document.getElementById('timesheetStatusFilter'),
  unsignedOnlyFilter: document.getElementById('unsignedOnlyFilter'),

  manualPunchForm: document.getElementById('manualPunchForm'),
  manualPunchNameInput: document.getElementById('manualPunchNameInput'),
  manualPunchActionInput: document.getElementById('manualPunchActionInput'),
  manualPunchDateInput: document.getElementById('manualPunchDateInput'),
  manualPunchTimeInput: document.getElementById('manualPunchTimeInput'),
  manualPunchReasonInput: document.getElementById('manualPunchReasonInput'),

  editFilterNameInput: document.getElementById('editFilterNameInput'),
  editSourceFilter: document.getElementById('editSourceFilter'),
  correctedOnlyFilter: document.getElementById('correctedOnlyFilter'),
  editPunchesBody: document.getElementById('editPunchesBody'),

  exportWorkerSelect: document.getElementById('exportWorkerSelect'),
  workerSummaryPreview: document.getElementById('workerSummaryPreview'),
  printTimesheetBtn: document.getElementById('printTimesheetBtn'),
  exportPayrollCsvBtn: document.getElementById('exportPayrollCsvBtn'),
  exportExceptionsCsvBtn: document.getElementById('exportExceptionsCsvBtn'),
  exportMissedPunchCsvBtn: document.getElementById('exportMissedPunchCsvBtn'),

  userProfileForm: document.getElementById('userProfileForm'),
  userUidInput: document.getElementById('userUidInput'),
  userNameInput: document.getElementById('userNameInput'),
  userEmailInput: document.getElementById('userEmailInput'),
  userRoleInput: document.getElementById('userRoleInput'),
  userBranchInput: document.getElementById('userBranchInput'),
  userActiveInput: document.getElementById('userActiveInput'),
  showInactiveUsersFilter: document.getElementById('showInactiveUsersFilter'),
  userListBody: document.getElementById('userListBody'),

  auditNameFilter: document.getElementById('auditNameFilter'),
  auditActionFilter: document.getElementById('auditActionFilter'),
  auditBody: document.getElementById('auditBody'),

  toast: document.getElementById('toast')
};

init();

function init() {
  wireEvents();

  const storedWorkerName = localStorage.getItem('workerPunchName') || '';
  if (storedWorkerName) {
    const pretty = prettifyHumanName(storedWorkerName);
    els.workerNameInput.value = pretty;
    els.workerNameValue.textContent = pretty;
    attachWorkerLiveView(pretty);
  }

  if (els.weekPicker) {
    els.weekPicker.value = formatDateInput(state.selectedWeekStart);
  }

  onAuthStateChanged(auth, async (user) => {
    clearAllListeners();

    if (!user) {
      state.me = null;
      state.profile = null;
      showLoggedOut();
      return;
    }

    try {
      state.me = user;
      const profileSnap = await getDoc(doc(db, 'users', user.uid));

      if (!profileSnap.exists()) {
        await signOut(auth);
        toast('No Firestore user profile exists for this account.', true);
        return;
      }

      const profile = profileSnap.data();
      if (profile.active === false) {
        await signOut(auth);
        toast('This user profile is inactive.', true);
        return;
      }

      state.profile = profile;
      showLoggedIn();
      attachManagerData();
    } catch (error) {
      console.error(error);
      toast(error.message || 'Sign-in setup failed.', true);
    }
  });
}

function wireEvents() {
  document.querySelectorAll('.worker-action-btn').forEach((btn) => {
    btn.addEventListener('click', () => handleWorkerPunch(btn.dataset.action));
  });

  els.workerNameInput?.addEventListener('input', () => {
    const value = prettifyHumanName(els.workerNameInput.value.trim());
    els.workerNameValue.textContent = value || '-';

    if (value) {
      localStorage.setItem('workerPunchName', value);
      attachWorkerLiveView(value);
    }
  });

  els.managerPortalBtn?.addEventListener('click', () => showManagerLogin());
  els.backToWorkerBtn?.addEventListener('click', () => showWorkerView());

  els.loginForm?.addEventListener('submit', handleLogin);
  els.resetPasswordBtn?.addEventListener('click', handlePasswordReset);
  els.signOutBtn?.addEventListener('click', async () => signOut(auth));

  els.tabBar?.addEventListener('click', (event) => {
    const btn = event.target.closest('.tab');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });

  els.weekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = new Date(`${els.weekPicker.value}T00:00:00`);
    refreshWeekListeners();
  });

  els.liveRangeFilter?.addEventListener('change', () => renderLivePunches(state.allPunchRows));
  els.liveNameFilter?.addEventListener('input', () => renderLivePunches(state.allPunchRows));
  els.liveFlagFilter?.addEventListener('change', () => renderLivePunches(state.allPunchRows));

  els.timesheetNameFilter?.addEventListener('input', renderDerivedTimesheets);
  els.timesheetStatusFilter?.addEventListener('change', renderDerivedTimesheets);
  els.unsignedOnlyFilter?.addEventListener('change', renderDerivedTimesheets);

  els.manualPunchForm?.addEventListener('submit', handleManualPunchSubmit);

  els.editFilterNameInput?.addEventListener('input', () => renderPunchEditor(state.selectedWeekPunchRows));
  els.editSourceFilter?.addEventListener('change', () => renderPunchEditor(state.selectedWeekPunchRows));
  els.correctedOnlyFilter?.addEventListener('change', () => renderPunchEditor(state.selectedWeekPunchRows));

  els.exportWorkerSelect?.addEventListener('change', renderWorkerSummaryPreview);
  els.printTimesheetBtn?.addEventListener('click', printSelectedTimesheet);
  els.exportPayrollCsvBtn?.addEventListener('click', exportPayrollCsv);
  els.exportExceptionsCsvBtn?.addEventListener('click', exportExceptionsCsv);
  els.exportMissedPunchCsvBtn?.addEventListener('click', exportMissedPunchCsv);

  els.userProfileForm?.addEventListener('submit', handleSaveProfile);
  els.showInactiveUsersFilter?.addEventListener('change', renderUsers);

  els.auditNameFilter?.addEventListener('input', renderAuditLog);
  els.auditActionFilter?.addEventListener('change', renderAuditLog);
}

function showWorkerView() {
  els.workerView.classList.remove('hidden');
  els.authView.classList.add('hidden');
  els.managerView.classList.add('hidden');
}

function showManagerLogin() {
  els.workerView.classList.add('hidden');
  els.authView.classList.remove('hidden');
  els.managerView.classList.add('hidden');
}

function showLoggedOut() {
  els.sessionChip?.classList.add('hidden');
  showWorkerView();
}

function showLoggedIn() {
  els.workerView.classList.add('hidden');
  els.authView.classList.add('hidden');
  els.managerView.classList.remove('hidden');
  els.sessionChip?.classList.remove('hidden');

  els.sessionName.textContent = state.profile?.name || state.me?.email || 'Signed in';
  els.sessionRole.textContent = state.profile?.role || 'manager';

  els.adminTabBtn?.classList.toggle('hidden', !isAdmin());
  switchTab('managerTab');
}

function switchTab(tabId) {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('active', tab.dataset.tab === tabId);
  });

  document.querySelectorAll('.tab-panel').forEach((panel) => {
    panel.classList.toggle('hidden', panel.id !== tabId);
  });
}

async function handleLogin(event) {
  event.preventDefault();

  try {
    await signInWithEmailAndPassword(
      auth,
      els.emailInput.value.trim(),
      els.passwordInput.value
    );
    els.passwordInput.value = '';
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not sign in.', true);
  }
}

async function handlePasswordReset() {
  const email = els.emailInput.value.trim();
  if (!email) {
    toast('Enter the email first, then tap reset.', true);
    return;
  }

  try {
    await sendPasswordResetEmail(auth, email);
    toast('Password reset email sent.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not send reset email.', true);
  }
}

async function handleWorkerPunch(action) {
  const rawName = els.workerNameInput.value.trim();

  if (!rawName) {
    toast('Enter your name first.', true);
    return;
  }

  const name = prettifyHumanName(rawName);
  const nameKey = normalizeName(name);

  if (nameKey.length < 2) {
    toast('Enter a valid name.', true);
    return;
  }

  const now = new Date();
  const nowMs = Date.now();
  const dateKey = formatDateKey(now);
  const weekKey = formatDateKey(getMondayDate(now));

  try {
    await addDoc(collection(db, 'punches'), {
      name,
      nameKey,
      action,
      timestamp: serverTimestamp(),
      timestampMs: nowMs,
      dateKey,
      weekKey,
      source: 'public_qr',
      corrected: false,
      createdAt: serverTimestamp()
    });

    localStorage.setItem('workerPunchName', name);
    els.workerNameInput.value = name;
    els.workerNameValue.textContent = name;
    els.workerLastActionValue.textContent = prettyAction(action);
    els.workerLastPunchValue.textContent = formatDateTime(nowMs);
    els.workerStatusValue.textContent = statusLabelForAction(action);
    els.workerStatusMessage.textContent = `${prettyAction(action)} recorded at ${formatDateTime(nowMs)}.`;

    attachWorkerLiveView(name);
    toast(`${prettyAction(action)} saved.`);
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save punch.', true);
  }
}

function attachWorkerLiveView(name) {
  if (state.workerUnsub) {
    try { state.workerUnsub(); } catch (_) {}
    state.workerUnsub = null;
  }

  const nameKey = normalizeName(name);
  if (!nameKey) return;

  const todayKey = formatDateKey(new Date());

  const q = query(
    collection(db, 'punches'),
    where('nameKey', '==', nameKey),
    where('dateKey', '==', todayKey),
    orderBy('timestampMs', 'desc'),
    limit(20)
  );

  state.workerUnsub = onSnapshot(q, (snap) => {
    const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

    if (!rows.length) {
      els.workerLastActionValue.textContent = '-';
      els.workerLastPunchValue.textContent = '-';
      els.workerStatusValue.textContent = 'Ready';
      els.workerStatusMessage.textContent = 'Enter your name and punch.';
      els.workerHistoryBody.innerHTML = '<tr><td colspan="2">No punches yet.</td></tr>';
      return;
    }

    const last = rows[0];

    els.workerNameValue.textContent = last.name || name;
    els.workerLastActionValue.textContent = prettyAction(last.action);
    els.workerLastPunchValue.textContent = formatDateTime(last.timestampMs);
    els.workerStatusValue.textContent = statusLabelForAction(last.action);
    els.workerStatusMessage.textContent = `${prettyAction(last.action)} recorded at ${formatDateTime(last.timestampMs)}.`;

    els.workerHistoryBody.innerHTML = rows.map((row) => `
      <tr>
        <td>${formatDateTime(row.timestampMs)}</td>
        <td>${prettyAction(row.action)}</td>
      </tr>
    `).join('');
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load worker punches.', true);
  });
}

function attachManagerData() {
  attachLivePunchesListener();
  attachWeekPunchesListener();
  attachTimesheetsListener();
  attachUsersListener();
  attachAuditListener();
}

function refreshWeekListeners() {
  clearWeekDependentListeners();

  attachWeekPunchesListener();
  attachTimesheetsListener();
}

function attachLivePunchesListener() {
  const q = query(collection(db, 'punches'), orderBy('timestampMs', 'desc'), limit(300));

  const unsub = onSnapshot(q, (snap) => {
    state.allPunchRows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    renderLivePunches(state.allPunchRows);
    renderActiveNow(state.allPunchRows);
  }, (error) => {
    console.error(error);
    toast(error.message || 'Live punch feed failed.', true);
  });

  state.unsubscribers.push({ key: 'live', unsub });
}

function attachWeekPunchesListener() {
  const weekKey = formatDateKey(state.selectedWeekStart);

  const q = query(
    collection(db, 'punches'),
    where('weekKey', '==', weekKey),
    orderBy('timestampMs', 'asc')
  );

  const unsub = onSnapshot(q, (snap) => {
    state.selectedWeekPunchRows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    renderDerivedTimesheets();
    renderPunchEditor(state.selectedWeekPunchRows);
    populateExportWorkerSelect();
    renderWorkerSummaryPreview();
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load weekly punches.', true);
  });

  state.unsubscribers.push({ key: 'week_punches', unsub });
}

function attachTimesheetsListener() {
  const weekKey = formatDateKey(state.selectedWeekStart);

  const q = query(
    collection(db, 'timesheets'),
    where('weekKey', '==', weekKey)
  );

  const unsub = onSnapshot(q, (snap) => {
    const map = {};
    snap.docs.forEach((d) => {
      map[d.id] = { id: d.id, ...d.data() };
    });
    state.selectedWeekTimesheetDocs = map;
    renderDerivedTimesheets();
    renderWorkerSummaryPreview();
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load timesheets.', true);
  });

  state.unsubscribers.push({ key: 'timesheets', unsub });
}

function attachUsersListener() {
  if (!isAdmin()) return;

  const q = query(collection(db, 'users'), orderBy('name', 'asc'));
  const unsub = onSnapshot(q, (snap) => {
    state.userRows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    renderUsers();
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load users.', true);
  });

  state.unsubscribers.push({ key: 'users', unsub });
}

function attachAuditListener() {
  const q = query(collection(db, 'audit_logs'), orderBy('timestampMs', 'desc'), limit(250));

  const unsub = onSnapshot(q, (snap) => {
    state.auditRows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    renderAuditLog();
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load audit log.', true);
  });

  state.unsubscribers.push({ key: 'audit', unsub });
}

function renderLivePunches(rows) {
  const filtered = getFilteredLivePunchRows(rows);

  if (!filtered.length) {
    els.livePunchBody.innerHTML = '<tr><td colspan="5">No live data yet.</td></tr>';
    return;
  }

  els.livePunchBody.innerHTML = filtered.map((row) => {
    const flags = getPunchFlags(row, rows);
    return `
      <tr>
        <td>${formatDateTime(row.timestampMs)}</td>
        <td>${escapeHtml(row.name || '-')}</td>
        <td><span class="action-pill action-${escapeHtml(row.action || 'unknown')}">${prettyAction(row.action)}</span></td>
        <td>${escapeHtml(row.source || '-')}</td>
        <td>${formatFlags(flags)}</td>
      </tr>
    `;
  }).join('');
}

function getFilteredLivePunchRows(rows) {
  const range = els.liveRangeFilter?.value || 'today';
  const nameFilter = String(els.liveNameFilter?.value || '').trim().toLowerCase();
  const flagFilter = els.liveFlagFilter?.value || '';
  const todayKey = formatDateKey(new Date());
  const currentWeekKey = formatDateKey(getMondayDate(new Date()));

  return rows.filter((row) => {
    const matchesName = !nameFilter || String(row.name || '').toLowerCase().includes(nameFilter);

    let matchesRange = true;
    if (range === 'today') matchesRange = row.dateKey === todayKey;
    if (range === 'week') matchesRange = row.weekKey === currentWeekKey;

    const flags = getPunchFlags(row, rows);
    let matchesFlag = true;
    if (flagFilter === 'duplicates') matchesFlag = flags.includes('Duplicate');
    if (flagFilter === 'open') matchesFlag = flags.includes('Open Shift');
    if (flagFilter === 'late_lunch') matchesFlag = flags.includes('Late Lunch');
    if (flagFilter === 'no_lunch') matchesFlag = flags.includes('No Lunch');

    return matchesName && matchesRange && matchesFlag;
  });
}

function renderActiveNow(rows) {
  const latestByName = new Map();

  rows.forEach((row) => {
    const key = row.nameKey || normalizeName(row.name || '');
    if (!key) return;
    if (!latestByName.has(key)) latestByName.set(key, row);
  });

  const active = [...latestByName.values()].filter((row) =>
    row.action === 'clock_in' || row.action === 'end_lunch'
  );

  if (!active.length) {
    els.activeNowList.innerHTML = '<div class="empty-state">Nobody is currently clocked in.</div>';
    return;
  }

  els.activeNowList.innerHTML = active
    .sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')))
    .map((row) => {
      const flags = getPunchFlags(row, rows);
      return `
        <div class="person-row">
          <div class="person-meta">
            <strong>${escapeHtml(row.name || '-')}</strong>
            <span>${prettyAction(row.action)}</span>
            ${flags.length ? `<div class="tiny">${formatFlags(flags)}</div>` : ''}
          </div>
          <div class="pill">${formatTime(row.timestampMs)}</div>
        </div>
      `;
    })
    .join('');
}

function renderDerivedTimesheets() {
  const rows = getDerivedTimesheetRows();
  const nameFilter = String(els.timesheetNameFilter?.value || '').trim().toLowerCase();
  const statusFilter = els.timesheetStatusFilter?.value || '';
  const unsignedOnly = els.unsignedOnlyFilter?.value === 'yes';

  const filtered = rows.filter((row) => {
    const matchesName = !nameFilter || String(row.name || '').toLowerCase().includes(nameFilter);

    const hasExceptions = row.flags.length > 0;
    let matchesStatus = true;
    if (statusFilter === 'open') matchesStatus = row.status === 'open';
    if (statusFilter === 'signed') matchesStatus = row.status === 'signed';
    if (statusFilter === 'exception') matchesStatus = hasExceptions;

    const matchesUnsigned = !unsignedOnly || row.status !== 'signed';

    return matchesName && matchesStatus && matchesUnsigned;
  });

  if (!filtered.length) {
    els.timesheetBody.innerHTML = '<tr><td colspan="7">No timesheets yet.</td></tr>';
    return;
  }

  els.timesheetBody.innerHTML = filtered.map((row) => {
    const signedAt = row.managerSignedAt?.seconds
      ? formatDateTime(row.managerSignedAt.seconds * 1000)
      : '-';

    return `
      <tr>
        <td>${escapeHtml(row.name)}</td>
        <td>${escapeHtml(row.weekKey)}</td>
        <td>${Number(row.weeklyHours || 0).toFixed(2)} (${row.daysWorked} day${row.daysWorked === 1 ? '' : 's'})</td>
        <td>${escapeHtml(row.status)}</td>
        <td>${formatFlags(row.flags)}</td>
        <td>${escapeHtml(row.managerSignedBy || '-')}${signedAt !== '-' ? `<br><span class="tiny">${signedAt}</span>` : ''}</td>
        <td>
          ${row.status === 'signed'
            ? `<button class="ghost-btn reopen-btn" data-id="${row.id}" type="button">Reopen</button>`
            : `<button class="primary-btn sign-btn" data-id="${row.id}" type="button">Sign</button>`
          }
        </td>
      </tr>
    `;
  }).join('');

  els.timesheetBody.querySelectorAll('.sign-btn').forEach((btn) => {
    btn.addEventListener('click', () => signTimesheet(btn.dataset.id));
  });

  els.timesheetBody.querySelectorAll('.reopen-btn').forEach((btn) => {
    btn.addEventListener('click', () => reopenTimesheet(btn.dataset.id));
  });
}

function getDerivedTimesheetRows() {
  const weekKey = formatDateKey(state.selectedWeekStart);
  const grouped = new Map();

  state.selectedWeekPunchRows.forEach((p) => {
    const key = p.nameKey || normalizeName(p.name || '');
    if (!key) return;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(p);
  });

  const rows = [];

  grouped.forEach((personPunches, nameKey) => {
    const displayName = personPunches[0]?.name || nameKey;
    const totals = buildWeekTotals(personPunches);
    const timesheetId = `${weekKey}_${nameKey}`;
    const saved = state.selectedWeekTimesheetDocs[timesheetId] || null;

    rows.push({
      id: timesheetId,
      name: displayName,
      nameKey,
      weekKey,
      weeklyHours: totals.weeklyHours,
      daysWorked: totals.daysWorked,
      dailyTotals: totals.dailyTotals,
      lastPunchAction: totals.lastAction,
      lastPunchAtMs: totals.lastPunchAtMs,
      flags: totals.flags,
      status: saved?.status || 'open',
      managerSignedBy: saved?.managerSignedBy || '',
      managerSignedAt: saved?.managerSignedAt || null,
      approvalNote: saved?.approvalNote || ''
    });
  });

  rows.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return rows;
}

function buildWeekTotals(punches) {
  const sorted = [...punches].sort((a, b) => (a.timestampMs || 0) - (b.timestampMs || 0));
  const byDay = {};
  let currentIn = null;
  let weeklyMinutes = 0;
  let lastAction = '-';
  let lastPunchAtMs = 0;
  const flags = [];

  sorted.forEach((punch) => {
    const timeMs = punch.timestampMs || 0;
    const dateKey = punch.dateKey || formatDateKey(new Date(timeMs));

    if (!byDay[dateKey]) {
      byDay[dateKey] = {
        clock_in: '',
        start_lunch: '',
        end_lunch: '',
        clock_out: '',
        minutes: 0,
        hadLunchStart: false,
        hadLunchEnd: false,
        actionList: []
      };
    }

    const day = byDay[dateKey];
    day.actionList.push(punch.action);

    if (day.actionList.length >= 2) {
      const prevAction = day.actionList[day.actionList.length - 2];
      if (prevAction === punch.action) {
        pushUnique(flags, 'Duplicate');
      }
    }

    lastAction = punch.action;
    lastPunchAtMs = Math.max(lastPunchAtMs, timeMs);

    if (punch.action === 'clock_in') {
      day.clock_in = formatTime(timeMs);
      currentIn = timeMs;
    }

    if (punch.action === 'start_lunch') {
      day.start_lunch = formatTime(timeMs);
      day.hadLunchStart = true;
      if (currentIn) {
        const diff = Math.max(0, Math.round((timeMs - currentIn) / 60000));
        weeklyMinutes += diff;
        day.minutes += diff;
        currentIn = null;
      }
    }

    if (punch.action === 'end_lunch') {
      day.end_lunch = formatTime(timeMs);
      day.hadLunchEnd = true;
      currentIn = timeMs;
    }

    if (punch.action === 'clock_out') {
      day.clock_out = formatTime(timeMs);
      if (currentIn) {
        const diff = Math.max(0, Math.round((timeMs - currentIn) / 60000));
        weeklyMinutes += diff;
        day.minutes += diff;
        currentIn = null;
      }
    }
  });

  if (currentIn) {
    pushUnique(flags, 'Open Shift');
  }

  const dailyTotals = Object.fromEntries(
    Object.entries(byDay).map(([dateKey, value]) => {
      const dayFlags = [];

      if (value.clock_in && !value.clock_out) dayFlags.push('Open Shift');
      if (value.start_lunch && !value.end_lunch) dayFlags.push('Missing Lunch Return');
      if (!value.start_lunch && value.clock_in && value.clock_out && value.minutes >= 360) dayFlags.push('No Lunch');

      if (value.start_lunch && value.end_lunch) {
        const lunchMinutes = timeRangeMinutesFromStrings(value.start_lunch, value.end_lunch);
        if (lunchMinutes > 45) dayFlags.push('Late Lunch');
      }

      dayFlags.forEach((f) => pushUnique(flags, f));

      return [dateKey, {
        clock_in: value.clock_in,
        start_lunch: value.start_lunch,
        end_lunch: value.end_lunch,
        clock_out: value.clock_out,
        hours: Number((value.minutes / 60).toFixed(2)),
        flags: dayFlags
      }];
    })
  );

  return {
    dailyTotals,
    weeklyHours: Number((weeklyMinutes / 60).toFixed(2)),
    daysWorked: Object.keys(dailyTotals).length,
    lastAction,
    lastPunchAtMs,
    flags
  };
}

async function signTimesheet(timesheetId) {
  const row = getDerivedTimesheetRows().find((r) => r.id === timesheetId);
  if (!row) {
    toast('Could not find that weekly record.', true);
    return;
  }

  const approvalNote = prompt('Approval note (optional):', row.approvalNote || '') ?? '';

  try {
    await setDoc(doc(db, 'timesheets', timesheetId), {
      name: row.name,
      nameKey: row.nameKey,
      weekKey: row.weekKey,
      dailyTotals: row.dailyTotals,
      weeklyHours: row.weeklyHours,
      daysWorked: row.daysWorked,
      status: 'signed',
      approvalNote,
      flags: row.flags,
      managerSignedBy: state.profile?.name || state.me?.email || 'Manager',
      managerSignedAt: Timestamp.fromDate(new Date()),
      updatedAt: serverTimestamp(),
      lastPunchAction: row.lastPunchAction,
      lastPunchAtMs: row.lastPunchAtMs
    }, { merge: true });

    await addAuditLog({
      type: 'timesheet_sign',
      workerName: row.name,
      workerKey: row.nameKey,
      reason: approvalNote || 'Signed timesheet',
      oldValue: { status: 'open' },
      newValue: { status: 'signed', flags: row.flags, approvalNote }
    });

    toast('Timesheet signed.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not sign timesheet.', true);
  }
}

async function reopenTimesheet(timesheetId) {
  const row = getDerivedTimesheetRows().find((r) => r.id === timesheetId);
  if (!row) {
    toast('Could not find that weekly record.', true);
    return;
  }

  const reason = prompt('Reason for reopening this timesheet:', '');
  if (reason === null) return;
  if (!reason.trim()) {
    toast('Reason is required to reopen a timesheet.', true);
    return;
  }

  try {
    await setDoc(doc(db, 'timesheets', timesheetId), {
      name: row.name,
      nameKey: row.nameKey,
      weekKey: row.weekKey,
      dailyTotals: row.dailyTotals,
      weeklyHours: row.weeklyHours,
      daysWorked: row.daysWorked,
      status: 'open',
      approvalNote: '',
      flags: row.flags,
      managerSignedBy: '',
      managerSignedAt: null,
      updatedAt: serverTimestamp(),
      lastPunchAction: row.lastPunchAction,
      lastPunchAtMs: row.lastPunchAtMs
    }, { merge: true });

    await addAuditLog({
      type: 'timesheet_reopen',
      workerName: row.name,
      workerKey: row.nameKey,
      reason,
      oldValue: { status: 'signed' },
      newValue: { status: 'open' }
    });

    toast('Timesheet reopened.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not reopen timesheet.', true);
  }
}

async function handleManualPunchSubmit(event) {
  event.preventDefault();

  const name = prettifyHumanName(els.manualPunchNameInput.value.trim());
  const nameKey = normalizeName(name);
  const action = els.manualPunchActionInput.value;
  const date = els.manualPunchDateInput.value;
  const time = els.manualPunchTimeInput.value;
  const reason = els.manualPunchReasonInput.value.trim();

  if (!name || nameKey.length < 2) {
    toast('Enter a valid worker name.', true);
    return;
  }

  if (!ACTIONS.includes(action)) {
    toast('Invalid action.', true);
    return;
  }

  if (!date || !time) {
    toast('Date and time are required.', true);
    return;
  }

  if (!reason) {
    toast('Reason for correction is required.', true);
    return;
  }

  const ms = parseLocalEditString(`${date} ${time}`);
  if (!ms) {
    toast('Invalid date/time.', true);
    return;
  }

  const punchDate = new Date(ms);

  try {
    const ref = await addDoc(collection(db, 'punches'), {
      name,
      nameKey,
      action,
      timestamp: serverTimestamp(),
      timestampMs: ms,
      dateKey: formatDateKey(punchDate),
      weekKey: formatDateKey(getMondayDate(punchDate)),
      source: 'manager_manual',
      corrected: true,
      correctionReason: reason,
      createdAt: serverTimestamp(),
      createdBy: state.profile?.name || state.me?.email || 'Manager'
    });

    await addAuditLog({
      type: 'manual_add',
      workerName: name,
      workerKey: nameKey,
      reason,
      oldValue: null,
      newValue: {
        id: ref.id,
        action,
        timestampMs: ms,
        source: 'manager_manual'
      }
    });

    els.manualPunchForm.reset();
    els.manualPunchDateInput.value = formatDateInput(new Date());
    els.manualPunchTimeInput.value = currentTimeInputValue();
    toast('Manual punch added.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not add manual punch.', true);
  }
}

function renderPunchEditor(rows) {
  const nameFilter = String(els.editFilterNameInput?.value || '').trim().toLowerCase();
  const sourceFilter = els.editSourceFilter?.value || '';
  const correctedOnly = els.correctedOnlyFilter?.value === 'yes';

  const filtered = rows.filter((row) => {
    const matchesName = !nameFilter || String(row.name || '').toLowerCase().includes(nameFilter);
    const matchesSource = !sourceFilter || row.source === sourceFilter;
    const matchesCorrected = !correctedOnly || row.corrected === true;
    return matchesName && matchesSource && matchesCorrected;
  });

  if (!filtered.length) {
    els.editPunchesBody.innerHTML = '<tr><td colspan="7">No punches yet.</td></tr>';
    return;
  }

  els.editPunchesBody.innerHTML = filtered.map((row) => {
    const flags = getPunchFlags(row, state.selectedWeekPunchRows);
    return `
      <tr>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${prettyAction(row.action)}</td>
        <td>${formatDateOnly(row.timestampMs)}</td>
        <td>${formatTime(row.timestampMs)}</td>
        <td>${escapeHtml(row.source || '-')}</td>
        <td>${formatFlags(flags)}</td>
        <td>
          <button class="secondary-btn punch-edit-btn" data-id="${row.id}" type="button">Edit</button>
          <button class="danger-btn punch-delete-btn" data-id="${row.id}" type="button">Delete</button>
        </td>
      </tr>
    `;
  }).join('');

  els.editPunchesBody.querySelectorAll('.punch-edit-btn').forEach((btn) => {
    btn.addEventListener('click', () => editPunch(btn.dataset.id));
  });

  els.editPunchesBody.querySelectorAll('.punch-delete-btn').forEach((btn) => {
    btn.addEventListener('click', () => deletePunchRecord(btn.dataset.id));
  });
}

async function editPunch(punchId) {
  const row = state.selectedWeekPunchRows.find((r) => r.id === punchId);
  if (!row) {
    toast('Punch not found.', true);
    return;
  }

  const newName = prompt('Edit worker name:', row.name || '');
  if (newName === null) return;

  const newAction = prompt('Edit action (clock_in, start_lunch, end_lunch, clock_out):', row.action || 'clock_in');
  if (newAction === null) return;

  const newDateTime = prompt('Edit date/time (YYYY-MM-DD HH:MM):', toLocalEditString(row.timestampMs));
  if (newDateTime === null) return;

  const reason = prompt('Reason for this edit:', row.correctionReason || '');
  if (reason === null) return;
  if (!reason.trim()) {
    toast('Reason is required for edits.', true);
    return;
  }

  const prettyName = prettifyHumanName(newName);
  const nameKey = normalizeName(prettyName);
  const action = String(newAction).trim();
  const parsedMs = parseLocalEditString(newDateTime);

  if (!prettyName || nameKey.length < 2) {
    toast('Invalid name.', true);
    return;
  }

  if (!ACTIONS.includes(action)) {
    toast('Invalid action.', true);
    return;
  }

  if (!parsedMs) {
    toast('Invalid date/time. Use YYYY-MM-DD HH:MM', true);
    return;
  }

  const date = new Date(parsedMs);
  const oldValue = {
    name: row.name,
    action: row.action,
    timestampMs: row.timestampMs,
    source: row.source,
    correctionReason: row.correctionReason || ''
  };

  const newValue = {
    name: prettyName,
    action,
    timestampMs: parsedMs,
    source: 'manager_edit',
    correctionReason: reason
  };

  try {
    await updateDoc(doc(db, 'punches', punchId), {
      name: prettyName,
      nameKey,
      action,
      timestampMs: parsedMs,
      dateKey: formatDateKey(date),
      weekKey: formatDateKey(getMondayDate(date)),
      source: 'manager_edit',
      corrected: true,
      correctionReason: reason,
      editedAt: serverTimestamp(),
      editedBy: state.profile?.name || state.me?.email || 'Manager'
    });

    await addAuditLog({
      type: 'edit',
      workerName: prettyName,
      workerKey: nameKey,
      reason,
      oldValue,
      newValue
    });

    toast('Punch updated.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not update punch.', true);
  }
}

async function deletePunchRecord(punchId) {
  const row = state.selectedWeekPunchRows.find((r) => r.id === punchId);
  if (!row) {
    toast('Punch not found.', true);
    return;
  }

  const reason = prompt('Reason for deleting this punch:', '');
  if (reason === null) return;
  if (!reason.trim()) {
    toast('Reason is required for deletes.', true);
    return;
  }

  try {
    await deleteDoc(doc(db, 'punches', punchId));

    await addAuditLog({
      type: 'delete',
      workerName: row.name,
      workerKey: row.nameKey || normalizeName(row.name || ''),
      reason,
      oldValue: {
        id: row.id,
        name: row.name,
        action: row.action,
        timestampMs: row.timestampMs,
        source: row.source
      },
      newValue: null
    });

    toast('Punch deleted.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not delete punch.', true);
  }
}

async function addAuditLog({ type, workerName, workerKey, reason, oldValue, newValue }) {
  try {
    await addDoc(collection(db, 'audit_logs'), {
      type,
      workerName: workerName || '',
      workerKey: workerKey || '',
      reason: reason || '',
      oldValue: oldValue ?? null,
      newValue: newValue ?? null,
      actorName: state.profile?.name || state.me?.email || 'Manager',
      actorUid: state.me?.uid || '',
      timestamp: serverTimestamp(),
      timestampMs: Date.now()
    });
  } catch (error) {
    console.error('Audit log failed:', error);
  }
}

function renderUsers() {
  if (!isAdmin()) return;

  const showInactive = els.showInactiveUsersFilter?.value === 'yes';
  const rows = state.userRows.filter((row) => showInactive || row.active !== false);

  if (!rows.length) {
    els.userListBody.innerHTML = '<tr><td colspan="5">No users yet.</td></tr>';
    return;
  }

  els.userListBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.name || '-')}</td>
      <td>${escapeHtml(row.email || '-')}</td>
      <td>${escapeHtml(row.role || '-')}</td>
      <td>${escapeHtml(row.branch || '-')}</td>
      <td>${row.active === false ? 'No' : 'Yes'}</td>
    </tr>
  `).join('');
}

async function handleSaveProfile(event) {
  event.preventDefault();

  if (!isAdmin()) {
    toast('Only admins can save user profiles.', true);
    return;
  }

  try {
    const uid = els.userUidInput.value.trim();
    await setDoc(doc(db, 'users', uid), {
      name: prettifyHumanName(els.userNameInput.value.trim()),
      email: els.userEmailInput.value.trim().toLowerCase(),
      role: els.userRoleInput.value,
      branch: els.userBranchInput.value.trim(),
      active: els.userActiveInput.value === 'true',
      updatedAt: serverTimestamp()
    }, { merge: true });

    toast('User profile saved.');
    els.userProfileForm.reset();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save profile.', true);
  }
}

function renderAuditLog() {
  const nameFilter = String(els.auditNameFilter?.value || '').trim().toLowerCase();
  const actionFilter = els.auditActionFilter?.value || '';

  const filtered = state.auditRows.filter((row) => {
    const matchesName = !nameFilter || String(row.workerName || '').toLowerCase().includes(nameFilter);
    const matchesAction = !actionFilter || row.type === actionFilter;
    return matchesName && matchesAction;
  });

  if (!filtered.length) {
    els.auditBody.innerHTML = '<tr><td colspan="7">No audit history yet.</td></tr>';
    return;
  }

  els.auditBody.innerHTML = filtered.map((row) => `
    <tr>
      <td>${formatDateTime(row.timestampMs)}</td>
      <td>${escapeHtml(row.actorName || '-')}</td>
      <td>${escapeHtml(row.type || '-')}</td>
      <td>${escapeHtml(row.workerName || '-')}</td>
      <td>${escapeHtml(row.reason || '-')}</td>
      <td><pre class="table-pre">${escapeHtml(stringifyShort(row.oldValue))}</pre></td>
      <td><pre class="table-pre">${escapeHtml(stringifyShort(row.newValue))}</pre></td>
    </tr>
  `).join('');
}

function populateExportWorkerSelect() {
  const rows = getDerivedTimesheetRows();
  const current = els.exportWorkerSelect.value;

  els.exportWorkerSelect.innerHTML = '<option value="">All workers</option>' +
    rows.map((row) => `<option value="${escapeHtml(row.nameKey)}">${escapeHtml(row.name)}</option>`).join('');

  if (rows.some((row) => row.nameKey === current)) {
    els.exportWorkerSelect.value = current;
  }
}

function renderWorkerSummaryPreview() {
  const selected = els.exportWorkerSelect.value;
  const rows = getDerivedTimesheetRows();

  if (!selected) {
    els.workerSummaryPreview.innerHTML = 'Select a worker to preview summary.';
    return;
  }

  const row = rows.find((r) => r.nameKey === selected);
  if (!row) {
    els.workerSummaryPreview.innerHTML = 'No summary found for that worker.';
    return;
  }

  const dailyHtml = Object.entries(row.dailyTotals || {}).sort(([a], [b]) => a.localeCompare(b))
    .map(([dateKey, day]) => `
      <tr>
        <td>${escapeHtml(dateKey)}</td>
        <td>${escapeHtml(day.clock_in || '-')}</td>
        <td>${escapeHtml(day.start_lunch || '-')}</td>
        <td>${escapeHtml(day.end_lunch || '-')}</td>
        <td>${escapeHtml(day.clock_out || '-')}</td>
        <td>${Number(day.hours || 0).toFixed(2)}</td>
        <td>${formatFlags(day.flags || [])}</td>
      </tr>
    `).join('');

  els.workerSummaryPreview.innerHTML = `
    <div class="stack gap16">
      <div>
        <strong>${escapeHtml(row.name)}</strong><br />
        Week: ${escapeHtml(row.weekKey)}<br />
        Hours: ${Number(row.weeklyHours || 0).toFixed(2)}<br />
        Status: ${escapeHtml(row.status)}<br />
        Flags: ${formatFlags(row.flags)}
      </div>

      <div class="mini-table-wrap">
        <table>
          <thead>
            <tr>
              <th>Date</th>
              <th>Clock In</th>
              <th>Lunch Out</th>
              <th>Lunch In</th>
              <th>Clock Out</th>
              <th>Hours</th>
              <th>Flags</th>
            </tr>
          </thead>
          <tbody>${dailyHtml || '<tr><td colspan="7">No daily rows.</td></tr>'}</tbody>
        </table>
      </div>
    </div>
  `;
}

function printSelectedTimesheet() {
  const selected = els.exportWorkerSelect.value;
  if (!selected) {
    toast('Select a worker first.', true);
    return;
  }

  const row = getDerivedTimesheetRows().find((r) => r.nameKey === selected);
  if (!row) {
    toast('No timesheet found for that worker.', true);
    return;
  }

  const signedAt = row.managerSignedAt?.seconds
    ? formatDateTime(row.managerSignedAt.seconds * 1000)
    : '-';

  const dailyRows = Object.entries(row.dailyTotals || {}).sort(([a], [b]) => a.localeCompare(b))
    .map(([dateKey, day]) => `
      <tr>
        <td>${escapeHtml(dateKey)}</td>
        <td>${escapeHtml(day.clock_in || '-')}</td>
        <td>${escapeHtml(day.start_lunch || '-')}</td>
        <td>${escapeHtml(day.end_lunch || '-')}</td>
        <td>${escapeHtml(day.clock_out || '-')}</td>
        <td>${Number(day.hours || 0).toFixed(2)}</td>
        <td>${formatFlags(day.flags || [])}</td>
      </tr>
    `).join('');

  const html = `
    <div style="font-family: Arial, sans-serif; color:#111; padding:24px;">
      <h1 style="margin:0 0 12px;">Weekly Timesheet</h1>
      <div style="margin-bottom:16px; line-height:1.7;">
        <div><strong>Worker:</strong> ${escapeHtml(row.name)}</div>
        <div><strong>Week Start:</strong> ${escapeHtml(row.weekKey)}</div>
        <div><strong>Total Hours:</strong> ${Number(row.weeklyHours || 0).toFixed(2)}</div>
        <div><strong>Status:</strong> ${escapeHtml(row.status)}</div>
        <div><strong>Flags:</strong> ${formatFlags(row.flags)}</div>
        <div><strong>Signed By:</strong> ${escapeHtml(row.managerSignedBy || '-')}</div>
        <div><strong>Signed At:</strong> ${escapeHtml(signedAt)}</div>
      </div>

      <table style="width:100%; border-collapse:collapse;">
        <thead>
          <tr>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Date</th>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Clock In</th>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Lunch Out</th>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Lunch In</th>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Clock Out</th>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Hours</th>
            <th style="border:1px solid #ccc;padding:8px;text-align:left;">Flags</th>
          </tr>
        </thead>
        <tbody>${dailyRows}</tbody>
      </table>
    </div>
  `;

  openPrintWindow('Weekly Timesheet', html);
}

function exportPayrollCsv() {
  const rows = getDerivedTimesheetRows();
  const csvRows = [
    ['Worker', 'Week Start', 'Hours', 'Days Worked', 'Status', 'Signed By']
  ];

  rows.forEach((row) => {
    csvRows.push([
      row.name,
      row.weekKey,
      Number(row.weeklyHours || 0).toFixed(2),
      row.daysWorked,
      row.status,
      row.managerSignedBy || ''
    ]);
  });

  downloadCsv(`payroll_${formatDateKey(state.selectedWeekStart)}.csv`, csvRows);
}

function exportExceptionsCsv() {
  const rows = getDerivedTimesheetRows().filter((row) => row.flags.length > 0);

  const csvRows = [
    ['Worker', 'Week Start', 'Hours', 'Flags']
  ];

  rows.forEach((row) => {
    csvRows.push([
      row.name,
      row.weekKey,
      Number(row.weeklyHours || 0).toFixed(2),
      row.flags.join('; ')
    ]);
  });

  downloadCsv(`exceptions_${formatDateKey(state.selectedWeekStart)}.csv`, csvRows);
}

function exportMissedPunchCsv() {
  const rows = getDerivedTimesheetRows().filter((row) =>
    row.flags.includes('Open Shift') ||
    row.flags.includes('Missing Lunch Return')
  );

  const csvRows = [
    ['Worker', 'Week Start', 'Flags']
  ];

  rows.forEach((row) => {
    csvRows.push([
      row.name,
      row.weekKey,
      row.flags.join('; ')
    ]);
  });

  downloadCsv(`missed_punches_${formatDateKey(state.selectedWeekStart)}.csv`, csvRows);
}

function downloadCsv(filename, rows) {
  const content = rows
    .map((row) => row.map(csvEscape).join(','))
    .join('\n');

  const blob = new Blob([content], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}

function csvEscape(value) {
  const str = String(value ?? '');
  if (str.includes(',') || str.includes('"') || str.includes('\n')) {
    return `"${str.replaceAll('"', '""')}"`;
  }
  return str;
}

function getPunchFlags(row, rows) {
  const flags = [];
  const sameWorkerSameDay = rows
    .filter((r) => (r.nameKey || normalizeName(r.name || '')) === (row.nameKey || normalizeName(row.name || '')))
    .filter((r) => r.dateKey === row.dateKey)
    .sort((a, b) => (a.timestampMs || 0) - (b.timestampMs || 0));

  const index = sameWorkerSameDay.findIndex((r) => r.id === row.id);

  if (index > 0 && sameWorkerSameDay[index - 1].action === row.action) {
    flags.push('Duplicate');
  }

  if (row.action === 'clock_in' || row.action === 'end_lunch') {
    const laterRows = sameWorkerSameDay.filter((r) => (r.timestampMs || 0) > (row.timestampMs || 0));
    const hasClosingAction = laterRows.some((r) =>
      r.action === 'start_lunch' || r.action === 'clock_out'
    );

    if (!hasClosingAction) {
      flags.push('Open Shift');
    }
  }

  if (row.action === 'start_lunch') {
    const laterRows = sameWorkerSameDay.filter((r) => (r.timestampMs || 0) > (row.timestampMs || 0));
    const lunchReturn = laterRows.find((r) => r.action === 'end_lunch');
    if (!lunchReturn) {
      flags.push('Missing Lunch Return');
    } else {
      const minutes = Math.round((lunchReturn.timestampMs - row.timestampMs) / 60000);
      if (minutes > 45) flags.push('Late Lunch');
    }
  }

  if (row.action === 'clock_out') {
    const hasLunch = sameWorkerSameDay.some((r) => r.action === 'start_lunch');
    const firstClockIn = sameWorkerSameDay.find((r) => r.action === 'clock_in');
    if (!hasLunch && firstClockIn) {
      const workedMinutes = Math.round((row.timestampMs - firstClockIn.timestampMs) / 60000);
      if (workedMinutes >= 360) flags.push('No Lunch');
    }
  }

  return uniqueArray(flags);
}

function formatFlags(flags) {
  if (!flags || !flags.length) return '-';
  return uniqueArray(flags).map((flag) => `<span class="tiny-flag">${escapeHtml(flag)}</span>`).join(' ');
}

function renderAuditLog() {
  const nameFilter = String(els.auditNameFilter?.value || '').trim().toLowerCase();
  const typeFilter = els.auditActionFilter?.value || '';

  const filtered = state.auditRows.filter((row) => {
    const matchesName = !nameFilter || String(row.workerName || '').toLowerCase().includes(nameFilter);
    const matchesType = !typeFilter || row.type === typeFilter;
    return matchesName && matchesType;
  });

  if (!filtered.length) {
    els.auditBody.innerHTML = '<tr><td colspan="7">No audit history yet.</td></tr>';
    return;
  }

  els.auditBody.innerHTML = filtered.map((row) => `
    <tr>
      <td>${formatDateTime(row.timestampMs)}</td>
      <td>${escapeHtml(row.actorName || '-')}</td>
      <td>${escapeHtml(row.type || '-')}</td>
      <td>${escapeHtml(row.workerName || '-')}</td>
      <td>${escapeHtml(row.reason || '-')}</td>
      <td><pre class="table-pre">${escapeHtml(stringifyShort(row.oldValue))}</pre></td>
      <td><pre class="table-pre">${escapeHtml(stringifyShort(row.newValue))}</pre></td>
    </tr>
  `).join('');
}

function clearAllListeners() {
  state.unsubscribers.forEach((entry) => {
    try { entry.unsub(); } catch (_) {}
  });
  state.unsubscribers = [];
}

function clearWeekDependentListeners() {
  const keep = [];
  state.unsubscribers.forEach((entry) => {
    if (entry.key === 'week_punches' || entry.key === 'timesheets') {
      try { entry.unsub(); } catch (_) {}
    } else {
      keep.push(entry);
    }
  });
  state.unsubscribers = keep;
}

function isAdmin() {
  return state.profile?.role === 'admin';
}

function prettyAction(action) {
  return String(action || '-')
    .replaceAll('_', ' ')
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function statusLabelForAction(action) {
  const map = {
    clock_in: 'Clocked In',
    start_lunch: 'On Lunch',
    end_lunch: 'Back From Lunch',
    clock_out: 'Clocked Out'
  };
  return map[action] || 'Saved';
}

function prettifyHumanName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .replaceAll(' ', '_');
}

function stringifyShort(value) {
  if (value == null) return '-';
  try {
    return JSON.stringify(value, null, 2);
  } catch {
    return String(value);
  }
}

function openPrintWindow(title, html) {
  const win = window.open('', '_blank', 'width=1000,height=800');
  if (!win) {
    toast('Pop-up blocked. Allow pop-ups to print.', true);
    return;
  }

  win.document.write(`
    <!DOCTYPE html>
    <html>
      <head>
        <title>${escapeHtml(title)}</title>
        <style>
          body { font-family: Arial, sans-serif; margin: 24px; background:#fff; color:#111; }
          table { width:100%; border-collapse:collapse; }
          th, td { border:1px solid #ccc; padding:8px; text-align:left; vertical-align:top; }
          @media print { body { margin: 12px; } }
        </style>
      </head>
      <body>
        ${html}
        <script>
          window.onload = function() { window.print(); };
        <\/script>
      </body>
    </html>
  `);
  win.document.close();
}

function getMondayDate(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function formatDateKey(date) {
  return formatDateInput(date);
}

function formatDateInput(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function formatDateOnly(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleDateString([], {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
  });
}

function formatTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleTimeString([], {
    hour: 'numeric',
    minute: '2-digit'
  });
}

function formatDateTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleString([], {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  });
}

function toLocalEditString(ms) {
  if (!ms) return '';
  const d = new Date(ms);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hour = String(d.getHours()).padStart(2, '0');
  const minute = String(d.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day} ${hour}:${minute}`;
}

function parseLocalEditString(value) {
  const cleaned = String(value || '').trim().replace('T', ' ');
  const match = cleaned.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);
  if (!match) return 0;

  const [, y, m, d, h, min] = match;
  const date = new Date(
    Number(y),
    Number(m) - 1,
    Number(d),
    Number(h),
    Number(min),
    0,
    0
  );

  const ms = date.getTime();
  return Number.isFinite(ms) ? ms : 0;
}

function currentTimeInputValue() {
  const now = new Date();
  return `${String(now.getHours()).padStart(2, '0')}:${String(now.getMinutes()).padStart(2, '0')}`;
}

function timeRangeMinutesFromStrings(startStr, endStr) {
  if (!startStr || !endStr) return 0;

  const parse = (val) => {
    const d = new Date(`2000-01-01 ${val}`);
    return Number.isFinite(d.getTime()) ? (d.getHours() * 60) + d.getMinutes() : 0;
  };

  const start = parse(startStr);
  const end = parse(endStr);
  return Math.max(0, end - start);
}

function uniqueArray(arr) {
  return [...new Set(arr.filter(Boolean))];
}

function pushUnique(arr, value) {
  if (!arr.includes(value)) arr.push(value);
}

function toast(message, isError = false) {
  if (!els.toast) return;

  els.toast.textContent = message;
  els.toast.classList.remove('hidden');
  els.toast.style.borderColor = isError
    ? 'rgba(255,107,107,0.45)'
    : 'rgba(255,255,255,0.14)';

  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => els.toast.classList.add('hidden'), 3200);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
