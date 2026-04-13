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
  getDocs,
  setDoc,
  addDoc,
  updateDoc,
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

const state = {
  me: null,
  profile: null,
  scanner: null,
  unsubscribers: [],
  selectedWeekStart: getMondayDate(new Date()),
};

const els = {
  authCard: document.getElementById('authCard'),
  appShell: document.getElementById('appShell'),
  sessionChip: document.getElementById('sessionChip'),
  sessionName: document.getElementById('sessionName'),
  sessionRole: document.getElementById('sessionRole'),
  signOutBtn: document.getElementById('signOutBtn'),
  loginForm: document.getElementById('loginForm'),
  emailInput: document.getElementById('emailInput'),
  passwordInput: document.getElementById('passwordInput'),
  resetPasswordBtn: document.getElementById('resetPasswordBtn'),
  employeeCodeInput: document.getElementById('employeeCodeInput'),
  fillMyCodeBtn: document.getElementById('fillMyCodeBtn'),
  openScannerBtn: document.getElementById('openScannerBtn'),
  stopScannerBtn: document.getElementById('stopScannerBtn'),
  scannerWrap: document.getElementById('scannerWrap'),
  lastActionValue: document.getElementById('lastActionValue'),
  lastPunchValue: document.getElementById('lastPunchValue'),
  statusValue: document.getElementById('statusValue'),
  weekHoursValue: document.getElementById('weekHoursValue'),
  employeeStatusMessage: document.getElementById('employeeStatusMessage'),
  myPunchesBody: document.getElementById('myPunchesBody'),
  livePunchBody: document.getElementById('livePunchBody'),
  activeNowList: document.getElementById('activeNowList'),
  timesheetBody: document.getElementById('timesheetBody'),
  weekPicker: document.getElementById('weekPicker'),
  managerTabBtn: document.getElementById('managerTabBtn'),
  timesheetsTabBtn: document.getElementById('timesheetsTabBtn'),
  adminTabBtn: document.getElementById('adminTabBtn'),
  tabBar: document.getElementById('tabBar'),
  userProfileForm: document.getElementById('userProfileForm'),
  userUidInput: document.getElementById('userUidInput'),
  userNameInput: document.getElementById('userNameInput'),
  userEmailInput: document.getElementById('userEmailInput'),
  userEmployeeCodeInput: document.getElementById('userEmployeeCodeInput'),
  userRoleInput: document.getElementById('userRoleInput'),
  userActiveInput: document.getElementById('userActiveInput'),
  userListBody: document.getElementById('userListBody'),
  companyQrForm: document.getElementById('companyQrForm'),
  companyUrlInput: document.getElementById('companyUrlInput'),
  companyQrCanvas: document.getElementById('companyQrCanvas'),
  employeeQrForm: document.getElementById('employeeQrForm'),
  badgeNameInput: document.getElementById('badgeNameInput'),
  badgeCodeInput: document.getElementById('badgeCodeInput'),
  employeeQrCanvas: document.getElementById('employeeQrCanvas'),
  badgeLabel: document.getElementById('badgeLabel'),
  toast: document.getElementById('toast'),
};

init();

function init() {
  wireEvents();
  els.weekPicker.value = formatDateInput(state.selectedWeekStart);
  els.companyUrlInput.value = appSettings.defaultAppUrl || window.location.href;
  renderCompanyQr();

  onAuthStateChanged(auth, async (user) => {
    clearLiveListeners();

    if (!user) {
      state.me = null;
      state.profile = null;
      showLoggedOut();
      return;
    }

    state.me = user;
    const profileSnap = await getDoc(doc(db, 'users', user.uid));
    if (!profileSnap.exists()) {
      await signOut(auth);
      toast('No user profile found in Firestore. Add one in the users collection first.', true);
      return;
    }

    state.profile = profileSnap.data();
    showLoggedIn();
    attachRoleViews();
    attachEmployeeLiveViews();

    if (isManager()) {
      attachManagerLiveViews();
      attachTimesheetView();
      attachUsersView();
    }
  });
}

function wireEvents() {
  els.loginForm.addEventListener('submit', handleLogin);
  els.resetPasswordBtn.addEventListener('click', handlePasswordReset);
  els.signOutBtn.addEventListener('click', async () => {
    await stopScanner();
    await signOut(auth);
  });

  els.fillMyCodeBtn.addEventListener('click', () => {
    if (state.profile?.employeeId) {
      els.employeeCodeInput.value = state.profile.employeeId;
      toast('Filled your employee code.');
    }
  });

  document.querySelectorAll('.action-btn').forEach((btn) => {
    btn.addEventListener('click', () => handlePunch(btn.dataset.action));
  });

  els.openScannerBtn.addEventListener('click', startScanner);
  els.stopScannerBtn.addEventListener('click', stopScanner);
  els.weekPicker.addEventListener('change', () => {
    state.selectedWeekStart = new Date(`${els.weekPicker.value}T00:00:00`);
    if (isManager()) {
      attachTimesheetView();
    }
  });

  els.tabBar.addEventListener('click', (event) => {
    const btn = event.target.closest('.tab');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });

  els.userProfileForm.addEventListener('submit', handleSaveProfile);
  els.companyQrForm.addEventListener('submit', (e) => {
    e.preventDefault();
    renderCompanyQr();
  });
  els.employeeQrForm.addEventListener('submit', (e) => {
    e.preventDefault();
    renderEmployeeQr();
  });
}

async function handleLogin(event) {
  event.preventDefault();
  try {
    await signInWithEmailAndPassword(auth, els.emailInput.value.trim(), els.passwordInput.value);
    els.passwordInput.value = '';
  } catch (error) {
    toast(error.message, true);
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
    toast(error.message, true);
  }
}

function showLoggedOut() {
  els.authCard.classList.remove('hidden');
  els.appShell.classList.add('hidden');
  els.sessionChip.classList.add('hidden');
  switchTab('employeeTab');
}

function showLoggedIn() {
  els.authCard.classList.add('hidden');
  els.appShell.classList.remove('hidden');
  els.sessionChip.classList.remove('hidden');
  els.sessionName.textContent = state.profile?.name || state.me?.email || 'Signed in';
  els.sessionRole.textContent = state.profile?.role || 'employee';
}

function attachRoleViews() {
  const managerView = isManager();
  const adminView = isAdmin();

  els.managerTabBtn.classList.toggle('hidden', !managerView);
  els.timesheetsTabBtn.classList.toggle('hidden', !managerView);
  els.adminTabBtn.classList.toggle('hidden', !adminView);

  const firstTab = managerView ? 'employeeTab' : 'employeeTab';
  switchTab(firstTab);
}

function switchTab(tabId) {
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('active', tab.dataset.tab === tabId);
  });
  document.querySelectorAll('.tab-panel').forEach((panel) => {
    panel.classList.toggle('hidden', panel.id !== tabId);
  });
}

async function handlePunch(action) {
  try {
    const employeeId = els.employeeCodeInput.value.trim() || state.profile?.employeeId;
    if (!employeeId) {
      toast('Enter or scan an employee code first.', true);
      return;
    }

    if (employeeId !== state.profile?.employeeId && !isManager()) {
      toast('You can only punch using your own employee code.', true);
      return;
    }

    const userDoc = await findUserByEmployeeId(employeeId);
    if (!userDoc) {
      toast('No active user found for that employee code.', true);
      return;
    }

    const now = new Date();
    const dateKey = formatDateKey(now);
    const weekKey = formatDateKey(getMondayDate(now));
    const punchPayload = {
      uid: userDoc.id,
      employeeId,
      name: userDoc.data.name,
      role: userDoc.data.role,
      action,
      timestamp: serverTimestamp(),
      timestampMs: Date.now(),
      dateKey,
      weekKey,
      punchedByUid: state.me.uid,
      punchedByName: state.profile?.name || state.me.email,
      createdAt: serverTimestamp(),
    };

    await addDoc(collection(db, 'punches'), punchPayload);
    await recomputeTimesheetForWeek(userDoc.id, employeeId, userDoc.data.name, weekKey);

    localStorage.setItem('lastEmployeeCode', employeeId);
    els.employeeCodeInput.value = employeeId;
    toast(`${prettyAction(action)} saved.`);
  } catch (error) {
    toast(error.message, true);
  }
}

async function recomputeTimesheetForWeek(uid, employeeId, name, weekKey) {
  const punchesQuery = query(
    collection(db, 'punches'),
    where('uid', '==', uid),
    where('weekKey', '==', weekKey),
    orderBy('timestampMs', 'asc')
  );
  const snap = await getDocs(punchesQuery);
  const punches = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

  const totals = buildWeekTotals(punches);
  const timesheetId = `${employeeId}_${weekKey}`;
  const existingSnap = await getDoc(doc(db, 'timesheets', timesheetId));
  const existing = existingSnap.exists() ? existingSnap.data() : null;

  await setDoc(doc(db, 'timesheets', timesheetId), {
    uid,
    employeeId,
    name,
    weekKey,
    dailyTotals: totals.dailyTotals,
    weeklyHours: totals.weeklyHours,
    status: existing?.status === 'signed' ? 'signed' : 'open',
    managerSignedBy: existing?.managerSignedBy || '',
    managerSignedAt: existing?.managerSignedAt || null,
    updatedAt: serverTimestamp(),
    lastPunchAction: totals.lastAction,
    lastPunchAtMs: totals.lastPunchAtMs,
  }, { merge: true });
}

function buildWeekTotals(punches) {
  const byDay = {};
  let currentIn = null;
  let weeklyMinutes = 0;
  let lastAction = '-';
  let lastPunchAtMs = 0;

  punches.forEach((punch) => {
    const timeMs = punch.timestampMs || 0;
    const dateKey = punch.dateKey;
    byDay[dateKey] ||= { minutes: 0, actions: [] };
    byDay[dateKey].actions.push({ action: punch.action, timestampMs: timeMs });
    lastAction = punch.action;
    lastPunchAtMs = Math.max(lastPunchAtMs, timeMs);

    if (punch.action === 'clock_in' || punch.action === 'end_lunch') {
      currentIn = timeMs;
    }
    if ((punch.action === 'start_lunch' || punch.action === 'clock_out') && currentIn) {
      const diff = Math.max(0, Math.round((timeMs - currentIn) / 60000));
      weeklyMinutes += diff;
      byDay[dateKey].minutes += diff;
      currentIn = null;
    }
  });

  const dailyTotals = Object.fromEntries(
    Object.entries(byDay).map(([dateKey, value]) => [dateKey, Number((value.minutes / 60).toFixed(2))])
  );

  return {
    dailyTotals,
    weeklyHours: Number((weeklyMinutes / 60).toFixed(2)),
    lastAction,
    lastPunchAtMs,
  };
}

async function attachEmployeeLiveViews() {
  const employeeId = localStorage.getItem('lastEmployeeCode') || state.profile?.employeeId || '';
  if (employeeId) {
    els.employeeCodeInput.value = employeeId;
  }

  const todayKey = formatDateKey(new Date());
  const myPunchesQuery = query(
    collection(db, 'punches'),
    where('uid', '==', state.me.uid),
    where('dateKey', '==', todayKey),
    orderBy('timestampMs', 'desc')
  );

  state.unsubscribers.push(onSnapshot(myPunchesQuery, (snap) => {
    const rows = snap.docs.map((d) => d.data());
    renderMyPunches(rows);
  }));

  const weekKey = formatDateKey(getMondayDate(new Date()));
  const timesheetId = `${state.profile.employeeId}_${weekKey}`;
  state.unsubscribers.push(onSnapshot(doc(db, 'timesheets', timesheetId), (sheetSnap) => {
    if (!sheetSnap.exists()) {
      els.weekHoursValue.textContent = '0.00 hrs';
      return;
    }
    const data = sheetSnap.data();
    els.weekHoursValue.textContent = `${Number(data.weeklyHours || 0).toFixed(2)} hrs`;
    els.statusValue.textContent = data.status || 'open';
  }));
}

function renderMyPunches(rows) {
  if (!rows.length) {
    els.myPunchesBody.innerHTML = '<tr><td colspan="3">No punches yet.</td></tr>';
    els.lastActionValue.textContent = '-';
    els.lastPunchValue.textContent = '-';
    els.employeeStatusMessage.textContent = 'Waiting for a punch.';
    return;
  }

  const last = rows[0];
  els.lastActionValue.textContent = prettyAction(last.action);
  els.lastPunchValue.textContent = formatDateTime(last.timestampMs);
  els.employeeStatusMessage.textContent = `${prettyAction(last.action)} was saved at ${formatDateTime(last.timestampMs)}.`;

  const statusMap = {
    clock_in: 'Clocked in',
    start_lunch: 'On lunch',
    end_lunch: 'Back from lunch',
    clock_out: 'Clocked out'
  };
  els.statusValue.textContent = statusMap[last.action] || 'Open';

  els.myPunchesBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${formatTime(row.timestampMs)}</td>
        <td>${prettyAction(row.action)}</td>
        <td>${row.punchedByName || row.name}</td>
      </tr>
    `)
    .join('');
}

function attachManagerLiveViews() {
  const liveQuery = query(collection(db, 'punches'), orderBy('timestampMs', 'desc'), limit(40));
  state.unsubscribers.push(onSnapshot(liveQuery, (snap) => {
    const rows = snap.docs.map((d) => d.data());
    renderLivePunches(rows);
    renderActiveNow(rows);
  }));
}

function renderLivePunches(rows) {
  if (!rows.length) {
    els.livePunchBody.innerHTML = '<tr><td colspan="4">No live data yet.</td></tr>';
    return;
  }
  els.livePunchBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${formatDateTime(row.timestampMs)}</td>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${escapeHtml(row.employeeId || '-')}</td>
        <td>${prettyAction(row.action)}</td>
      </tr>
    `)
    .join('');
}

function renderActiveNow(rows) {
  const latestByEmployee = new Map();
  rows.forEach((row) => {
    if (!latestByEmployee.has(row.employeeId)) {
      latestByEmployee.set(row.employeeId, row);
    }
  });

  const active = [...latestByEmployee.values()].filter((row) => row.action === 'clock_in' || row.action === 'end_lunch');
  if (!active.length) {
    els.activeNowList.innerHTML = '<div class="empty-state">Nobody is currently clocked in.</div>';
    return;
  }

  els.activeNowList.innerHTML = active
    .map((row) => `
      <div class="person-row">
        <div class="person-meta">
          <strong>${escapeHtml(row.name || '-')}</strong>
          <span>${escapeHtml(row.employeeId || '-')}</span>
        </div>
        <div class="pill">${prettyAction(row.action)}</div>
      </div>
    `)
    .join('');
}

function attachTimesheetView() {
  const weekKey = formatDateKey(state.selectedWeekStart);
  const sheetQuery = query(
    collection(db, 'timesheets'),
    where('weekKey', '==', weekKey),
    orderBy('name', 'asc')
  );

  state.unsubscribers.push(onSnapshot(sheetQuery, (snap) => {
    const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
    renderTimesheets(rows);
  }));
}

function renderTimesheets(rows) {
  if (!rows.length) {
    els.timesheetBody.innerHTML = '<tr><td colspan="7">No timesheets yet.</td></tr>';
    return;
  }

  els.timesheetBody.innerHTML = rows.map((row) => {
    const signedAt = row.managerSignedAt?.seconds
      ? formatDateTime(row.managerSignedAt.seconds * 1000)
      : '-';

    return `
      <tr>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${escapeHtml(row.employeeId || '-')}</td>
        <td>${escapeHtml(row.weekKey || '-')}</td>
        <td>${Number(row.weeklyHours || 0).toFixed(2)}</td>
        <td>${escapeHtml(row.status || 'open')}</td>
        <td>${escapeHtml(row.managerSignedBy || '-')}${signedAt !== '-' ? `<br><span class="tiny">${signedAt}</span>` : ''}</td>
        <td>
          ${row.status === 'signed'
            ? `<button class="ghost-btn reopen-btn" data-id="${row.employeeId}_${row.weekKey}">Reopen</button>`
            : `<button class="primary-btn sign-btn" data-id="${row.employeeId}_${row.weekKey}">Sign</button>`}
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

async function signTimesheet(timesheetId) {
  try {
    await updateDoc(doc(db, 'timesheets', timesheetId), {
      status: 'signed',
      managerSignedBy: state.profile?.name || state.me.email,
      managerSignedAt: Timestamp.fromDate(new Date()),
      updatedAt: serverTimestamp(),
    });
    toast('Timesheet signed.');
  } catch (error) {
    toast(error.message, true);
  }
}

async function reopenTimesheet(timesheetId) {
  try {
    await updateDoc(doc(db, 'timesheets', timesheetId), {
      status: 'open',
      managerSignedBy: '',
      managerSignedAt: null,
      updatedAt: serverTimestamp(),
    });
    toast('Timesheet reopened.');
  } catch (error) {
    toast(error.message, true);
  }
}

function attachUsersView() {
  const usersQuery = query(collection(db, 'users'), orderBy('name', 'asc'));
  state.unsubscribers.push(onSnapshot(usersQuery, (snap) => {
    const rows = snap.docs.map((d) => d.data());
    if (!rows.length) {
      els.userListBody.innerHTML = '<tr><td colspan="5">No users yet.</td></tr>';
      return;
    }
    els.userListBody.innerHTML = rows.map((row) => `
      <tr>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${escapeHtml(row.email || '-')}</td>
        <td>${escapeHtml(row.employeeId || '-')}</td>
        <td>${escapeHtml(row.role || '-')}</td>
        <td>${row.active ? 'Yes' : 'No'}</td>
      </tr>
    `).join('');
  }));
}

async function handleSaveProfile(event) {
  event.preventDefault();
  try {
    const uid = els.userUidInput.value.trim();
    await setDoc(doc(db, 'users', uid), {
      name: els.userNameInput.value.trim(),
      email: els.userEmailInput.value.trim(),
      employeeId: els.userEmployeeCodeInput.value.trim(),
      role: els.userRoleInput.value,
      active: els.userActiveInput.value === 'true',
      updatedAt: serverTimestamp(),
    }, { merge: true });
    toast('User profile saved.');
    els.userProfileForm.reset();
  } catch (error) {
    toast(error.message, true);
  }
}

async function findUserByEmployeeId(employeeId) {
  const q = query(
    collection(db, 'users'),
    where('employeeId', '==', employeeId),
    where('active', '==', true),
    limit(1)
  );
  const snap = await getDocs(q);
  if (snap.empty) return null;
  const docSnap = snap.docs[0];
  return { id: docSnap.id, data: docSnap.data() };
}

async function startScanner() {
  if (!window.Html5Qrcode) {
    toast('Scanner library did not load.', true);
    return;
  }

  els.scannerWrap.classList.remove('hidden');
  els.stopScannerBtn.classList.remove('hidden');
  els.openScannerBtn.classList.add('hidden');

  state.scanner = new Html5Qrcode('reader');
  try {
    await state.scanner.start(
      { facingMode: 'environment' },
      { fps: 10, qrbox: { width: 220, height: 220 } },
      (decodedText) => {
        els.employeeCodeInput.value = decodedText.trim();
        localStorage.setItem('lastEmployeeCode', decodedText.trim());
        toast(`Scanned ${decodedText.trim()}`);
        stopScanner();
      }
    );
  } catch (error) {
    toast(`Scanner error: ${error}`, true);
    stopScanner();
  }
}

async function stopScanner() {
  if (state.scanner) {
    try {
      await state.scanner.stop();
      await state.scanner.clear();
    } catch (_) {
      // ignore cleanup errors
    }
  }
  state.scanner = null;
  els.scannerWrap.classList.add('hidden');
  els.stopScannerBtn.classList.add('hidden');
  els.openScannerBtn.classList.remove('hidden');
  const reader = document.getElementById('reader');
  if (reader) reader.innerHTML = '';
}

function renderCompanyQr() {
  const value = els.companyUrlInput.value.trim() || window.location.href;
  if (!window.QRCode) return;
  QRCode.toCanvas(els.companyQrCanvas, value, { width: 240 }, (error) => {
    if (error) toast(error.message, true);
  });
}

function renderEmployeeQr() {
  const code = els.badgeCodeInput.value.trim();
  const name = els.badgeNameInput.value.trim();
  if (!code) {
    toast('Enter an employee code first.', true);
    return;
  }
  if (!window.QRCode) return;
  QRCode.toCanvas(els.employeeQrCanvas, code, { width: 240 }, (error) => {
    if (error) {
      toast(error.message, true);
      return;
    }
    els.badgeLabel.textContent = `${name || 'Employee'} — ${code}`;
  });
}

function clearLiveListeners() {
  state.unsubscribers.forEach((unsub) => {
    try { unsub(); } catch (_) {}
  });
  state.unsubscribers = [];
}

function isManager() {
  return ['manager', 'admin'].includes(state.profile?.role);
}

function isAdmin() {
  return state.profile?.role === 'admin';
}

function prettyAction(action) {
  return String(action || '-')
    .replaceAll('_', ' ')
    .replace(/\b\w/g, (char) => char.toUpperCase());
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

function formatTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}

function formatDateTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleString([], {
    year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit'
  });
}

function toast(message, isError = false) {
  els.toast.textContent = message;
  els.toast.classList.remove('hidden');
  els.toast.style.borderColor = isError ? 'rgba(255,107,107,0.45)' : 'rgba(255,255,255,0.14)';
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
