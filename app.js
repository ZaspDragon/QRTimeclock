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
  unsubscribers: [],
  selectedWeekStart: getMondayDate(new Date()),
  workerUnsub: null,
};

const els = {
  workerNameInput: document.getElementById('workerNameInput'),
  workerNameValue: document.getElementById('workerNameValue'),
  workerLastActionValue: document.getElementById('workerLastActionValue'),
  workerLastPunchValue: document.getElementById('workerLastPunchValue'),
  workerStatusValue: document.getElementById('workerStatusValue'),
  workerStatusMessage: document.getElementById('workerStatusMessage'),

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
  userRoleInput: document.getElementById('userRoleInput'),
  userActiveInput: document.getElementById('userActiveInput'),
  userListBody: document.getElementById('userListBody'),

  companyQrForm: document.getElementById('companyQrForm'),
  companyUrlInput: document.getElementById('companyUrlInput'),
  companyQrCanvas: document.getElementById('companyQrCanvas'),

  toast: document.getElementById('toast'),
};

init();

function init() {
  injectWorkerHistoryUi();
  wireEvents();

  const storedWorkerName = localStorage.getItem('workerPunchName') || '';
  if (storedWorkerName) {
    const pretty = prettifyHumanName(storedWorkerName);
    els.workerNameInput.value = pretty;
    els.workerNameValue.textContent = pretty;
    attachWorkerLiveView(pretty);
  }

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

    try {
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
      attachManagerLiveViews();
      attachTimesheetView();

      if (isAdmin()) {
        attachUsersView();
      }
    } catch (error) {
      console.error(error);
      toast(error.message || 'Sign-in setup failed.', true);
    }
  });
}

function injectWorkerHistoryUi() {
  const workerCard = document.getElementById('workerCard');
  if (!workerCard || document.getElementById('workerHistoryBody')) return;

  const historyWrap = document.createElement('div');
  historyWrap.className = 'mini-table-wrap';
  historyWrap.innerHTML = `
    <table>
      <thead>
        <tr>
          <th>Date/Time</th>
          <th>Action</th>
        </tr>
      </thead>
      <tbody id="workerHistoryBody">
        <tr><td colspan="2">No punches yet.</td></tr>
      </tbody>
    </table>
  `;

  workerCard.appendChild(historyWrap);
}

function wireEvents() {
  document.querySelectorAll('.worker-action-btn').forEach((btn) => {
    btn.addEventListener('click', () => handleWorkerPunch(btn.dataset.action));
  });

  els.workerNameInput.addEventListener('input', () => {
    const value = prettifyHumanName(els.workerNameInput.value.trim());
    els.workerNameValue.textContent = value || '-';

    if (value) {
      localStorage.setItem('workerPunchName', value);
      attachWorkerLiveView(value);
    }
  });

  els.loginForm.addEventListener('submit', handleLogin);
  els.resetPasswordBtn.addEventListener('click', handlePasswordReset);

  els.signOutBtn.addEventListener('click', async () => {
    await signOut(auth);
  });

  els.weekPicker.addEventListener('change', () => {
    state.selectedWeekStart = new Date(`${els.weekPicker.value}T00:00:00`);
    if (state.me && isManager()) {
      clearTimesheetListenerOnly();
      attachTimesheetView();
    }
  });

  els.tabBar.addEventListener('click', (event) => {
    const btn = event.target.closest('.tab');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });

  els.userProfileForm?.addEventListener('submit', handleSaveProfile);

  els.companyQrForm.addEventListener('submit', (e) => {
    e.preventDefault();
    renderCompanyQr();
  });
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
      createdAt: serverTimestamp(),
    });

    localStorage.setItem('workerPunchName', name);
    els.workerNameInput.value = name;
    els.workerNameValue.textContent = name;
    els.workerLastActionValue.textContent = prettyAction(action);
    els.workerLastPunchValue.textContent = formatDateTime(nowMs);
    els.workerStatusValue.textContent = statusLabelForAction(action);
    els.workerStatusMessage.textContent = `${prettyAction(action)} saved for ${name} at ${formatDateTime(nowMs)}.`;

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
  const workerHistoryBody = document.getElementById('workerHistoryBody');

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
      if (workerHistoryBody) {
        workerHistoryBody.innerHTML = '<tr><td colspan="2">No punches yet.</td></tr>';
      }
      return;
    }

    const last = rows[0];
    els.workerLastActionValue.textContent = prettyAction(last.action);
    els.workerLastPunchValue.textContent = formatDateTime(last.timestampMs);
    els.workerStatusValue.textContent = statusLabelForAction(last.action);

    const clockedInAt = findLatestClockInTime(rows);
    els.workerStatusMessage.textContent = clockedInAt
      ? `${statusLabelForAction(last.action)}. Clocked in at ${formatDateTime(clockedInAt)}.`
      : `${statusLabelForAction(last.action)} at ${formatDateTime(last.timestampMs)}.`;

    if (workerHistoryBody) {
      workerHistoryBody.innerHTML = rows.map((row) => `
        <tr>
          <td>${formatDateTime(row.timestampMs)}</td>
          <td>${prettyAction(row.action)}</td>
        </tr>
      `).join('');
    }
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load worker punches.', true);
  });
}

function findLatestClockInTime(rows) {
  const sorted = [...rows].sort((a, b) => (b.timestampMs || 0) - (a.timestampMs || 0));
  for (const row of sorted) {
    if (row.action === 'clock_in') return row.timestampMs || 0;
  }
  return 0;
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

function showLoggedOut() {
  els.authCard.classList.remove('hidden');
  els.appShell.classList.add('hidden');
  els.sessionChip.classList.add('hidden');
}

function showLoggedIn() {
  els.authCard.classList.add('hidden');
  els.appShell.classList.remove('hidden');
  els.sessionChip.classList.remove('hidden');
  els.sessionName.textContent = state.profile?.name || state.me?.email || 'Signed in';
  els.sessionRole.textContent = state.profile?.role || 'manager';
}

function attachRoleViews() {
  els.managerTabBtn.classList.remove('hidden');
  els.timesheetsTabBtn.classList.remove('hidden');
  els.adminTabBtn.classList.toggle('hidden', !isAdmin());
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

function attachManagerLiveViews() {
  const liveQuery = query(
    collection(db, 'punches'),
    orderBy('timestampMs', 'desc'),
    limit(80)
  );

  state.unsubscribers.push(
    onSnapshot(
      liveQuery,
      async (snap) => {
        const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        renderLivePunches(rows);
        renderActiveNow(rows);

        try {
          await recomputeAllTimesheetsForWeek(formatDateKey(state.selectedWeekStart));
        } catch (error) {
          console.error(error);
        }
      },
      (error) => {
        console.error(error);
        toast(error.message || 'Live punch feed failed.', true);
      }
    )
  );
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
        <td>${prettyAction(row.action)}</td>
        <td>${escapeHtml(row.source || '-')}</td>
      </tr>
    `)
    .join('');
}

function renderActiveNow(rows) {
  const latestByName = new Map();

  rows.forEach((row) => {
    const key = row.nameKey || normalizeName(row.name || '');
    if (!key) return;

    if (!latestByName.has(key)) {
      latestByName.set(key, row);
    }
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
    .map((row) => `
      <div class="person-row">
        <div class="person-meta">
          <strong>${escapeHtml(row.name || '-')}</strong>
          <span>${prettyAction(row.action)}</span>
        </div>
        <div class="pill">${formatTime(row.timestampMs)}</div>
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

  state.unsubscribers.push(
    onSnapshot(
      sheetQuery,
      (snap) => {
        const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        renderTimesheets(rows);
      },
      (error) => {
        console.error(error);
        toast(error.message || 'Timesheet view failed.', true);
      }
    )
  );
}

function renderTimesheets(rows) {
  if (!rows.length) {
    els.timesheetBody.innerHTML = '<tr><td colspan="6">No timesheets yet.</td></tr>';
    return;
  }

  els.timesheetBody.innerHTML = rows.map((row) => {
    const signedAt = row.managerSignedAt?.seconds
      ? formatDateTime(row.managerSignedAt.seconds * 1000)
      : '-';

    return `
      <tr>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${escapeHtml(row.weekKey || '-')}</td>
        <td>${Number(row.weeklyHours || 0).toFixed(2)}</td>
        <td>${escapeHtml(row.status || 'open')}</td>
        <td>${escapeHtml(row.managerSignedBy || '-')}${signedAt !== '-' ? `<br><span class="tiny">${signedAt}</span>` : ''}</td>
        <td>
          ${row.status === 'signed'
            ? `<button class="ghost-btn reopen-btn" data-id="${row.id}">Reopen</button>`
            : `<button class="primary-btn sign-btn" data-id="${row.id}">Sign</button>`}
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
      managerSignedBy: state.profile?.name || state.me?.email || 'Manager',
      managerSignedAt: Timestamp.fromDate(new Date()),
      updatedAt: serverTimestamp(),
    });
    toast('Timesheet signed.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not sign timesheet.', true);
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
    console.error(error);
    toast(error.message || 'Could not reopen timesheet.', true);
  }
}

async function recomputeAllTimesheetsForWeek(weekKey) {
  if (!state.me || !isManager()) return;

  const punchesQuery = query(
    collection(db, 'punches'),
    where('weekKey', '==', weekKey),
    orderBy('timestampMs', 'asc')
  );

  const snap = await getDocs(punchesQuery);
  const punches = snap.docs.map((d) => ({ id: d.id, ...d.data() }));

  const byName = new Map();

  punches.forEach((p) => {
    const key = p.nameKey || normalizeName(p.name || '');
    if (!key) return;

    if (!byName.has(key)) {
      byName.set(key, []);
    }

    byName.get(key).push(p);
  });

  for (const [nameKey, personPunches] of byName.entries()) {
    const displayName = personPunches[0]?.name || nameKey;
    const totals = buildWeekTotals(personPunches);
    const timesheetId = `${weekKey}_${nameKey}`;
    const existingSnap = await getDoc(doc(db, 'timesheets', timesheetId));
    const existing = existingSnap.exists() ? existingSnap.data() : null;

    await setDoc(doc(db, 'timesheets', timesheetId), {
      name: displayName,
      nameKey,
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
}

function buildWeekTotals(punches) {
  const sorted = [...punches].sort((a, b) => (a.timestampMs || 0) - (b.timestampMs || 0));
  const byDay = {};
  let currentIn = null;
  let weeklyMinutes = 0;
  let lastAction = '-';
  let lastPunchAtMs = 0;

  sorted.forEach((punch) => {
    const timeMs = punch.timestampMs || 0;
    const dateKey = punch.dateKey || formatDateKey(new Date(timeMs));

    if (!byDay[dateKey]) {
      byDay[dateKey] = { minutes: 0 };
    }

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
    Object.entries(byDay).map(([dateKey, value]) => [
      dateKey,
      Number((value.minutes / 60).toFixed(2))
    ])
  );

  return {
    dailyTotals,
    weeklyHours: Number((weeklyMinutes / 60).toFixed(2)),
    lastAction,
    lastPunchAtMs,
  };
}

function attachUsersView() {
  const usersQuery = query(collection(db, 'users'), orderBy('name', 'asc'));

  state.unsubscribers.push(
    onSnapshot(
      usersQuery,
      (snap) => {
        const rows = snap.docs.map((d) => d.data());

        if (!rows.length) {
          els.userListBody.innerHTML = '<tr><td colspan="4">No users yet.</td></tr>';
          return;
        }

        els.userListBody.innerHTML = rows
          .map((row) => `
            <tr>
              <td>${escapeHtml(row.name || '-')}</td>
              <td>${escapeHtml(row.email || '-')}</td>
              <td>${escapeHtml(row.role || '-')}</td>
              <td>${row.active ? 'Yes' : 'No'}</td>
            </tr>
          `)
          .join('');
      },
      (error) => {
        console.error(error);
        toast(error.message || 'Could not load users.', true);
      }
    )
  );
}

async function handleSaveProfile(event) {
  event.preventDefault();

  try {
    const uid = els.userUidInput.value.trim();
    await setDoc(doc(db, 'users', uid), {
      name: prettifyHumanName(els.userNameInput.value.trim()),
      email: els.userEmailInput.value.trim().toLowerCase(),
      role: els.userRoleInput.value,
      active: els.userActiveInput.value === 'true',
      updatedAt: serverTimestamp(),
    }, { merge: true });

    toast('User profile saved.');
    els.userProfileForm.reset();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save profile.', true);
  }
}

function clearLiveListeners() {
  state.unsubscribers.forEach((unsub) => {
    try { unsub(); } catch (_) {}
  });
  state.unsubscribers = [];
}

function clearTimesheetListenerOnly() {
  clearLiveListeners();

  if (state.me && isManager()) {
    attachManagerLiveViews();
    if (isAdmin()) attachUsersView();
  }
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

function renderCompanyQr() {
  const value = els.companyUrlInput.value.trim() || window.location.href;
  if (!window.QRCode) return;

  QRCode.toCanvas(els.companyQrCanvas, value, { width: 240 }, (error) => {
    if (error) {
      console.error(error);
      toast(error.message || 'Could not generate QR.', true);
    }
  });
}

function toast(message, isError = false) {
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
