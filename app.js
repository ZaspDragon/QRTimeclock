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

const state = {
  me: null,
  profile: null,
  unsubscribers: [],
  selectedWeekStart: getMondayDate(new Date()),
  workerUnsub: null,
  allPunchRows: [],
  selectedWeekPunchRows: [],
  selectedWeekTimesheetDocs: {},
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

async function init() {
  injectWorkerHistoryUi();
  injectManagerPunchEditorUi();
  injectAgencyExportUi();
  wireEvents();

  const storedWorkerName = localStorage.getItem('workerPunchName') || '';
  if (storedWorkerName) {
    const pretty = prettifyHumanName(storedWorkerName);
    if (els.workerNameInput) els.workerNameInput.value = pretty;
    if (els.workerNameValue) els.workerNameValue.textContent = pretty;
    attachWorkerLiveView(pretty);
  }

  if (els.weekPicker) {
    els.weekPicker.value = formatDateInput(state.selectedWeekStart);
  }

  if (els.companyUrlInput) {
    els.companyUrlInput.value = appSettings.defaultAppUrl || window.location.href;
  }

  await ensureQrCodeLibrary();
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
      attachManagerPunchEditor();
      attachAgencyExportEvents();

      if (isAdmin()) {
        attachUsersView();
      }
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
    if (els.workerNameValue) els.workerNameValue.textContent = value || '-';

    if (value) {
      localStorage.setItem('workerPunchName', value);
      attachWorkerLiveView(value);
    }
  });

  els.loginForm?.addEventListener('submit', handleLogin);
  els.resetPasswordBtn?.addEventListener('click', handlePasswordReset);

  els.signOutBtn?.addEventListener('click', async () => {
    await signOut(auth);
  });

  els.weekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = new Date(`${els.weekPicker.value}T00:00:00`);
    if (state.me && isManager()) {
      clearTimesheetListenerOnly();
      attachTimesheetView();
      attachManagerPunchEditor();
    }
  });

  els.tabBar?.addEventListener('click', (event) => {
    const btn = event.target.closest('.tab');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });

  els.userProfileForm?.addEventListener('submit', handleSaveProfile);

  els.companyQrForm?.addEventListener('submit', async (e) => {
    e.preventDefault();
    await ensureQrCodeLibrary();
    renderCompanyQr();
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

function injectManagerPunchEditorUi() {
  const managerTab = document.getElementById('managerTab');
  if (!managerTab || document.getElementById('managerPunchEditorCard')) return;

  const card = document.createElement('article');
  card.className = 'card';
  card.id = 'managerPunchEditorCard';
  card.innerHTML = `
    <div class="card-head">
      <h2>Edit punches</h2>
      <p>Managers can correct punch names, times, actions, or delete bad punches.</p>
    </div>

    <div style="display:grid;gap:12px;margin-bottom:14px;">
      <label>
        <span>Filter by name</span>
        <input id="managerPunchFilterInput" type="text" placeholder="Type a worker name" />
      </label>
    </div>

    <div class="mini-table-wrap tall">
      <table>
        <thead>
          <tr>
            <th>Date/Time</th>
            <th>Name</th>
            <th>Action</th>
            <th>Source</th>
            <th>Actions</th>
          </tr>
        </thead>
        <tbody id="managerPunchEditorBody">
          <tr><td colspan="5">No punches yet.</td></tr>
        </tbody>
      </table>
    </div>
  `;

  managerTab.appendChild(card);
}

function injectAgencyExportUi() {
  const tabBar = document.getElementById('tabBar');
  const appShell = document.getElementById('appShell');
  if (!tabBar || !appShell || document.getElementById('agencyTabBtn')) return;

  const adminBtn = document.getElementById('adminTabBtn');
  const agencyBtn = document.createElement('button');
  agencyBtn.className = 'tab';
  agencyBtn.type = 'button';
  agencyBtn.dataset.tab = 'agencyTab';
  agencyBtn.id = 'agencyTabBtn';
  agencyBtn.textContent = 'Agency Export';

  if (adminBtn) {
    adminBtn.insertAdjacentElement('afterend', agencyBtn);
  } else {
    tabBar.appendChild(agencyBtn);
  }

  const agencySection = document.createElement('section');
  agencySection.id = 'agencyTab';
  agencySection.className = 'tab-panel hidden';
  agencySection.innerHTML = `
    <div class="card">
      <div class="card-head">
        <h2>Temp Agency Export</h2>
        <p>Preview exactly what the agency will receive, then print or save it as a PDF.</p>
      </div>

      <div class="grid-form compact-form" style="margin-bottom:16px;">
        <label>
          <span>Worker</span>
          <select id="agencyWorkerSelect">
            <option value="">Select a worker</option>
          </select>
        </label>

        <div class="form-actions full-width">
          <button id="agencyPreviewBtn" class="primary-btn" type="button">Preview Sheet</button>
          <button id="agencyPrintBtn" class="secondary-btn" type="button">Print / Save PDF</button>
        </div>
      </div>

      <div id="agencyPreviewWrap" class="mini-table-wrap">
        <div id="agencyPreview" style="padding:18px;">
          <div class="empty-state">Choose a worker and click Preview Sheet.</div>
        </div>
      </div>
    </div>
  `;

  appShell.appendChild(agencySection);
}

  const agencySection = document.createElement('section');
  agencySection.id = 'agencyTab';
  agencySection.className = 'tab-panel hidden';
  agencySection.innerHTML = `
    <div class="card">
      <div class="card-head">
        <h2>Temp Agency Export</h2>
        <p>Preview exactly what the agency will receive, then print or save it as a PDF.</p>
      </div>

      <div class="grid-form compact-form" style="margin-bottom:16px;">
        <label>
          <span>Worker</span>
          <select id="agencyWorkerSelect">
            <option value="">Select a worker</option>
          </select>
        </label>

        <div class="form-actions full-width">
          <button id="agencyPreviewBtn" class="primary-btn" type="button">Preview Sheet</button>
          <button id="agencyPrintBtn" class="secondary-btn" type="button">Print / Save PDF</button>
        </div>
      </div>

      <div id="agencyPreviewWrap" class="mini-table-wrap">
        <div id="agencyPreview" style="padding:18px;">
          <div class="empty-state">Choose a worker and click Preview Sheet.</div>
        </div>
      </div>
    </div>
  `;

  qrTabPanel.parentNode.insertBefore(agencySection, qrTabPanel);
}

async function handleWorkerPunch(action) {
  const rawName = els.workerNameInput?.value.trim();
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
    if (els.workerNameInput) els.workerNameInput.value = name;
    if (els.workerNameValue) els.workerNameValue.textContent = name;
    if (els.workerLastActionValue) els.workerLastActionValue.textContent = prettyAction(action);
    if (els.workerLastPunchValue) els.workerLastPunchValue.textContent = formatDateTime(nowMs);
    if (els.workerStatusValue) els.workerStatusValue.textContent = statusLabelForAction(action);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = `${prettyAction(action)} saved for ${name} at ${formatDateTime(nowMs)}.`;
    }

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
      if (els.workerLastActionValue) els.workerLastActionValue.textContent = '-';
      if (els.workerLastPunchValue) els.workerLastPunchValue.textContent = '-';
      if (els.workerStatusValue) els.workerStatusValue.textContent = 'Ready';
      if (els.workerStatusMessage) els.workerStatusMessage.textContent = 'Enter your name and punch.';
      if (workerHistoryBody) {
        workerHistoryBody.innerHTML = '<tr><td colspan="2">No punches yet.</td></tr>';
      }
      return;
    }

    const last = rows[0];
    if (els.workerNameValue) els.workerNameValue.textContent = last.name || name;
    if (els.workerLastActionValue) els.workerLastActionValue.textContent = prettyAction(last.action);
    if (els.workerLastPunchValue) els.workerLastPunchValue.textContent = formatDateTime(last.timestampMs);
    if (els.workerStatusValue) els.workerStatusValue.textContent = statusLabelForAction(last.action);

    const clockedInAt = findLatestClockInTime(rows);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = clockedInAt
        ? `${statusLabelForAction(last.action)}. Clocked in at ${formatDateTime(clockedInAt)}.`
        : `${statusLabelForAction(last.action)} at ${formatDateTime(last.timestampMs)}.`;
    }

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
      els.emailInput?.value.trim(),
      els.passwordInput?.value
    );
    if (els.passwordInput) els.passwordInput.value = '';
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not sign in.', true);
  }
}

async function handlePasswordReset() {
  const email = els.emailInput?.value.trim();
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
  els.authCard?.classList.remove('hidden');
  els.appShell?.classList.add('hidden');
  els.sessionChip?.classList.add('hidden');
}

function showLoggedIn() {
  els.authCard?.classList.add('hidden');
  els.appShell?.classList.remove('hidden');
  els.sessionChip?.classList.remove('hidden');
  if (els.sessionName) els.sessionName.textContent = state.profile?.name || state.me?.email || 'Signed in';
  if (els.sessionRole) els.sessionRole.textContent = state.profile?.role || 'manager';
}

function attachRoleViews() {
  els.managerTabBtn?.classList.remove('hidden');
  els.timesheetsTabBtn?.classList.remove('hidden');
  els.adminTabBtn?.classList.toggle('hidden', !isAdmin());
  document.getElementById('agencyTabBtn')?.classList.remove('hidden');
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
    limit(120)
  );

  state.unsubscribers.push(
    onSnapshot(
      liveQuery,
      (snap) => {
        const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        state.allPunchRows = rows;
        renderLivePunches(rows);
        renderActiveNow(rows);
        renderManagerPunchEditor(rows);
      },
      (error) => {
        console.error(error);
        toast(error.message || 'Live punch feed failed.', true);
      }
    )
  );
}

function renderLivePunches(rows) {
  if (!els.livePunchBody) return;

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
  if (!els.activeNowList) return;

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

function attachManagerPunchEditor() {
  const filterInput = document.getElementById('managerPunchFilterInput');
  if (!filterInput || filterInput.dataset.wired === 'yes') return;

  filterInput.dataset.wired = 'yes';
  filterInput.addEventListener('input', () => {
    renderManagerPunchEditor(state.allPunchRows);
  });
}

function renderManagerPunchEditor(rows) {
  const body = document.getElementById('managerPunchEditorBody');
  const filterInput = document.getElementById('managerPunchFilterInput');
  if (!body) return;

  const filter = String(filterInput?.value || '').trim().toLowerCase();
  const filtered = rows.filter((row) => {
    if (!filter) return true;
    return String(row.name || '').toLowerCase().includes(filter);
  });

  if (!filtered.length) {
    body.innerHTML = '<tr><td colspan="5">No punches found.</td></tr>';
    return;
  }

  body.innerHTML = filtered.map((row) => `
    <tr>
      <td>${formatDateTime(row.timestampMs)}</td>
      <td>${escapeHtml(row.name || '-')}</td>
      <td>${prettyAction(row.action)}</td>
      <td>${escapeHtml(row.source || '-')}</td>
      <td>
        <button class="secondary-btn manager-edit-punch-btn" data-id="${row.id}" type="button">Edit</button>
        <button class="danger-btn manager-delete-punch-btn" data-id="${row.id}" type="button">Delete</button>
      </td>
    </tr>
  `).join('');

  body.querySelectorAll('.manager-edit-punch-btn').forEach((btn) => {
    btn.addEventListener('click', () => editPunch(btn.dataset.id));
  });

  body.querySelectorAll('.manager-delete-punch-btn').forEach((btn) => {
    btn.addEventListener('click', () => deletePunchRecord(btn.dataset.id));
  });
}

async function editPunch(punchId) {
  const row = state.allPunchRows.find((r) => r.id === punchId);
  if (!row) {
    toast('Punch not found.', true);
    return;
  }

  const newName = prompt('Edit worker name:', row.name || '');
  if (newName === null) return;

  const newAction = prompt(
    'Edit action (clock_in, start_lunch, end_lunch, clock_out):',
    row.action || 'clock_in'
  );
  if (newAction === null) return;

  const newDateTime = prompt(
    'Edit date/time (example: 2026-04-14 07:26):',
    toLocalEditString(row.timestampMs)
  );
  if (newDateTime === null) return;

  const prettyName = prettifyHumanName(newName);
  const nameKey = normalizeName(prettyName);
  const action = String(newAction).trim();

  if (!prettyName || nameKey.length < 2) {
    toast('Invalid name.', true);
    return;
  }

  if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(action)) {
    toast('Invalid action.', true);
    return;
  }

  const parsedMs = parseLocalEditString(newDateTime);
  if (!parsedMs) {
    toast('Invalid date/time format. Use YYYY-MM-DD HH:MM', true);
    return;
  }

  const date = new Date(parsedMs);
  const dateKey = formatDateKey(date);
  const weekKey = formatDateKey(getMondayDate(date));

  try {
    await updateDoc(doc(db, 'punches', punchId), {
      name: prettyName,
      nameKey,
      action,
      timestampMs: parsedMs,
      dateKey,
      weekKey,
      editedAt: serverTimestamp(),
      editedBy: state.profile?.name || state.me?.email || 'Manager'
    });

    toast('Punch updated.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not update punch.', true);
  }
}

async function deletePunchRecord(punchId) {
  const okay = confirm('Delete this punch?');
  if (!okay) return;

  try {
    await deleteDoc(doc(db, 'punches', punchId));
    toast('Punch deleted.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not delete punch.', true);
  }
}

function attachTimesheetView() {
  const weekKey = formatDateKey(state.selectedWeekStart);

  const punchesQuery = query(
    collection(db, 'punches'),
    where('weekKey', '==', weekKey),
    orderBy('timestampMs', 'asc')
  );

  const timesheetsQuery = query(
    collection(db, 'timesheets'),
    where('weekKey', '==', weekKey)
  );

  state.unsubscribers.push(
    onSnapshot(
      punchesQuery,
      (snap) => {
        state.selectedWeekPunchRows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
        renderDerivedTimesheets();
        populateAgencyWorkerSelect();
        renderAgencyPreview();
      },
      (error) => {
        console.error(error);
        toast(error.message || 'Could not load weekly punches.', true);
      }
    )
  );

  state.unsubscribers.push(
    onSnapshot(
      timesheetsQuery,
      (snap) => {
        const map = {};
        snap.docs.forEach((d) => {
          map[d.id] = { id: d.id, ...d.data() };
        });
        state.selectedWeekTimesheetDocs = map;
        renderDerivedTimesheets();
        renderAgencyPreview();
      },
      (error) => {
        console.error(error);
        toast(error.message || 'Could not load weekly signoffs.', true);
      }
    )
  );
}

function renderDerivedTimesheets() {
  if (!els.timesheetBody) return;

  const rows = getDerivedTimesheetRows();

  if (!rows.length) {
    els.timesheetBody.innerHTML = '<tr><td colspan="6">No timesheets yet.</td></tr>';
    return;
  }

  els.timesheetBody.innerHTML = rows.map((row) => {
    const signedAt = row.managerSignedAt?.seconds
      ? formatDateTime(row.managerSignedAt.seconds * 1000)
      : '-';

    const hoursText = `${Number(row.weeklyHours || 0).toFixed(2)} (${row.daysWorked || 0} day${row.daysWorked === 1 ? '' : 's'})`;

    return `
      <tr>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${escapeHtml(row.weekKey || '-')}</td>
        <td>${hoursText}</td>
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
      status: saved?.status || 'open',
      managerSignedBy: saved?.managerSignedBy || '',
      managerSignedAt: saved?.managerSignedAt || null,
    });
  });

  rows.sort((a, b) => String(a.name).localeCompare(String(b.name)));
  return rows;
}

async function signTimesheet(timesheetId) {
  const weekKey = formatDateKey(state.selectedWeekStart);
  const row = buildCurrentTimesheetRow(timesheetId, weekKey);

  if (!row) {
    toast('Could not find that weekly record.', true);
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
      status: 'signed',
      managerSignedBy: state.profile?.name || state.me?.email || 'Manager',
      managerSignedAt: Timestamp.fromDate(new Date()),
      updatedAt: serverTimestamp(),
      lastPunchAction: row.lastPunchAction,
      lastPunchAtMs: row.lastPunchAtMs,
    }, { merge: true });

    toast('Timesheet signed.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not sign timesheet.', true);
  }
}

async function reopenTimesheet(timesheetId) {
  const weekKey = formatDateKey(state.selectedWeekStart);
  const row = buildCurrentTimesheetRow(timesheetId, weekKey);

  if (!row) {
    toast('Could not find that weekly record.', true);
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
      managerSignedBy: '',
      managerSignedAt: null,
      updatedAt: serverTimestamp(),
      lastPunchAction: row.lastPunchAction,
      lastPunchAtMs: row.lastPunchAtMs,
    }, { merge: true });

    toast('Timesheet reopened.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not reopen timesheet.', true);
  }
}

function buildCurrentTimesheetRow(timesheetId, weekKey) {
  const rows = getDerivedTimesheetRows();
  return rows.find((row) => row.id === timesheetId && row.weekKey === weekKey) || null;
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
      byDay[dateKey] = {
        clock_in: '',
        start_lunch: '',
        end_lunch: '',
        clock_out: '',
        minutes: 0
      };
    }

    lastAction = punch.action;
    lastPunchAtMs = Math.max(lastPunchAtMs, timeMs);

    if (punch.action === 'clock_in') {
      byDay[dateKey].clock_in = formatTime(timeMs);
      currentIn = timeMs;
    }

    if (punch.action === 'start_lunch') {
      byDay[dateKey].start_lunch = formatTime(timeMs);
      if (currentIn) {
        const diff = Math.max(0, Math.round((timeMs - currentIn) / 60000));
        weeklyMinutes += diff;
        byDay[dateKey].minutes += diff;
        currentIn = null;
      }
    }

    if (punch.action === 'end_lunch') {
      byDay[dateKey].end_lunch = formatTime(timeMs);
      currentIn = timeMs;
    }

    if (punch.action === 'clock_out') {
      byDay[dateKey].clock_out = formatTime(timeMs);
      if (currentIn) {
        const diff = Math.max(0, Math.round((timeMs - currentIn) / 60000));
        weeklyMinutes += diff;
        byDay[dateKey].minutes += diff;
        currentIn = null;
      }
    }
  });

  const dailyTotals = Object.fromEntries(
    Object.entries(byDay).map(([dateKey, value]) => [
      dateKey,
      {
        clock_in: value.clock_in,
        start_lunch: value.start_lunch,
        end_lunch: value.end_lunch,
        clock_out: value.clock_out,
        hours: Number((value.minutes / 60).toFixed(2))
      }
    ])
  );

  const daysWorked = Object.keys(dailyTotals).length;

  return {
    dailyTotals,
    weeklyHours: Number((weeklyMinutes / 60).toFixed(2)),
    daysWorked,
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
    const uid = els.userUidInput?.value.trim();
    await setDoc(doc(db, 'users', uid), {
      name: prettifyHumanName(els.userNameInput?.value.trim()),
      email: els.userEmailInput?.value.trim().toLowerCase(),
      role: els.userRoleInput?.value,
      active: els.userActiveInput?.value === 'true',
      updatedAt: serverTimestamp(),
    }, { merge: true });

    toast('User profile saved.');
    els.userProfileForm?.reset();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save profile.', true);
  }
}

function attachAgencyExportEvents() {
  const previewBtn = document.getElementById('agencyPreviewBtn');
  const printBtn = document.getElementById('agencyPrintBtn');
  const select = document.getElementById('agencyWorkerSelect');

  if (previewBtn && previewBtn.dataset.wired !== 'yes') {
    previewBtn.dataset.wired = 'yes';
    previewBtn.addEventListener('click', () => renderAgencyPreview());
  }

  if (printBtn && printBtn.dataset.wired !== 'yes') {
    printBtn.dataset.wired = 'yes';
    printBtn.addEventListener('click', () => printAgencyPreview());
  }

  if (select && select.dataset.wired !== 'yes') {
    select.dataset.wired = 'yes';
    select.addEventListener('change', () => renderAgencyPreview());
  }
}

function populateAgencyWorkerSelect() {
  const select = document.getElementById('agencyWorkerSelect');
  if (!select) return;

  const current = select.value;
  const rows = getDerivedTimesheetRows();

  select.innerHTML = '<option value="">Select a worker</option>' +
    rows.map((row) => `<option value="${escapeHtml(row.nameKey)}">${escapeHtml(row.name)}</option>`).join('');

  if (rows.some((row) => row.nameKey === current)) {
    select.value = current;
  }
}

function renderAgencyPreview() {
  const wrap = document.getElementById('agencyPreview');
  const select = document.getElementById('agencyWorkerSelect');
  if (!wrap || !select) return;

  const selectedNameKey = select.value;
  if (!selectedNameKey) {
    wrap.innerHTML = '<div class="empty-state">Choose a worker and click Preview Sheet.</div>';
    return;
  }

  const weekKey = formatDateKey(state.selectedWeekStart);
  const row = getDerivedTimesheetRows().find((r) => r.nameKey === selectedNameKey && r.weekKey === weekKey);

  if (!row) {
    wrap.innerHTML = '<div class="empty-state">No weekly sheet found for that worker.</div>';
    return;
  }

  const signedAt = row.managerSignedAt?.seconds
    ? formatDateTime(row.managerSignedAt.seconds * 1000)
    : '-';

  const dailyRows = buildAgencyDailyRows(row.dailyTotals);

  wrap.innerHTML = `
    <div id="agencyPrintableSheet" style="background:#fff;color:#111;border-radius:12px;padding:24px;min-height:200px;">
      <div style="display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap;margin-bottom:18px;">
        <div>
          <h2 style="margin:0 0 8px;font-size:28px;">Temp Agency Weekly Time Sheet</h2>
          <div style="font-size:15px;line-height:1.6;">
            <div><strong>Worker:</strong> ${escapeHtml(row.name)}</div>
            <div><strong>Week Start:</strong> ${escapeHtml(row.weekKey)}</div>
            <div><strong>Status:</strong> ${escapeHtml(row.status)}</div>
          </div>
        </div>
        <div style="font-size:14px;line-height:1.7;text-align:right;">
          <div><strong>Company:</strong> ${escapeHtml(appSettings.companyName || 'Company')}</div>
          <div><strong>Generated:</strong> ${escapeHtml(formatDateTime(Date.now()))}</div>
        </div>
      </div>

      <table style="width:100%;border-collapse:collapse;margin-bottom:18px;">
        <thead>
          <tr>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Date</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Clock In</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Lunch Out</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Lunch In</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Clock Out</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Hours</th>
          </tr>
        </thead>
        <tbody>
          ${dailyRows}
        </tbody>
      </table>

      <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-top:24px;">
        <div style="font-size:15px;line-height:1.8;">
          <div><strong>Total Hours:</strong> ${Number(row.weeklyHours || 0).toFixed(2)}</div>
          <div><strong>Days Worked:</strong> ${Number(row.daysWorked || 0)}</div>
        </div>

        <div style="font-size:15px;line-height:1.8;text-align:right;">
          <div><strong>Manager:</strong> ${escapeHtml(row.managerSignedBy || '-')}</div>
          <div><strong>Signed:</strong> ${escapeHtml(signedAt)}</div>
        </div>
      </div>
    </div>
  `;
}

function buildAgencyDailyRows(dailyTotals) {
  const keys = Object.keys(dailyTotals || {}).sort();
  if (!keys.length) {
    return `
      <tr>
        <td colspan="6" style="border:1px solid #bbb;padding:10px;">No punches recorded for this week.</td>
      </tr>
    `;
  }

  return keys.map((dateKey) => {
    const row = dailyTotals[dateKey] || {};
    return `
      <tr>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(dateKey)}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.clock_in || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.start_lunch || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.end_lunch || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.clock_out || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${Number(row.hours || 0).toFixed(2)}</td>
      </tr>
    `;
  }).join('');
}

function printAgencyPreview() {
  const sheet = document.getElementById('agencyPrintableSheet');
  if (!sheet) {
    toast('Preview the sheet first.', true);
    return;
  }

  const win = window.open('', '_blank', 'width=1000,height=800');
  if (!win) {
    toast('Pop-up blocked. Allow pop-ups to print.', true);
    return;
  }

  win.document.write(`
    <!DOCTYPE html>
    <html>
      <head>
        <title>Agency Time Sheet</title>
        <style>
          body {
            font-family: Arial, sans-serif;
            margin: 24px;
            color: #111;
            background: #fff;
          }
          @media print {
            body {
              margin: 12px;
            }
          }
        </style>
      </head>
      <body>
        ${sheet.outerHTML}
        <script>
          window.onload = function() {
            window.print();
          };
        <\/script>
      </body>
    </html>
  `);
  win.document.close();
}

async function ensureQrCodeLibrary() {
  if (window.QRCode && typeof window.QRCode.toCanvas === 'function') {
    return true;
  }

  const existing = document.querySelector('script[data-qrcode-lib="yes"]');
  if (existing) {
    await waitForQrCode(2000);
    return !!window.QRCode;
  }

  const script = document.createElement('script');
  script.src = 'https://cdn.jsdelivr.net/npm/qrcode@1.5.3/build/qrcode.min.js';
  script.async = true;
  script.dataset.qrcodeLib = 'yes';

  const loaded = new Promise((resolve, reject) => {
    script.onload = () => resolve(true);
    script.onerror = () => reject(new Error('Could not load QRCode library.'));
  });

  document.head.appendChild(script);

  try {
    await loaded;
    await waitForQrCode(1000);
    return !!window.QRCode;
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not load QRCode library.', true);
    return false;
  }
}

function waitForQrCode(timeoutMs = 1000) {
  return new Promise((resolve) => {
    const started = Date.now();

    function tick() {
      if (window.QRCode && typeof window.QRCode.toCanvas === 'function') {
        resolve(true);
        return;
      }

      if (Date.now() - started >= timeoutMs) {
        resolve(false);
        return;
      }

      setTimeout(tick, 50);
    }

    tick();
  });
}

function renderCompanyQr() {
  const value = els.companyUrlInput?.value.trim() || window.location.href;

  if (!window.QRCode || typeof window.QRCode.toCanvas !== 'function') {
    toast('QRCode library did not load.', true);
    return;
  }

  QRCode.toCanvas(els.companyQrCanvas, value, { width: 240 }, (error) => {
    if (error) {
      console.error(error);
      toast(error.message || 'Could not generate QR.', true);
      return;
    }
    toast('QR generated.');
  });
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
    attachTimesheetView();
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
