import { firebaseConfig, appSettings } from './firebase-config.js';
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  getAuth,
  createUserWithEmailAndPassword,
  signInWithEmailAndPassword,
  onAuthStateChanged,
  signOut,
  updateProfile,
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
  writeBatch,
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

// Branch deployment config. Change CURRENT_SITE_ID to "OHC" when deploying an OHC-specific manager app.
const CURRENT_COMPANY_ID = 'chadwell';
const CURRENT_SITE_ID = 'OH01';
const BRANCH_OPTIONS = [
  { siteId: 'OH01', label: 'OH01' },
  { siteId: 'OHC', label: 'OHC' },
];

const state = {
  me: null,
  profile: null,
  companyId: CURRENT_COMPANY_ID,
  siteId: CURRENT_SITE_ID,
  agencyId: null,         // from user profile (null = direct company user)
  companyDoc: null,       // loaded from companies/{companyId}
  unsubscribers: [],
  selectedWeekStart: getMondayDate(new Date()),
  workerUnsub: null,
  workerEmployee: null,      // looked-up employee record for public punch
  publicEmployees: [],       // cached employee list for public QR autocomplete
  allPunchRows: [],
  selectedWeekPunchRows: [],
  selectedWeekTimesheetDocs: {},
  allEmployees: [],
  allMissedRequests: [],
  approvalFilter: 'pending',
  creatingPendingProfile: false,
};

const els = {
  workerNameInput: document.getElementById('workerNameInput'),
  workerBranchSelect: document.getElementById('workerBranchSelect'),
  workerAutocompleteList: document.getElementById('workerAutocompleteList'),
  workerLookupStatus: document.getElementById('workerLookupStatus'),
  workerNameValue: document.getElementById('workerNameValue'),
  workerLastActionValue: document.getElementById('workerLastActionValue'),
  workerLastPunchValue: document.getElementById('workerLastPunchValue'),
  workerStatusValue: document.getElementById('workerStatusValue'),
  workerStatusMessage: document.getElementById('workerStatusMessage'),
  workerHistoryBody: document.getElementById('workerHistoryBody'),

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
  signupForm: document.getElementById('signupForm'),
  signupNameInput: document.getElementById('signupNameInput'),
  signupEmailInput: document.getElementById('signupEmailInput'),
  signupPasswordInput: document.getElementById('signupPasswordInput'),
  signupRequestedRoleInput: document.getElementById('signupRequestedRoleInput'),
  signupSiteInput: document.getElementById('signupSiteInput'),

  livePunchBody: document.getElementById('livePunchBody'),
  activeNowList: document.getElementById('activeNowList'),
  timesheetBody: document.getElementById('timesheetBody'),
  weekPicker: document.getElementById('weekPicker'),
  managerTabBtn: document.getElementById('managerTabBtn'),
  timesheetsTabBtn: document.getElementById('timesheetsTabBtn'),
  editPunchesTabBtn: document.getElementById('editPunchesTabBtn'),
  adminTabBtn: document.getElementById('adminTabBtn'),
  agencyTabBtn: document.getElementById('agencyTabBtn'),
  tabBar: document.getElementById('tabBar'),

  manualPunchForm: document.getElementById('manualPunchForm'),
  manualPunchNameInput: document.getElementById('manualPunchNameInput'),
  manualPunchActionInput: document.getElementById('manualPunchActionInput'),
  manualPunchDateInput: document.getElementById('manualPunchDateInput'),
  manualPunchTimeInput: document.getElementById('manualPunchTimeInput'),
  editFilterNameInput: document.getElementById('editFilterNameInput'),
  editPunchesBody: document.getElementById('editPunchesBody'),

  userProfileForm: document.getElementById('userProfileForm'),
  userUidInput: document.getElementById('userUidInput'),
  userNameInput: document.getElementById('userNameInput'),
  userEmailInput: document.getElementById('userEmailInput'),
  userRoleInput: document.getElementById('userRoleInput'),
  userActiveInput: document.getElementById('userActiveInput'),
  userSiteIdsInput: document.getElementById('userSiteIdsInput'),
  userListBody: document.getElementById('userListBody'),
  pendingUserListBody: document.getElementById('pendingUserListBody'),

  myTimecardTabBtn: document.getElementById('myTimecardTabBtn'),
  missedPunchTabBtn: document.getElementById('missedPunchTabBtn'),
  missedPunchForm: document.getElementById('missedPunchForm'),
  mpActionInput: document.getElementById('mpActionInput'),
  mpDateInput: document.getElementById('mpDateInput'),
  mpTimeInput: document.getElementById('mpTimeInput'),
  mpReasonInput: document.getElementById('mpReasonInput'),
  myMissedPunchBody: document.getElementById('myMissedPunchBody'),
  myTimecardWeekPicker: document.getElementById('myTimecardWeekPicker'),
  myTcTotalHours: document.getElementById('myTcTotalHours'),
  myTcDaysWorked: document.getElementById('myTcDaysWorked'),
  myTcLastPunch: document.getElementById('myTcLastPunch'),
  myTcStatus: document.getElementById('myTcStatus'),
  myTimecardBody: document.getElementById('myTimecardBody'),

  approvalsTabBtn: document.getElementById('approvalsTabBtn'),
  approvalFilterAll: document.getElementById('approvalFilterAll'),
  approvalFilterPending: document.getElementById('approvalFilterPending'),
  approvalFilterApproved: document.getElementById('approvalFilterApproved'),
  approvalFilterDenied: document.getElementById('approvalFilterDenied'),
  approvalListBody: document.getElementById('approvalListBody'),

  employeesTabBtn: document.getElementById('employeesTabBtn'),
  employeeForm: document.getElementById('employeeForm'),
  employeeDocId: document.getElementById('employeeDocId'),
  empNameInput: document.getElementById('empNameInput'),
  empNumberInput: document.getElementById('empNumberInput'),
  empAgencySelect: document.getElementById('empAgencySelect'),
  empSiteInput: document.getElementById('empSiteInput'),
  empStatusSelect: document.getElementById('empStatusSelect'),
  empCancelEditBtn: document.getElementById('empCancelEditBtn'),
  empFilterInput: document.getElementById('empFilterInput'),
  employeeListBody: document.getElementById('employeeListBody'),

  agencyWorkerSelect: document.getElementById('agencyWorkerSelect'),
  agencyPreviewBtn: document.getElementById('agencyPreviewBtn'),
  agencyPrintBtn: document.getElementById('agencyPrintBtn'),
  agencyPreview: document.getElementById('agencyPreview'),

  toast: document.getElementById('toast'),
};

init();

async function init() {
  wireEvents();
  setupBranchSelectors();

  // Load employees for public QR autocomplete
  loadPublicEmployees();

  // Restore last-used worker name
  const storedWorkerName = localStorage.getItem('workerPunchName') || '';
  if (storedWorkerName) {
    const pretty = prettifyHumanName(storedWorkerName);
    if (els.workerNameInput) els.workerNameInput.value = pretty;
    if (els.workerNameValue) els.workerNameValue.textContent = pretty;
    // Try to match to an existing employee
    restoreWorkerFromName(pretty);
    attachWorkerLiveView(pretty);
  }

  if (els.weekPicker) {
    els.weekPicker.value = formatDateInput(state.selectedWeekStart);
  }

  if (els.manualPunchDateInput) {
    els.manualPunchDateInput.value = formatDateInput(new Date());
  }

  if (els.manualPunchTimeInput) {
    els.manualPunchTimeInput.value = formatTimeForInput(Date.now());
  }

  onAuthStateChanged(auth, async (user) => {
    clearLiveListeners();

    if (!user) {
      state.me = null;
      state.profile = null;
      showLoggedOut();
      return;
    }

    try {
      if (state.creatingPendingProfile) return;
      state.me = user;
      const profileSnap = await getDoc(doc(db, 'users', user.uid));

      if (!profileSnap.exists()) {
        await signOut(auth);
        toast('No user profile found. Ask an admin to approve your account.', true);
        return;
      }

      state.profile = normalizeUserProfile(profileSnap.data(), user);
      if (state.profile.active !== true) {
        await signOut(auth);
        toast('Your account is not active yet. Ask an admin to approve your account.', true);
        return;
      }

      if (state.profile.companyId !== CURRENT_COMPANY_ID) {
        await signOut(auth);
        toast('You do not have access to this company.', true);
        return;
      }

      if (!profileCanAccessSite(state.profile, CURRENT_SITE_ID)) {
        await signOut(auth);
        toast('You do not have access to this branch.', true);
        return;
      }

      state.companyId = CURRENT_COMPANY_ID;
      state.siteId = CURRENT_SITE_ID;
      state.agencyId = state.profile.agencyId || null;

      // Load company doc if companyId exists
      if (state.companyId) {
        try {
          const compSnap = await getDoc(doc(db, 'companies', state.companyId));
          state.companyDoc = compSnap.exists() ? compSnap.data() : null;
        } catch (_) {
          state.companyDoc = null;
        }
      }

      showLoggedIn();
      attachRoleViews();
      attachManagerLiveViews();
      attachTimesheetView();
      attachUsersViewIfAdmin();
      attachPendingUsersViewIfAdmin();
      populateAgencyWorkerSelect();
      renderAgencyPreview();
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

  // Name input with autocomplete
  els.workerNameInput?.addEventListener('input', debounce(handleWorkerNameAutocomplete, 250));
  els.workerNameInput?.addEventListener('focus', () => handleWorkerNameAutocomplete());
  els.workerNameInput?.addEventListener('keydown', handleAutocompleteKeydown);

  // Close autocomplete on outside click
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.autocomplete-wrap')) {
      hideAutocomplete();
    }
  });

  els.loginForm?.addEventListener('submit', handleLogin);
  els.signupForm?.addEventListener('submit', handleSignupRequest);
  els.resetPasswordBtn?.addEventListener('click', handlePasswordReset);

  els.workerBranchSelect?.addEventListener('change', () => {
    localStorage.setItem('workerPunchSiteId', getPublicSiteId());
    state.workerEmployee = null;
    loadPublicEmployees().then(() => {
      handleWorkerNameAutocomplete();
      const typed = prettifyHumanName(els.workerNameInput?.value.trim() || '');
      if (typed) attachWorkerLiveView(typed);
    });
  });

  els.signOutBtn?.addEventListener('click', async () => {
    await signOut(auth);
  });

  els.weekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = new Date(`${els.weekPicker.value}T00:00:00`);
    if (state.me && isManager()) {
      clearTimesheetListenerOnly();
      attachTimesheetView();
    }
  });

  els.tabBar?.addEventListener('click', (event) => {
    const btn = event.target.closest('.tab');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });

  els.manualPunchForm?.addEventListener('submit', handleManualPunchSubmit);

  els.editFilterNameInput?.addEventListener('input', () => {
    renderEditPunchesTable(state.allPunchRows);
  });

  els.userProfileForm?.addEventListener('submit', handleSaveProfile);

  els.missedPunchForm?.addEventListener('submit', handleMissedPunchSubmit);

  els.myTimecardWeekPicker?.addEventListener('change', () => {
    if (state.me && isEmployee()) {
      clearMyTimecardListener();
      attachMyTimecardView();
    }
  });

  [['approvalFilterAll', 'all'], ['approvalFilterPending', 'pending'],
   ['approvalFilterApproved', 'approved'], ['approvalFilterDenied', 'denied']].forEach(([id, val]) => {
    els[id]?.addEventListener('click', () => {
      state.approvalFilter = val;
      renderApprovalList(state.allMissedRequests);
    });
  });

  els.employeeForm?.addEventListener('submit', handleSaveEmployee);
  els.empCancelEditBtn?.addEventListener('click', cancelEmployeeEdit);
  els.empFilterInput?.addEventListener('input', () => renderEmployeeList(state.allEmployees || []));

  els.agencyPreviewBtn?.addEventListener('click', () => renderAgencyPreview());
  els.agencyPrintBtn?.addEventListener('click', () => printAgencyPreview());
  els.agencyWorkerSelect?.addEventListener('change', () => renderAgencyPreview());
}

// ─── Public employee loading & autocomplete ─────────────
function setupBranchSelectors() {
  populateBranchSelect(els.workerBranchSelect, localStorage.getItem('workerPunchSiteId') || CURRENT_SITE_ID);
  populateBranchSelect(els.signupSiteInput, CURRENT_SITE_ID);
  populateBranchSelect(els.empSiteInput, CURRENT_SITE_ID);
}

function populateBranchSelect(selectEl, selectedSiteId = CURRENT_SITE_ID) {
  if (!selectEl) return;
  selectEl.innerHTML = BRANCH_OPTIONS
    .map((branch) => `<option value="${branch.siteId}">${branch.label}</option>`)
    .join('');
  selectEl.value = BRANCH_OPTIONS.some((branch) => branch.siteId === selectedSiteId)
    ? selectedSiteId
    : CURRENT_SITE_ID;
}

function getPublicSiteId() {
  const selected = String(els.workerBranchSelect?.value || CURRENT_SITE_ID).trim();
  return BRANCH_OPTIONS.some((branch) => branch.siteId === selected) ? selected : CURRENT_SITE_ID;
}

function getCurrentCompanyId() {
  return CURRENT_COMPANY_ID;
}

function getCurrentSiteId() {
  return CURRENT_SITE_ID;
}

function getAllowedSiteIds(profile = state.profile) {
  if (!profile) return [];
  if (Array.isArray(profile.siteIds) && profile.siteIds.length) {
    return profile.siteIds.map((siteId) => String(siteId || '').trim()).filter(Boolean);
  }
  const singleSiteId = String(profile.siteId || '').trim();
  return singleSiteId ? [singleSiteId] : [];
}

function parseSiteIds(value, fallbackToCurrent = true) {
  const raw = Array.isArray(value) ? value : String(value || '').split(',');
  const allowed = new Set(BRANCH_OPTIONS.map((branch) => branch.siteId));
  const sites = raw
    .map((siteId) => String(siteId || '').trim())
    .filter((siteId) => allowed.has(siteId));
  if (!sites.length && !fallbackToCurrent) return [];
  return [...new Set(sites.length ? sites : [CURRENT_SITE_ID])];
}

function profileCanAccessSite(profile, siteId) {
  return String(profile?.companyId || '').trim() === CURRENT_COMPANY_ID
    && getAllowedSiteIds(profile).includes(siteId);
}

function canUseSite(siteId) {
  return profileCanAccessSite(state.profile, siteId);
}

function branchConstraints(siteId = getCurrentSiteId()) {
  return [
    where('companyId', '==', getCurrentCompanyId()),
    where('siteId', '==', siteId),
  ];
}

function branchPayload(siteId = getCurrentSiteId()) {
  return {
    companyId: getCurrentCompanyId(),
    siteId,
  };
}

async function loadPublicEmployees() {
  try {
    const constraints = [
      ...branchConstraints(getPublicSiteId()),
      where('status', '==', 'active'),
    ];

    const q = query(collection(db, 'employees'), ...constraints);
    const snap = await getDocs(q);
    state.publicEmployees = snap.docs
      .map((d) => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (a.name || '').localeCompare(b.name || ''));
  } catch (error) {
    console.warn('Could not load employees for autocomplete:', error.message);
    state.publicEmployees = [];
  }
}

function restoreWorkerFromName(name) {
  const nameKey = normalizeName(name);
  const match = state.publicEmployees.find((e) => normalizeName(e.name) === nameKey);
  if (match) {
    state.workerEmployee = match;
    if (els.workerLookupStatus) {
      els.workerLookupStatus.textContent = `✓ Welcome back, ${match.name}. Ready to punch.`;
      els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
    }
  }
}

let _acActiveIndex = -1;

function handleWorkerNameAutocomplete() {
  const raw = els.workerNameInput?.value.trim() || '';
  const typed = prettifyHumanName(raw);
  if (els.workerNameValue) els.workerNameValue.textContent = typed || '-';

  if (typed.length < 2) {
    state.workerEmployee = null;
    hideAutocomplete();
    if (els.workerLookupStatus) {
      els.workerLookupStatus.textContent = 'Type your name to begin.';
      els.workerLookupStatus.style.borderColor = '';
    }
    return;
  }

  const lower = typed.toLowerCase();
  const matches = state.publicEmployees.filter((e) =>
    (e.name || '').toLowerCase().includes(lower)
  ).slice(0, 8);

  // Check for exact match
  const exactMatch = state.publicEmployees.find(
    (e) => normalizeName(e.name) === normalizeName(typed)
  );

  if (exactMatch) {
    state.workerEmployee = exactMatch;
    if (els.workerLookupStatus) {
      els.workerLookupStatus.textContent = `✓ Found: ${exactMatch.name} (${exactMatch.employeeNumber || 'new'}). Ready to punch.`;
      els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
    }
    hideAutocomplete();
    localStorage.setItem('workerPunchName', exactMatch.name);
    attachWorkerLiveView(exactMatch.name);
    return;
  }

  // No exact match — show suggestions + "new worker" option
  state.workerEmployee = null;
  if (els.workerLookupStatus) {
    els.workerLookupStatus.textContent = matches.length
      ? `${matches.length} match${matches.length > 1 ? 'es' : ''} found. Select or keep typing.`
      : `New worker — "${typed}" will be created on first punch.`;
    els.workerLookupStatus.style.borderColor = matches.length ? '' : 'rgba(59,213,118,0.4)';
  }

  renderAutocomplete(matches, typed);
}

function renderAutocomplete(matches, typed) {
  const list = els.workerAutocompleteList;
  if (!list) return;

  _acActiveIndex = -1;
  let html = '';

  matches.forEach((emp, i) => {
    html += `<li data-index="${i}" data-emp-id="${emp.id}">
      ${escapeHTML(emp.name)}<span class="emp-num">${escapeHTML(emp.employeeNumber || '')}</span>
    </li>`;
  });

  // Always show "new worker" option at the bottom if typed name doesn't match
  if (typed.length >= 2) {
    html += `<li class="new-worker" data-index="${matches.length}" data-new="true">
      + Create new: "${escapeHTML(typed)}"
    </li>`;
  }

  list.innerHTML = html;
  list.hidden = false;

  // Wire click handlers
  list.querySelectorAll('li').forEach((li) => {
    li.addEventListener('click', () => {
      if (li.dataset.new === 'true') {
        selectNewWorker(typed);
      } else {
        const emp = matches[parseInt(li.dataset.index)];
        selectAutocompleteEmployee(emp);
      }
    });
  });
}

function handleAutocompleteKeydown(e) {
  const list = els.workerAutocompleteList;
  if (!list || list.hidden) return;

  const items = list.querySelectorAll('li');
  if (!items.length) return;

  if (e.key === 'ArrowDown') {
    e.preventDefault();
    _acActiveIndex = Math.min(_acActiveIndex + 1, items.length - 1);
    updateActiveItem(items);
  } else if (e.key === 'ArrowUp') {
    e.preventDefault();
    _acActiveIndex = Math.max(_acActiveIndex - 1, 0);
    updateActiveItem(items);
  } else if (e.key === 'Enter' && _acActiveIndex >= 0) {
    e.preventDefault();
    items[_acActiveIndex]?.click();
  } else if (e.key === 'Escape') {
    hideAutocomplete();
  }
}

function updateActiveItem(items) {
  items.forEach((li, i) => li.classList.toggle('active', i === _acActiveIndex));
}

function selectAutocompleteEmployee(emp) {
  state.workerEmployee = emp;
  if (els.workerNameInput) els.workerNameInput.value = emp.name;
  if (els.workerNameValue) els.workerNameValue.textContent = emp.name;
  if (els.workerLookupStatus) {
    els.workerLookupStatus.textContent = `✓ Found: ${emp.name} (${emp.employeeNumber || ''}). Ready to punch.`;
    els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
  }
  hideAutocomplete();
  localStorage.setItem('workerPunchName', emp.name);
  attachWorkerLiveView(emp.name);
}

function selectNewWorker(typed) {
  const pretty = prettifyHumanName(typed);
  // Set a placeholder worker employee — will be auto-created on punch
  state.workerEmployee = { _isNew: true, name: pretty, nameKey: normalizeName(pretty) };
  if (els.workerNameInput) els.workerNameInput.value = pretty;
  if (els.workerNameValue) els.workerNameValue.textContent = pretty;
  if (els.workerLookupStatus) {
    els.workerLookupStatus.textContent = `✓ New worker: "${pretty}". Punch to clock in and create profile.`;
    els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
  }
  hideAutocomplete();
  localStorage.setItem('workerPunchName', pretty);
}

function hideAutocomplete() {
  if (els.workerAutocompleteList) {
    els.workerAutocompleteList.hidden = true;
    els.workerAutocompleteList.innerHTML = '';
  }
  _acActiveIndex = -1;
}

function escapeHTML(str) {
  const div = document.createElement('div');
  div.textContent = str;
  return div.innerHTML;
}

async function handleWorkerPunch(action) {
  let emp = state.workerEmployee;
  const typedName = prettifyHumanName(els.workerNameInput?.value.trim() || '');

  // If nothing selected but a name was typed, treat as new worker
  if (!emp && typedName.length >= 2) {
    emp = { _isNew: true, name: typedName, nameKey: normalizeName(typedName) };
    state.workerEmployee = emp;
  }

  if (!emp || !emp.name) {
    toast('Type your name first.', true);
    return;
  }

  if (emp.status === 'inactive' || emp.status === 'terminated') {
    toast('Your employee record is not active. Contact your manager.', true);
    return;
  }

  const publicSiteId = getPublicSiteId();

  // Auto-create employee if new
  if (emp._isNew) {
    try {
      const empNumber = await generateNextPublicEmployeeNumber();
      const scope = {
        employeeNumber: empNumber,
        nameKey: normalizeName(emp.name),
        companyId: getCurrentCompanyId(),
        agencyId: '',
        siteId: publicSiteId
      };
      const existingEmployee = await findExistingEmployeeForUpsert(scope);
      const employeeId = existingEmployee?.id || buildStableEmployeeId(empNumber, scope.agencyId, scope.siteId || scope.companyId);
      const newPayload = {
        name: emp.name,
        nameKey: scope.nameKey,
        normalizedName: scope.nameKey,
        employeeNumber: existingEmployee?.employeeNumber || empNumber,
        employeeNumberKey: normalizeWorkerNumber(existingEmployee?.employeeNumber || empNumber),
        companyId: scope.companyId,
        agencyId: '',
        assignedSiteId: publicSiteId,
        siteId: publicSiteId,
        status: 'active',
        active: true,
        employeeId,
        source: 'auto_created',
        updatedAt: serverTimestamp(),
      };
      if (!existingEmployee) {
        newPayload.createdAt = serverTimestamp();
        await setDoc(doc(db, 'employees', employeeId), newPayload, { merge: true });
      }

      emp = { id: employeeId, ...existingEmployee, ...newPayload };
      state.workerEmployee = emp;

      // Update local cache
      if (!state.publicEmployees.some((row) => row.id === employeeId || row.employeeId === employeeId)) {
        state.publicEmployees.push(emp);
      }

      if (els.workerLookupStatus) {
        els.workerLookupStatus.textContent = existingEmployee
          ? `Found: ${emp.name} (${emp.employeeNumber || empNumber}). Punching...`
          : `Created: ${emp.name} (${empNumber}). Punching...`;
        els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
      }
    } catch (error) {
      console.warn('Auto-create employee skipped (rules may not be deployed yet):', error.message);
      // Still allow the punch — employee record will be created by manager or once rules are deployed
      emp = { name: emp.name, nameKey: normalizeName(emp.name), employeeId: '', employeeNumber: '' };
      state.workerEmployee = emp;
    }
  }

  const name = emp.name || '';
  const nameKey = normalizeName(name);
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
      employeeId: emp.employeeId || emp.id || '',
      employeeNumber: emp.employeeNumber || '',
      companyId: getCurrentCompanyId(),
      siteId: emp.siteId || publicSiteId,
      agencyId: emp.agencyId || '',
    });

    if (els.workerLastActionValue) els.workerLastActionValue.textContent = prettyAction(action);
    if (els.workerLastPunchValue) els.workerLastPunchValue.textContent = formatDateTime(nowMs);
    if (els.workerStatusValue) els.workerStatusValue.textContent = statusLabelForAction(action);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = `${prettyAction(action)} saved for ${name} at ${formatDateTime(nowMs)}.`;
    }

    attachWorkerLiveView(name);
    localStorage.setItem('workerPunchName', name);
    toast(`${prettyAction(action)} saved.`);
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save punch.', true);
  }
}

async function generateNextPublicEmployeeNumber() {
  // Check both local cache and generate next EMP number
  const prefix = 'EMP-';
  const existing = (state.publicEmployees || [])
    .map((e) => e.employeeNumber || '')
    .filter((n) => n.startsWith(prefix))
    .map((n) => parseInt(n.replace(prefix, ''), 10))
    .filter((n) => !isNaN(n));

  const maxNum = existing.length ? Math.max(...existing) : 1000;
  return prefix + String(maxNum + 1);
}

async function handleManualPunchSubmit(event) {
  event.preventDefault();

  if (!canEditPunches()) {
    toast('You need edit-punch permission to add manual punches.', true);
    return;
  }

  const name = prettifyHumanName(els.manualPunchNameInput?.value.trim());
  const nameKey = normalizeName(name);
  const action = els.manualPunchActionInput?.value;
  const dateValue = els.manualPunchDateInput?.value;
  const timeValue = els.manualPunchTimeInput?.value;

  if (!name || !nameKey) {
    toast('Enter a valid name.', true);
    return;
  }

  if (!action || !dateValue || !timeValue) {
    toast('Fill out all manual punch fields.', true);
    return;
  }

  const parsedMs = parseLocalDateAndTime(dateValue, timeValue);
  if (!parsedMs) {
    toast('Invalid date or time.', true);
    return;
  }

  const punchDate = new Date(parsedMs);
  const dateKey = formatDateKey(punchDate);
  const weekKey = formatDateKey(getMondayDate(punchDate));

  try {
    await addDoc(collection(db, 'punches'), {
      name,
      nameKey,
      action,
      timestamp: serverTimestamp(),
      timestampMs: parsedMs,
      dateKey,
      weekKey,
      source: 'manual_manager',
      createdAt: serverTimestamp(),
      createdBy: state.profile?.name || state.me?.email || 'Manager',
      ...branchPayload(),
      agencyId: '',
      employeeId: '',
    });

    await addDoc(collection(db, 'punch_edits'), {
      type: 'manual_add',
      name,
      nameKey,
      action,
      timestampMs: parsedMs,
      dateKey,
      weekKey,
      source: 'manual_manager',
      editedBy: state.profile?.name || state.me?.email || 'Manager',
      editedAt: serverTimestamp(),
      ...branchPayload(),
    });

    els.manualPunchForm?.reset();
    if (els.manualPunchDateInput) els.manualPunchDateInput.value = formatDateInput(new Date());
    if (els.manualPunchTimeInput) els.manualPunchTimeInput.value = formatTimeForInput(Date.now());

    toast('Manual punch added.');
    await logAudit('punch_manual_added', 'punch', nameKey, {}, {
      name,
      nameKey,
      action,
      timestampMs: parsedMs,
      ...branchPayload(),
    }, 'Manual manager punch');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not add manual punch.', true);
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
    ...branchConstraints(getPublicSiteId()),
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
      if (els.workerHistoryBody) {
        els.workerHistoryBody.innerHTML = '<tr><td colspan="2">No punches yet.</td></tr>';
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

    if (els.workerHistoryBody) {
      els.workerHistoryBody.innerHTML = rows.map((row) => `
        <tr>
          <td>${formatDateTime(row.timestampMs)}</td>
          <td>${prettyAction(row.action)}</td>
        </tr>
      `).join('');
    }
  }, (error) => {
    // Permission errors are expected for unauthenticated workers (original rules restrict reads to managers)
    if (error.code === 'permission-denied' || (error.message && error.message.includes('permissions'))) {
      console.info('Live punch view not available (read requires authentication). Punch data was saved.');
    } else {
      console.error('Live view error:', error);
      toast(error.message || 'Could not load worker punches.', true);
    }
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

async function handleSignupRequest(event) {
  event.preventDefault();

  const name = prettifyHumanName(els.signupNameInput?.value.trim());
  const email = String(els.signupEmailInput?.value || '').trim().toLowerCase();
  const password = els.signupPasswordInput?.value || '';
  const requestedRole = els.signupRequestedRoleInput?.value || 'manager';
  const siteId = els.signupSiteInput?.value || CURRENT_SITE_ID;

  if (!name || !email || password.length < 6) {
    toast('Enter a name, email, and password with at least 6 characters.', true);
    return;
  }

  try {
    state.creatingPendingProfile = true;
    const credential = await createUserWithEmailAndPassword(auth, email, password);
    if (name) {
      await updateProfile(credential.user, { displayName: name });
    }

    await setDoc(doc(db, 'users', credential.user.uid), {
      uid: credential.user.uid,
      active: false,
      approvalStatus: 'pending',
      requestedRole,
      role: 'worker',
      companyId: getCurrentCompanyId(),
      siteId,
      siteIds: [siteId],
      email,
      name,
      displayName: name,
      permissions: {
        canEditPunches: false,
        canDeletePunches: false,
        canMergeWorkers: false,
        manageUsers: false,
      },
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp(),
    }, { merge: true });

    await signOut(auth);
    els.signupForm?.reset();
    populateBranchSelect(els.signupSiteInput, CURRENT_SITE_ID);
    toast('Account request submitted. Ask an admin to approve your account.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not request account access.', true);
  } finally {
    state.creatingPendingProfile = false;
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
  state.companyId = CURRENT_COMPANY_ID;
  state.siteId = CURRENT_SITE_ID;
  state.agencyId = null;
  state.companyDoc = null;
  els.authCard?.classList.remove('hidden');
  els.appShell?.classList.add('hidden');
  els.sessionChip?.classList.add('hidden');
  // Restore public worker card
  const workerCard = document.getElementById('workerCard');
  if (workerCard) workerCard.classList.remove('hidden');
  // Reset header
  const headerP = document.querySelector('.topbar p');
  if (headerP) headerP.textContent = 'Mobile punch tracking with live manager visibility and weekly signoff.';
}

function showLoggedIn() {
  els.authCard?.classList.add('hidden');
  els.appShell?.classList.remove('hidden');
  els.sessionChip?.classList.remove('hidden');
  // Hide the public worker card for logged-in users
  const workerCard = document.getElementById('workerCard');
  if (workerCard) workerCard.classList.add('hidden');
  if (els.sessionName) els.sessionName.textContent = state.profile?.name || state.me?.email || 'Signed in';
  if (els.sessionRole) {
    const roleParts = [state.profile?.role || 'manager'];
    roleParts.push(CURRENT_SITE_ID);
    if (state.agencyId) roleParts.push('agency');
    els.sessionRole.textContent = roleParts.join(' · ');
  }

  // Show company name in header
  const companyDisplayName = state.companyDoc?.name || (state.companyId ? state.companyId : appSettings.companyName);
  const headerP = document.querySelector('.topbar p');
  if (headerP) headerP.textContent = companyDisplayName + ' — TimeClock Pro';
}

function getCompanyName() {
  return state.companyDoc?.name || state.companyId || appSettings.companyName;
}

/** Returns true if current user is scoped to an agency */
function isAgencyUser() {
  return !!state.agencyId;
}

const AGENCY_NAMES = {
  sterling_staffing: 'Sterling Staffing',
  excel_staffing: 'Excel Staffing',
};

function agencyLabel(agencyId) {
  if (!agencyId) return 'Direct';
  return AGENCY_NAMES[agencyId] || agencyId;
}

function attachRoleViews() {
  const emp = isEmployee();
  const mgr = isManager();

  // Employee-only tabs
  els.myTimecardTabBtn?.classList.toggle('hidden', !emp);
  els.missedPunchTabBtn?.classList.toggle('hidden', !emp);

  // Manager/admin tabs
  els.managerTabBtn?.classList.toggle('hidden', emp);
  els.timesheetsTabBtn?.classList.toggle('hidden', emp);
  els.editPunchesTabBtn?.classList.toggle('hidden', emp);
  els.approvalsTabBtn?.classList.toggle('hidden', !mgr);
  els.employeesTabBtn?.classList.toggle('hidden', !mgr);
  els.adminTabBtn?.classList.toggle('hidden', !isAdmin());
  els.agencyTabBtn?.classList.toggle('hidden', emp);

  if (emp) {
    if (els.myTimecardWeekPicker) {
      els.myTimecardWeekPicker.value = formatDateInput(state.selectedWeekStart);
    }
    if (els.mpDateInput) els.mpDateInput.value = formatDateInput(new Date());
    switchTab('myTimecardTab');
    attachMyTimecardView();
    attachMyMissedPunchView();
  } else {
    switchTab('managerTab');
    if (mgr) {
      attachEmployeesView();
      attachApprovalView();
    }
  }
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
  const constraints = [...branchConstraints()];
  if (isAgencyUser()) constraints.push(where('agencyId', '==', state.agencyId));
  constraints.push(orderBy('timestampMs', 'desc'));
  constraints.push(limit(250));

  const liveQuery = query(
    collection(db, 'punches'),
    ...constraints
  );

  state.unsubscribers.push(
    onSnapshot(
      liveQuery,
      (snap) => {
        const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() })).filter(isActivePunchRecord);
        state.allPunchRows = rows;
        renderLivePunches(rows);
        renderActiveNow(rows);
        renderEditPunchesTable(rows);
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

function renderEditPunchesTable(rows) {
  if (!els.editPunchesBody) return;

  const filter = String(els.editFilterNameInput?.value || '').trim().toLowerCase();
  const filtered = rows.filter((row) => {
    if (!filter) return true;
    return String(row.name || '').toLowerCase().includes(filter);
  });

  if (!filtered.length) {
    els.editPunchesBody.innerHTML = '<tr><td colspan="8">No punches found.</td></tr>';
    return;
  }

  els.editPunchesBody.innerHTML = filtered.map((row) => {
    const editedAtText = row.editedAt?.seconds
      ? formatDateTime(row.editedAt.seconds * 1000)
      : '-';

    const rowClass = row.editedBy ? 'class="edited-row"' : '';

    return `
      <tr ${rowClass}>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${prettyAction(row.action)}</td>
        <td>${formatDateOnly(row.timestampMs)}</td>
        <td>${formatTime(row.timestampMs)}</td>
        <td>${escapeHtml(row.source || '-')}</td>
        <td>${escapeHtml(row.editedBy || '-')}</td>
        <td>${escapeHtml(editedAtText)}</td>
        <td>
          <button class="secondary-btn manager-edit-punch-btn" data-id="${row.id}" type="button" ${canEditPunches() ? '' : 'disabled'}>Edit</button>
          <button class="danger-btn manager-delete-punch-btn" data-id="${row.id}" type="button" ${canDeletePunches() ? '' : 'disabled'}>Delete</button>
        </td>
      </tr>
    `;
  }).join('');

  els.editPunchesBody.querySelectorAll('.manager-edit-punch-btn').forEach((btn) => {
    btn.addEventListener('click', () => editPunch(btn.dataset.id));
  });

  els.editPunchesBody.querySelectorAll('.manager-delete-punch-btn').forEach((btn) => {
    btn.addEventListener('click', () => deletePunchRecord(btn.dataset.id));
  });
}

async function editPunch(punchId) {
  if (!canEditPunches()) {
    toast('You need edit-punch permission to edit punches.', true);
    return;
  }

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

  const updatedPayload = {
    name: prettyName,
    nameKey,
    action,
    timestampMs: parsedMs,
    dateKey,
    weekKey,
    editedAt: serverTimestamp(),
    editedBy: state.profile?.name || state.me?.email || 'Manager'
  };

  try {
    await addDoc(collection(db, 'punch_edits'), {
      punchId,
      type: 'edit',
      original: {
        name: row.name || '',
        nameKey: row.nameKey || '',
        action: row.action || '',
        timestampMs: row.timestampMs || 0,
        dateKey: row.dateKey || '',
        weekKey: row.weekKey || '',
        source: row.source || '',
        editedBy: row.editedBy || '',
      },
      updated: {
        name: prettyName,
        nameKey,
        action,
        timestampMs: parsedMs,
        dateKey,
        weekKey,
        source: row.source || ''
      },
      editedBy: state.profile?.name || state.me?.email || 'Manager',
      editedAt: serverTimestamp(),
      ...branchPayload()
    });

    await updateDoc(doc(db, 'punches', punchId), { ...updatedPayload, ...branchPayload(row.siteId || CURRENT_SITE_ID) });
    await logAudit('punch_edited', 'punch', punchId, row || {}, updatedPayload, 'Punch edited from manager dashboard');

    toast('Punch updated.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not update punch.', true);
  }
}

async function deletePunchRecord(punchId) {
  if (!canDeletePunches()) {
    toast('You need delete-punch permission to delete punches.', true);
    return;
  }

  const row = state.allPunchRows.find((r) => r.id === punchId);
  const okay = confirm('Delete this punch?');
  if (!okay) return;

  try {
    await addDoc(collection(db, 'punch_edits'), {
      punchId,
      type: 'delete',
      original: row || null,
      editedBy: state.profile?.name || state.me?.email || 'Manager',
      editedAt: serverTimestamp(),
      ...branchPayload(row?.siteId || CURRENT_SITE_ID)
    });

    const deletePayload = {
      status: 'deleted',
      active: false,
      deletedAt: serverTimestamp(),
      deletedBy: state.profile?.name || state.me?.email || 'Manager',
      deleteReason: 'Manager deleted from punch editor',
      updatedAt: serverTimestamp()
    };
    await updateDoc(doc(db, 'punches', punchId), deletePayload);
    await logAudit('punch_deleted', 'punch', punchId, row || {}, deletePayload, 'Soft delete from punch editor');
    toast('Punch marked deleted.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not delete punch.', true);
  }
}

function attachTimesheetView() {
  const weekKey = formatDateKey(state.selectedWeekStart);

  const punchConstraints = [
    ...branchConstraints(),
    where('weekKey', '==', weekKey),
  ];
  if (isAgencyUser()) punchConstraints.push(where('agencyId', '==', state.agencyId));
  punchConstraints.push(orderBy('timestampMs', 'asc'));

  const punchesQuery = query(collection(db, 'punches'), ...punchConstraints);

  const tsConstraints = [
    ...branchConstraints(),
    where('weekKey', '==', weekKey),
  ];
  if (isAgencyUser()) tsConstraints.push(where('agencyId', '==', state.agencyId));

  const timesheetsQuery = query(collection(db, 'timesheets'), ...tsConstraints);

  state.unsubscribers.push(
    onSnapshot(
      punchesQuery,
      (snap) => {
        state.selectedWeekPunchRows = snap.docs.map((d) => ({ id: d.id, ...d.data() })).filter(isActivePunchRecord);
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
      companyId: personPunches[0]?.companyId || getCurrentCompanyId(),
      agencyId: personPunches[0]?.agencyId || state.agencyId || '',
      siteId: personPunches[0]?.siteId || personPunches[0]?.assignedSiteId || getCurrentSiteId(),
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
      companyId: row.companyId || getCurrentCompanyId(),
      agencyId: row.agencyId || state.agencyId || '',
      siteId: row.siteId || getCurrentSiteId(),
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
      companyId: row.companyId || getCurrentCompanyId(),
      agencyId: row.agencyId || state.agencyId || '',
      siteId: row.siteId || getCurrentSiteId(),
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

function isEmployee() {
  return hasAnyRole(['employee', 'worker']);
}

/* ───────────────────────────────────────────────────
   MY TIMECARD (employee self-service view)
   ─────────────────────────────────────────────────── */

function attachMyTimecardView() {
  const weekStart = els.myTimecardWeekPicker?.value
    ? new Date(`${els.myTimecardWeekPicker.value}T00:00:00`)
    : state.selectedWeekStart;

  const weekKey = formatDateKey(weekStart);
  const employeeId = state.profile?.employeeId || null;
  const nameKey = normalizeName(state.profile?.name || '');

  if (!employeeId && !nameKey) {
    toast('Your profile is missing employeeId. Ask your manager.', true);
    return;
  }

  // Query punches by employeeId (preferred) or nameKey (legacy fallback)
  const constraints = [where('weekKey', '==', weekKey)];
  if (employeeId) {
    constraints.push(where('employeeId', '==', employeeId));
  } else {
    constraints.push(where('nameKey', '==', nameKey));
  }
  constraints.push(...branchConstraints());
  constraints.push(orderBy('timestampMs', 'asc'));

  const q = query(collection(db, 'punches'), ...constraints);

  state._myTcUnsub = onSnapshot(q, (snap) => {
    const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() })).filter(isActivePunchRecord);
    renderMyTimecard(rows);
  }, (error) => {
    console.error(error);
    toast(error.message || 'Could not load your timecard.', true);
  });

  state.unsubscribers.push(state._myTcUnsub);
}

function clearMyTimecardListener() {
  if (state._myTcUnsub) {
    try { state._myTcUnsub(); } catch (_) {}
    state._myTcUnsub = null;
  }
}

function renderMyTimecard(punches) {
  const totals = buildWeekTotals(punches);

  if (els.myTcTotalHours) els.myTcTotalHours.textContent = Number(totals.weeklyHours || 0).toFixed(2);
  if (els.myTcDaysWorked) els.myTcDaysWorked.textContent = String(totals.daysWorked || 0);
  if (els.myTcLastPunch) els.myTcLastPunch.textContent = totals.lastPunchAtMs ? formatDateTime(totals.lastPunchAtMs) : '-';
  if (els.myTcStatus) els.myTcStatus.textContent = totals.lastAction ? statusLabelForAction(totals.lastAction) : '-';

  if (!els.myTimecardBody) return;

  const daily = totals.dailyTotals;
  const keys = Object.keys(daily).sort();

  if (!keys.length) {
    els.myTimecardBody.innerHTML = '<tr><td colspan="6">No punches this week.</td></tr>';
    return;
  }

  els.myTimecardBody.innerHTML = keys.map((dateKey) => {
    const d = daily[dateKey];
    return `
      <tr>
        <td>${escapeHtml(dateKey)}</td>
        <td>${escapeHtml(d.clock_in || '-')}</td>
        <td>${escapeHtml(d.start_lunch || '-')}</td>
        <td>${escapeHtml(d.end_lunch || '-')}</td>
        <td>${escapeHtml(d.clock_out || '-')}</td>
        <td>${Number(d.hours || 0).toFixed(2)}</td>
      </tr>
    `;
  }).join('');
}

/* ───────────────────────────────────────────────────
   MISSED PUNCH REQUESTS
   ─────────────────────────────────────────────────── */

async function handleMissedPunchSubmit(event) {
  event.preventDefault();

  if (!state.me || !state.profile) {
    toast('You must be signed in.', true);
    return;
  }

  const action = els.mpActionInput?.value;
  const dateValue = els.mpDateInput?.value;
  const timeValue = els.mpTimeInput?.value;
  const reason = els.mpReasonInput?.value.trim();

  if (!action || !dateValue || !timeValue || !reason) {
    toast('Fill out all fields.', true);
    return;
  }

  const requestedTimestampMs = parseLocalDateAndTime(dateValue, timeValue);
  if (!requestedTimestampMs) {
    toast('Invalid date or time.', true);
    return;
  }

  try {
    await addDoc(collection(db, 'missedPunchRequests'), {
      uid: state.me.uid,
      employeeId: state.profile.employeeId || '',
      ...branchPayload(),
      agencyId: state.agencyId || '',
      name: state.profile.name || '',
      requestedAction: action,
      requestedDate: dateValue,
      requestedTime: timeValue,
      requestedTimestampMs,
      reason,
      status: 'pending',
      reviewedBy: '',
      reviewedAt: null,
      approvedBy: '',
      approvedAt: null,
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp(),
    });

    els.missedPunchForm?.reset();
    if (els.mpDateInput) els.mpDateInput.value = formatDateInput(new Date());
    toast('Missed punch request submitted.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not submit request.', true);
  }
}

function attachMyMissedPunchView() {
  if (!state.me) return;

  const constraints = [
    ...branchConstraints(),
    where('uid', '==', state.me.uid),
  ];

  const q = query(collection(db, 'missedPunchRequests'), ...constraints);

  state.unsubscribers.push(
    onSnapshot(q, (snap) => {
      const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      rows.sort((a, b) => (b.requestedTimestampMs || 0) - (a.requestedTimestampMs || 0));
      renderMyMissedPunches(rows);
    }, (error) => {
      console.error(error);
    })
  );
}

function renderMyMissedPunches(rows) {
  if (!els.myMissedPunchBody) return;

  if (!rows.length) {
    els.myMissedPunchBody.innerHTML = '<tr><td colspan="6">No requests yet.</td></tr>';
    return;
  }

  els.myMissedPunchBody.innerHTML = rows.map((r) => {
    const statusClass = r.status === 'approved' ? 'color:var(--good)' :
                        r.status === 'denied' ? 'color:var(--danger)' : 'color:var(--warn)';
    return `
      <tr>
        <td>${escapeHtml(r.requestedDate || '-')}</td>
        <td>${escapeHtml(r.requestedTime || '-')}</td>
        <td>${prettyAction(r.requestedAction)}</td>
        <td>${escapeHtml(r.reason || '-')}</td>
        <td><span style="${statusClass};font-weight:700;text-transform:capitalize;">${escapeHtml(r.status || 'pending')}</span></td>
        <td>${escapeHtml(r.reviewedBy || '-')}</td>
      </tr>
    `;
  }).join('');
}

/* ───────────────────────────────────────────────────
   MANAGER APPROVAL DASHBOARD
   ─────────────────────────────────────────────────── */

function attachApprovalView() {
  const constraints = [...branchConstraints()];
  if (isAgencyUser()) constraints.push(where('agencyId', '==', state.agencyId));

  const q = query(collection(db, 'missedPunchRequests'), ...constraints);

  state.unsubscribers.push(
    onSnapshot(q, (snap) => {
      state.allMissedRequests = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      state.allMissedRequests.sort((a, b) => (b.requestedTimestampMs || 0) - (a.requestedTimestampMs || 0));
      renderApprovalList(state.allMissedRequests);
    }, (error) => {
      console.error(error);
      toast(error.message || 'Could not load approval requests.', true);
    })
  );
}

function renderApprovalList(requests) {
  if (!els.approvalListBody) return;

  const filtered = state.approvalFilter === 'all'
    ? requests
    : requests.filter((r) => r.status === state.approvalFilter);

  if (!filtered.length) {
    els.approvalListBody.innerHTML = `<tr><td colspan="7">No ${state.approvalFilter} requests.</td></tr>`;
    return;
  }

  els.approvalListBody.innerHTML = filtered.map((r) => {
    const statusClass = r.status === 'approved' ? 'color:var(--good)' :
                        r.status === 'denied' ? 'color:var(--danger)' : 'color:var(--warn)';
    const actions = r.status === 'pending'
      ? `<button class="primary-btn approve-req-btn" data-id="${r.id}" type="button" style="margin-right:6px;">Approve</button>
         <button class="danger-btn deny-req-btn" data-id="${r.id}" type="button">Deny</button>`
      : `<span class="tiny">${escapeHtml(r.reviewedBy || '-')}</span>`;

    return `
      <tr>
        <td>${escapeHtml(r.name || '-')}</td>
        <td>${escapeHtml(r.requestedDate || '-')}</td>
        <td>${escapeHtml(r.requestedTime || '-')}</td>
        <td>${prettyAction(r.requestedAction)}</td>
        <td>${escapeHtml(r.reason || '-')}</td>
        <td><span style="${statusClass};font-weight:700;text-transform:capitalize;">${escapeHtml(r.status)}</span></td>
        <td>${actions}</td>
      </tr>
    `;
  }).join('');

  els.approvalListBody.querySelectorAll('.approve-req-btn').forEach((btn) => {
    btn.addEventListener('click', () => approveRequest(btn.dataset.id));
  });

  els.approvalListBody.querySelectorAll('.deny-req-btn').forEach((btn) => {
    btn.addEventListener('click', () => denyRequest(btn.dataset.id));
  });
}

async function approveRequest(requestId) {
  if (!isManager()) { toast('Only managers can approve.', true); return; }

  const req = state.allMissedRequests.find((r) => r.id === requestId);
  if (!req) { toast('Request not found.', true); return; }

  const managerName = state.profile?.name || state.me?.email || 'Manager';
  const now = Timestamp.fromDate(new Date());

  try {
    // 1. Update the request status
    await updateDoc(doc(db, 'missedPunchRequests', requestId), {
      status: 'approved',
      reviewedBy: managerName,
      reviewedAt: now,
      approvedBy: managerName,
      approvedAt: now,
      updatedAt: serverTimestamp(),
    });
    await logAudit('missed_punch_approved', 'missedPunchRequest', requestId, req, {
      status: 'approved',
      reviewedBy: managerName,
    }, 'Missed punch request approved');

    // 2. Auto-create the actual punch
    const punchDate = new Date(req.requestedTimestampMs);
    const dateKey = formatDateKey(punchDate);
    const weekKey = formatDateKey(getMondayDate(punchDate));

    await addDoc(collection(db, 'punches'), {
      name: req.name || '',
      nameKey: normalizeName(req.name || ''),
      action: req.requestedAction,
      timestamp: serverTimestamp(),
      timestampMs: req.requestedTimestampMs,
      dateKey,
      weekKey,
      source: 'missed_punch_approved',
      createdAt: serverTimestamp(),
      createdBy: managerName,
      approvedBy: managerName,
      approvedAt: now,
      companyId: req.companyId || getCurrentCompanyId(),
      siteId: req.siteId || getCurrentSiteId(),
      agencyId: req.agencyId || '',
      employeeId: req.employeeId || '',
    });

    toast('Request approved — punch created.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not approve request.', true);
  }
}

async function denyRequest(requestId) {
  if (!isManager()) { toast('Only managers can deny.', true); return; }

  const req = state.allMissedRequests.find((r) => r.id === requestId);
  if (!req) { toast('Request not found.', true); return; }

  const reason = prompt('Denial reason (optional):') || '';
  const managerName = state.profile?.name || state.me?.email || 'Manager';

  try {
    await updateDoc(doc(db, 'missedPunchRequests', requestId), {
      status: 'denied',
      reviewedBy: managerName,
      reviewedAt: Timestamp.fromDate(new Date()),
      editReason: reason,
      updatedAt: serverTimestamp(),
    });
    await logAudit('missed_punch_denied', 'missedPunchRequest', requestId, req, {
      status: 'denied',
      reviewedBy: managerName,
      reason,
    }, 'Missed punch request denied');

    toast('Request denied.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not deny request.', true);
  }
}

/* ───────────────────────────────────────────────────
   EMPLOYEES COLLECTION (employees/{employeeId})
   ─────────────────────────────────────────────────── */

function attachEmployeesView() {
  const empConstraints = [...branchConstraints()];
  if (isAgencyUser()) empConstraints.push(where('agencyId', '==', state.agencyId));
  empConstraints.push(orderBy('name', 'asc'));

  const empQuery = query(collection(db, 'employees'), ...empConstraints);

  state.unsubscribers.push(
    onSnapshot(empQuery, (snap) => {
      state.allEmployees = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderEmployeeList(state.allEmployees);
    }, (error) => {
      console.error(error);
      toast(error.message || 'Could not load employees.', true);
    })
  );
}

function renderEmployeeList(employees) {
  if (!els.employeeListBody) return;

  const filter = String(els.empFilterInput?.value || '').trim().toLowerCase();
  const filtered = employees.filter((e) => {
    if (!filter) return true;
    return (
      String(e.name || '').toLowerCase().includes(filter) ||
      String(e.employeeNumber || '').toLowerCase().includes(filter)
    );
  });

  if (!filtered.length) {
    els.employeeListBody.innerHTML = '<tr><td colspan="6">No employees found.</td></tr>';
    return;
  }

  els.employeeListBody.innerHTML = filtered.map((emp) => `
    <tr>
      <td>${escapeHtml(emp.employeeNumber || '-')}</td>
      <td>${escapeHtml(emp.name || '-')}</td>
      <td>${escapeHtml(agencyLabel(emp.agencyId))}</td>
      <td>${escapeHtml(emp.siteId || emp.assignedSiteId || '-')}</td>
      <td><span class="tiny-flag">${escapeHtml(emp.status || 'active')}</span></td>
      <td>
        <button class="secondary-btn emp-edit-btn" data-id="${emp.id}" type="button">Edit</button>
      </td>
    </tr>
  `).join('');

  els.employeeListBody.querySelectorAll('.emp-edit-btn').forEach((btn) => {
    btn.addEventListener('click', () => loadEmployeeForEdit(btn.dataset.id));
  });
}

function loadEmployeeForEdit(empId) {
  const emp = (state.allEmployees || []).find((e) => e.id === empId);
  if (!emp) { toast('Employee not found.', true); return; }

  if (els.employeeDocId) els.employeeDocId.value = empId;
  if (els.empNameInput) els.empNameInput.value = emp.name || '';
  if (els.empNumberInput) els.empNumberInput.value = emp.employeeNumber || '';
  if (els.empAgencySelect) els.empAgencySelect.value = emp.agencyId || '';
  if (els.empSiteInput) els.empSiteInput.value = emp.siteId || emp.assignedSiteId || CURRENT_SITE_ID;
  if (els.empStatusSelect) els.empStatusSelect.value = emp.status || 'active';
  els.empCancelEditBtn?.classList.remove('hidden');
}

function cancelEmployeeEdit() {
  els.employeeForm?.reset();
  if (els.employeeDocId) els.employeeDocId.value = '';
  els.empCancelEditBtn?.classList.add('hidden');
}

async function handleSaveEmployee(event) {
  event.preventDefault();

  if (!isManager()) {
    toast('Only managers and admins can manage employees.', true);
    return;
  }

  const name = prettifyHumanName(els.empNameInput?.value.trim());
  const nameKey = normalizeName(name);

  if (!name || nameKey.length < 2) {
    toast('Enter a valid employee name.', true);
    return;
  }

  let employeeNumber = els.empNumberInput?.value.trim();
  const agencyId = els.empAgencySelect?.value || '';
  const requestedSiteId = String(els.empSiteInput?.value || CURRENT_SITE_ID).trim();
  const assignedSiteId = canUseSite(requestedSiteId) ? requestedSiteId : CURRENT_SITE_ID;
  if (requestedSiteId !== assignedSiteId) {
    toast('You do not have access to assign that branch.', true);
    return;
  }
  const status = els.empStatusSelect?.value || 'active';
  const existingId = els.employeeDocId?.value || '';

  // Auto-generate employee number if blank
  if (!employeeNumber) {
    employeeNumber = await generateNextEmployeeNumber();
  }

  const payload = {
    name,
    nameKey,
    normalizedName: nameKey,
    employeeNumber,
    employeeNumberKey: normalizeWorkerNumber(employeeNumber),
    companyId: getCurrentCompanyId(),
    agencyId,
    assignedSiteId,
    siteId: assignedSiteId,
    status,
    active: status === 'active',
    updatedAt: serverTimestamp(),
  };

  try {
    if (existingId) {
      // Update existing employee
      await setDoc(doc(db, 'employees', existingId), payload, { merge: true });
      await logAudit('employee_updated', 'employee', existingId, {}, payload, 'Admin employee update');
      toast('Employee updated.');
    } else {
      const reusableEmployee = await findExistingEmployeeForUpsert({
        employeeNumber,
        nameKey,
        companyId: getCurrentCompanyId(),
        agencyId,
        siteId: assignedSiteId
      });
      const employeeId = reusableEmployee?.id || buildStableEmployeeId(employeeNumber, agencyId, assignedSiteId || getCurrentCompanyId());
      payload.employeeId = employeeId;
      if (!reusableEmployee) {
        payload.createdAt = serverTimestamp();
      }

      await setDoc(doc(db, 'employees', employeeId), payload, { merge: true });
      await logAudit(reusableEmployee ? 'employee_reused_existing' : 'employee_created', 'employee', employeeId, reusableEmployee || {}, payload, 'Admin employee save with stable ID');
      toast(reusableEmployee ? 'Existing active employee reused and updated.' : 'Employee created: ' + employeeNumber);
    }

    cancelEmployeeEdit();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save employee.', true);
  }
}

async function generateNextEmployeeNumber() {
  // Find the highest existing employee number and increment
  const prefix = 'EMP-';
  const existing = (state.allEmployees || [])
    .map((e) => e.employeeNumber || '')
    .filter((n) => n.startsWith(prefix))
    .map((n) => parseInt(n.replace(prefix, ''), 10))
    .filter((n) => !isNaN(n));

  const maxNum = existing.length ? Math.max(...existing) : 1000;
  return prefix + String(maxNum + 1);
}

function attachUsersViewIfAdmin() {
  if (!canManageUsers()) return;
  attachUsersView();
}

function attachPendingUsersViewIfAdmin() {
  if (!canManageUsers() || !els.pendingUserListBody) return;

  const pendingQuery = query(
    collection(db, 'users'),
    where('companyId', '==', getCurrentCompanyId()),
    where('siteIds', 'array-contains', getCurrentSiteId()),
    where('approvalStatus', '==', 'pending')
  );

  state.unsubscribers.push(
    onSnapshot(pendingQuery, (snap) => {
      const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderPendingUsers(rows);
    }, (error) => {
      console.error(error);
      toast(error.message || 'Could not load pending users.', true);
    })
  );
}

function renderPendingUsers(rows) {
  if (!els.pendingUserListBody) return;

  if (!rows.length) {
    els.pendingUserListBody.innerHTML = '<tr><td colspan="6">No pending users.</td></tr>';
    return;
  }

  els.pendingUserListBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.name || row.displayName || '-')}</td>
      <td>${escapeHtml(row.email || '-')}</td>
      <td>${escapeHtml(row.requestedRole || '-')}</td>
      <td>${escapeHtml(row.companyId || '-')}</td>
      <td>${escapeHtml(getAllowedSiteIds(row).join(', ') || '-')}</td>
      <td>
        <button class="primary-btn approve-user-btn" data-id="${row.id}" type="button">Approve</button>
        <button class="danger-btn deny-user-btn" data-id="${row.id}" type="button">Deny</button>
      </td>
    </tr>
  `).join('');

  els.pendingUserListBody.querySelectorAll('.approve-user-btn').forEach((btn) => {
    btn.addEventListener('click', () => approvePendingUser(btn.dataset.id, rows));
  });

  els.pendingUserListBody.querySelectorAll('.deny-user-btn').forEach((btn) => {
    btn.addEventListener('click', () => denyPendingUser(btn.dataset.id, rows));
  });
}

async function approvePendingUser(uid, rows) {
  if (!canManageUsers()) {
    toast('You do not have permission to approve users.', true);
    return;
  }

  const row = rows.find((item) => item.id === uid);
  if (!row) return;

  const requestedRole = row.requestedRole || 'manager';
  const role = prompt('Approve role (manager, admin, worker):', requestedRole) || requestedRole;
  if (!['manager', 'admin', 'worker'].includes(role)) {
    toast('Invalid role.', true);
    return;
  }

  const siteIds = parseSiteIds(prompt('Allowed siteIds, comma-separated:', getAllowedSiteIds(row).join(', ') || CURRENT_SITE_ID));
  const fullAccess = isOwnerOrSuperAdminRole(role);
  const payload = {
    active: true,
    approvalStatus: 'approved',
    role,
    companyId: getCurrentCompanyId(),
    siteId: siteIds[0],
    siteIds,
    permissions: {
      canEditPunches: fullAccess || row.permissions?.canEditPunches === true,
      canDeletePunches: fullAccess || row.permissions?.canDeletePunches === true,
      canMergeWorkers: fullAccess || row.permissions?.canMergeWorkers === true,
      manageUsers: fullAccess || row.permissions?.manageUsers === true,
    },
    approvedBy: state.me.uid,
    approvedAt: serverTimestamp(),
    updatedAt: serverTimestamp(),
  };

  await setDoc(doc(db, 'users', uid), payload, { merge: true });
  await logAudit('user_approved', 'user', uid, row, payload, 'Pending user approved');
  toast('User approved.');
}

async function denyPendingUser(uid, rows) {
  if (!canManageUsers()) {
    toast('You do not have permission to deny users.', true);
    return;
  }

  const row = rows.find((item) => item.id === uid);
  const payload = {
    active: false,
    approvalStatus: 'denied',
    deniedBy: state.me.uid,
    deniedAt: serverTimestamp(),
    updatedAt: serverTimestamp(),
  };
  await setDoc(doc(db, 'users', uid), payload, { merge: true });
  await logAudit('user_denied', 'user', uid, row || {}, payload, 'Pending user denied');
  toast('User denied/deactivated.');
}

function attachUsersView() {
  const userConstraints = [
    where('companyId', '==', getCurrentCompanyId()),
    where('siteIds', 'array-contains', getCurrentSiteId()),
  ];
  userConstraints.push(orderBy('name', 'asc'));

  const usersQuery = query(collection(db, 'users'), ...userConstraints);

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
              <td>${escapeHtml(row.role || '-')}<br><span class="tiny">${escapeHtml(getAllowedSiteIds(row).join(', ') || '-')}</span></td>
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
    const profileRef = doc(db, 'users', uid);
    const existingProfileSnap = await getDoc(profileRef);
    const existingProfile = existingProfileSnap.exists() ? existingProfileSnap.data() : {};
    const role = els.userRoleInput?.value || existingProfile.role || 'manager';
    const fullAccess = isOwnerOrSuperAdminRole(role);
    const displayName = prettifyHumanName(els.userNameInput?.value.trim()) || existingProfile.displayName || existingProfile.name || '';
    const requestedSiteIds = parseSiteIds(els.userSiteIdsInput?.value || existingProfile.siteIds || existingProfile.siteId || CURRENT_SITE_ID);
    const profilePayload = {
      uid,
      name: displayName,
      displayName,
      email: els.userEmailInput?.value.trim().toLowerCase(),
      role,
      active: els.userActiveInput?.value === 'true',
      agencyId: existingProfile.agencyId || state.agencyId || '',
      companyId: getCurrentCompanyId(),
      siteId: requestedSiteIds[0] || CURRENT_SITE_ID,
      siteIds: requestedSiteIds,
      approvalStatus: els.userActiveInput?.value === 'true' ? 'approved' : (existingProfile.approvalStatus || 'pending'),
      permissions: {
        canEditPunches: fullAccess || existingProfile.permissions?.canEditPunches === true,
        canDeletePunches: fullAccess || existingProfile.permissions?.canDeletePunches === true,
        canMergeWorkers: fullAccess || existingProfile.permissions?.canMergeWorkers === true,
        manageUsers: fullAccess || existingProfile.permissions?.manageUsers === true
      },
      updatedAt: serverTimestamp(),
    };
    if (!existingProfileSnap.exists()) {
      profilePayload.createdAt = serverTimestamp();
    }

    await setDoc(profileRef, profilePayload, { merge: true });
    const roleChanged = existingProfile.role && existingProfile.role !== profilePayload.role;
    const permissionsChanged = JSON.stringify(existingProfile.permissions || {}) !== JSON.stringify(profilePayload.permissions || {});
    await logAudit(
      roleChanged ? 'user_role_changed' : (permissionsChanged ? 'user_permissions_changed' : 'user_profile_saved'),
      'user',
      uid,
      existingProfile,
      profilePayload,
      'Dashboard access updated'
    );

    toast('User profile saved.');
    els.userProfileForm?.reset();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save profile.', true);
  }
}

function populateAgencyWorkerSelect() {
  if (!els.agencyWorkerSelect) return;

  const current = els.agencyWorkerSelect.value;
  const rows = getDerivedTimesheetRows();

  els.agencyWorkerSelect.innerHTML = '<option value="">Select a worker</option>' +
    rows.map((row) => `<option value="${escapeHtml(row.nameKey)}">${escapeHtml(row.name)}</option>`).join('');

  if (rows.some((row) => row.nameKey === current)) {
    els.agencyWorkerSelect.value = current;
  }
}

function renderAgencyPreview() {
  if (!els.agencyPreview || !els.agencyWorkerSelect) return;

  const selectedNameKey = els.agencyWorkerSelect.value;
  if (!selectedNameKey) {
    els.agencyPreview.innerHTML = '<div class="empty-state">Choose a worker and click Preview Sheet.</div>';
    return;
  }

  const weekKey = formatDateKey(state.selectedWeekStart);
  const row = getDerivedTimesheetRows().find((r) => r.nameKey === selectedNameKey && r.weekKey === weekKey);

  if (!row) {
    els.agencyPreview.innerHTML = '<div class="empty-state">No weekly sheet found for that worker.</div>';
    return;
  }

  const signedAt = row.managerSignedAt?.seconds
    ? formatDateTime(row.managerSignedAt.seconds * 1000)
    : '-';

  const dailyRows = buildAgencyDailyRows(row.dailyTotals);

  els.agencyPreview.innerHTML = `
    <div id="agencyPrintableSheet" style="background:#fff;color:#111;border-radius:12px;padding:24px;min-height:200px;">
      <div style="display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap;margin-bottom:18px;">
        <div>
          <h2 style="margin:0 0 8px;font-size:28px;">Weekly Time Sheet</h2>
          <div style="font-size:15px;line-height:1.6;">
            <div><strong>Worker:</strong> ${escapeHtml(row.name)}</div>
            <div><strong>Week Start:</strong> ${escapeHtml(row.weekKey)}</div>
            <div><strong>Status:</strong> ${escapeHtml(row.status)}</div>
          </div>
        </div>
        <div style="font-size:14px;line-height:1.7;text-align:right;">
          <div><strong>Company:</strong> ${escapeHtml(getCompanyName())}</div>
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
    attachUsersViewIfAdmin();
    attachPendingUsersViewIfAdmin();
  }
}

function currentRole() {
  return String(state.profile?.role || '').trim();
}

function hasAnyRole(roles) {
  return roles.includes(currentRole());
}

function hasPermission(permissionName) {
  return state.profile?.permissions?.[permissionName] === true;
}

function isOwnerOrSuperAdminRole(role) {
  return ['owner', 'superAdmin'].includes(String(role || '').trim());
}

function isOwnerOrSuperAdmin() {
  return isOwnerOrSuperAdminRole(currentRole());
}

function isManager() {
  return hasAnyRole(['owner', 'superAdmin', 'admin', 'manager']);
}

function isAdmin() {
  return hasAnyRole(['owner', 'superAdmin', 'admin']);
}

function canEditPunches() {
  if (isOwnerOrSuperAdmin()) return true;
  return hasAnyRole(['admin', 'manager']) && hasPermission('canEditPunches');
}

function canDeletePunches() {
  if (isOwnerOrSuperAdmin()) return true;
  return hasAnyRole(['admin', 'manager']) && hasPermission('canDeletePunches');
}

function canMergeWorkers() {
  if (isOwnerOrSuperAdmin()) return true;
  return hasAnyRole(['admin']) && hasPermission('canMergeWorkers');
}

function canManageUsers() {
  if (isOwnerOrSuperAdmin()) return true;
  return hasAnyRole(['admin']) && hasPermission('manageUsers');
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

function normalizeWorkerNumber(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[^a-z0-9_-]/g, '');
}

function normalizeScopeId(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '_')
    .replace(/[^a-z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function buildStableEmployeeId(employeeNumber, agencyId, siteId) {
  const numberKey = normalizeWorkerNumber(employeeNumber);
  const agencyKey = normalizeScopeId(agencyId || 'direct');
  const siteKey = normalizeScopeId(siteId || 'site');
  return `employee_${agencyKey}_${siteKey}_${numberKey || normalizeScopeId(crypto.randomUUID())}`;
}

function isActiveEmployeeRecord(record) {
  if (!record || typeof record !== 'object') return false;
  if (record.active === false) return false;
  const status = String(record.status || 'active').trim().toLowerCase();
  return !['inactive', 'removed', 'terminated', 'disabled', 'archived', 'merged'].includes(status);
}

function isActivePunchRecord(record) {
  if (!record || typeof record !== 'object') return false;
  if (record.active === false) return false;
  const status = String(record.status || '').trim().toLowerCase();
  return status !== 'deleted';
}

function normalizeUserProfile(profile = {}, authUser = null) {
  const status = String(profile.status || '').trim().toLowerCase();
  const role = profile.role === 'worker' ? 'worker' : (profile.role || 'manager');
  const fullAccess = isOwnerOrSuperAdminRole(role);
  const active = typeof profile.active === 'boolean' ? profile.active : status !== 'inactive';
  const displayName = prettifyHumanName(
    profile.displayName
    || profile.name
    || [profile.firstName, profile.lastName].filter(Boolean).join(' ')
    || authUser?.displayName
    || profile.email
    || authUser?.email
    || ''
  );

  return {
    ...profile,
    uid: profile.uid || authUser?.uid || '',
    name: displayName || profile.email || authUser?.email || 'Signed in',
    displayName: profile.displayName || displayName || profile.email || authUser?.email || 'Signed in',
    email: String(profile.email || authUser?.email || '').trim().toLowerCase(),
    role,
    active,
    companyId: profile.companyId || '',
    agencyId: profile.agencyId || '',
    siteId: profile.siteId || profile.assignedSiteId || '',
    siteIds: parseSiteIds(profile.siteIds || profile.siteId || profile.assignedSiteId || '', false),
    permissions: {
      canEditPunches: fullAccess || profile.permissions?.canEditPunches === true,
      canDeletePunches: fullAccess || profile.permissions?.canDeletePunches === true,
      canMergeWorkers: fullAccess || profile.permissions?.canMergeWorkers === true,
      manageUsers: fullAccess || profile.permissions?.manageUsers === true
    },
    employeeId: profile.employeeId || profile.workerId || '',
    workerId: profile.workerId || profile.employeeId || ''
  };
}

async function findExistingEmployeeForUpsert({ employeeNumber, nameKey, companyId = '', agencyId = '', siteId = '' }) {
  const employeesRef = collection(db, 'employees');
  const matches = [];
  const employeeNumberKey = normalizeWorkerNumber(employeeNumber);
  const normalizedSiteId = String(siteId || '').trim();
  const normalizedAgencyId = String(agencyId || '').trim();
  const normalizedCompanyId = String(companyId || '').trim();

  const addMatches = (snap) => {
    snap.docs.forEach((record) => {
      const row = { id: record.id, ...record.data() };
      if (!isActiveEmployeeRecord(row)) return;
      const sameCompany = String(row.companyId || '').trim() === normalizedCompanyId;
      const sameAgency = String(row.agencyId || '').trim() === normalizedAgencyId;
      const sameSite = String(row.assignedSiteId || row.siteId || '').trim() === normalizedSiteId;
      if (sameCompany && sameAgency && sameSite) matches.push(row);
    });
  };

  if (employeeNumberKey) {
    addMatches(await getDocs(query(
      employeesRef,
      where('companyId', '==', normalizedCompanyId),
      where('agencyId', '==', normalizedAgencyId),
      where('assignedSiteId', '==', normalizedSiteId),
      where('status', '==', 'active'),
      where('employeeNumberKey', '==', employeeNumberKey)
    )));
    if (!matches.length) {
      addMatches(await getDocs(query(
        employeesRef,
        where('companyId', '==', normalizedCompanyId),
        where('agencyId', '==', normalizedAgencyId),
        where('assignedSiteId', '==', normalizedSiteId),
        where('status', '==', 'active'),
        where('employeeNumber', '==', employeeNumber)
      )));
    }
  }

  if (!matches.length && nameKey) {
    addMatches(await getDocs(query(
      employeesRef,
      where('companyId', '==', normalizedCompanyId),
      where('agencyId', '==', normalizedAgencyId),
      where('assignedSiteId', '==', normalizedSiteId),
      where('status', '==', 'active'),
      where('nameKey', '==', nameKey)
    )));
  }

  return matches[0] || null;
}

async function logAudit(action, entityType, entityId, oldValue = {}, newValue = {}, reason = '') {
  if (!state.me) return;
  try {
    await addDoc(collection(db, 'auditLogs'), {
      ...branchPayload(),
      agencyId: state.agencyId || '',
      userId: state.me.uid || '',
      actorId: state.me.uid || '',
      actorRole: state.profile?.role || '',
      role: state.profile?.role || '',
      action,
      eventType: action,
      entityType,
      entityId,
      affectedRecord: entityId,
      oldValue,
      newValue,
      reason,
      timestamp: serverTimestamp(),
      createdAt: serverTimestamp()
    });
  } catch (error) {
    console.warn('Audit log write failed:', error.message);
  }
}

async function loadPunchesForEmployee(employee) {
  const ids = [...new Set([employee?.id, employee?.employeeId, employee?.workerId].filter(Boolean))];
  const snapshots = [];
  for (const id of ids) {
    snapshots.push(await getDocs(query(collection(db, 'punches'), ...branchConstraints(), where('employeeId', '==', id))));
    snapshots.push(await getDocs(query(collection(db, 'punches'), ...branchConstraints(), where('workerId', '==', id))));
  }
  const rows = new Map();
  snapshots.forEach((snap) => {
    snap.docs.forEach((record) => {
      rows.set(record.id, { id: record.id, ...record.data() });
    });
  });
  return [...rows.values()];
}

async function previewDuplicateEmployeeMerge(primaryEmployeeId, duplicateEmployeeIds, options = {}) {
  const [primarySnap, ...duplicateSnaps] = await Promise.all([
    getDoc(doc(db, 'employees', primaryEmployeeId)),
    ...duplicateEmployeeIds.map((id) => getDoc(doc(db, 'employees', id)))
  ]);
  if (!primarySnap.exists()) throw new Error('Primary employee not found.');
  const primary = { id: primarySnap.id, ...primarySnap.data() };
  const duplicates = duplicateSnaps.map((snap) => {
    if (!snap.exists()) throw new Error('One or more duplicate employees were not found.');
    return { id: snap.id, ...snap.data() };
  });

  const conflicts = [];
  duplicates.forEach((duplicate) => {
    if (normalizeWorkerNumber(primary.employeeNumber) && normalizeWorkerNumber(duplicate.employeeNumber) && normalizeWorkerNumber(primary.employeeNumber) !== normalizeWorkerNumber(duplicate.employeeNumber)) {
      conflicts.push(`employeeNumber: ${duplicate.id}`);
    }
    if (String(primary.agencyId || '') !== String(duplicate.agencyId || '')) {
      conflicts.push(`agencyId: ${duplicate.id}`);
    }
    if (String(primary.assignedSiteId || primary.siteId || '') !== String(duplicate.assignedSiteId || duplicate.siteId || '')) {
      conflicts.push(`siteId: ${duplicate.id}`);
    }
  });

  if (conflicts.length && !(options.overrideConflicts === true && isOwnerOrSuperAdmin())) {
    return { ok: false, blocked: true, conflicts, primary, duplicates, punchCount: 0, dateRange: null };
  }

  const punchRows = (await Promise.all(duplicates.map(loadPunchesForEmployee))).flat();
  const uniquePunches = [...new Map(punchRows.map((row) => [row.id, row])).values()];
  const timestamps = uniquePunches.map((row) => Number(row.timestampMs || 0)).filter(Boolean).sort((a, b) => a - b);

  return {
    ok: true,
    dryRun: true,
    primary,
    duplicates,
    punchCount: uniquePunches.length,
    dateRange: timestamps.length ? { start: formatDateKey(new Date(timestamps[0])), end: formatDateKey(new Date(timestamps[timestamps.length - 1])) } : null,
    agency: primary.agencyId || '',
    site: primary.assignedSiteId || primary.siteId || '',
    employeeNumber: primary.employeeNumber || '',
    punches: uniquePunches
  };
}

async function mergeDuplicateEmployees(primaryEmployeeId, duplicateEmployeeIds, options = {}) {
  if (!canMergeWorkers()) throw new Error('You do not have permission to merge workers.');
  const preview = await previewDuplicateEmployeeMerge(primaryEmployeeId, duplicateEmployeeIds, options);
  if (!preview.ok) return preview;
  if (options.dryRun !== false) return preview;

  const now = serverTimestamp();
  const batches = [];
  let batch = writeBatch(db);
  let writes = 0;
  const commitWhenFull = () => {
    if (writes < 450) return;
    batches.push(batch.commit());
    batch = writeBatch(db);
    writes = 0;
  };

  preview.punches.forEach((punch) => {
    batch.update(doc(db, 'punches', punch.id), {
      employeeId: preview.primary.id,
      workerId: preview.primary.workerId || preview.primary.employeeId || preview.primary.id,
      name: preview.primary.name || punch.name || '',
      nameKey: preview.primary.nameKey || normalizeName(preview.primary.name || punch.name || ''),
      employeeNumber: preview.primary.employeeNumber || punch.employeeNumber || '',
      mergedFromEmployeeId: punch.employeeId || punch.workerId || '',
      mergePreservedOriginal: {
        employeeId: punch.employeeId || '',
        workerId: punch.workerId || '',
        name: punch.name || '',
        nameKey: punch.nameKey || '',
        employeeNumber: punch.employeeNumber || ''
      },
      updatedAt: now
    });
    writes += 1;
    commitWhenFull();
  });

  preview.duplicates.forEach((duplicate) => {
    batch.update(doc(db, 'employees', duplicate.id), {
      active: false,
      status: 'merged',
      mergedInto: preview.primary.id,
      mergedAt: now,
      mergedBy: state.me.uid || '',
      updatedAt: now
    });
    writes += 1;
    commitWhenFull();
  });

  const mergeLogRef = doc(collection(db, 'mergeLogs'));
  batch.set(mergeLogRef, {
    ...branchPayload(preview.site || getCurrentSiteId()),
    primaryWorkerId: preview.primary.id,
    duplicateWorkerIds: preview.duplicates.map((worker) => worker.id),
    punchCount: preview.punchCount,
    dateRange: preview.dateRange,
    agencyId: preview.agency,
    siteId: preview.site || getCurrentSiteId(),
    employeeNumber: preview.employeeNumber,
    mergedBy: state.me.uid || '',
    mergedByName: state.profile?.name || state.me?.email || '',
    overrideConflicts: options.overrideConflicts === true,
    timestamp: now,
    createdAt: now
  });
  writes += 1;

  if (writes) batches.push(batch.commit());
  await Promise.all(batches);
  await logAudit('worker_merge_completed', 'worker', preview.primary.id, {
    duplicateWorkerIds: preview.duplicates.map((worker) => worker.id)
  }, {
    primaryWorkerId: preview.primary.id,
    movedPunches: preview.punchCount,
    mergeLogId: mergeLogRef.id
  }, 'Safe duplicate employee merge');

  return { ...preview, dryRun: false, mergeLogId: mergeLogRef.id };
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

function parseLocalDateAndTime(dateValue, timeValue) {
  if (!dateValue || !timeValue) return 0;
  const cleaned = `${dateValue} ${timeValue}`;
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

function formatDateOnly(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleDateString([], {
    year: 'numeric',
    month: 'short',
    day: 'numeric'
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

function formatTimeForInput(ms) {
  const d = new Date(ms);
  const hour = String(d.getHours()).padStart(2, '0');
  const minute = String(d.getMinutes()).padStart(2, '0');
  return `${hour}:${minute}`;
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

function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

window.QRTimeClockAdminTools = {
  previewDuplicateEmployeeMerge,
  mergeDuplicateEmployees
};
