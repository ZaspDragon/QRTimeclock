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
  deleteDoc,
  serverTimestamp,
  query,
  where,
  orderBy,
  limit,
  onSnapshot,
  Timestamp,
  writeBatch
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const state = {
  me: null,
  profile: null,
  companyId: null,        // from user profile
  agencyId: null,         // from user profile (null = direct company user)
  companyDoc: null,       // loaded from companies/{companyId}
  unsubscribers: [],
  selectedWeekStart: getMondayDate(new Date()),
  workerUnsub: null,
  workerEmployee: null,      // looked-up employee record for public punch
  publicEmployees: [],       // cached employee list for public QR autocomplete
  publicEmployeeRecords: [], // full active list retained for historical ID lookup
  allPunchRows: [],
  selectedWeekPunchRows: [],
  selectedWeekTimesheetDocs: {},
  allEmployees: [],
  allMissedRequests: [],
  approvalFilter: 'pending',
  workerPunchSaving: false,
  employeeStatusFilter: 'active',
  duplicateGroups: [],
  managerTimeLookup: null,
};

const els = {
  workerNameInput: document.getElementById('workerNameInput'),
  workerAutocompleteList: document.getElementById('workerAutocompleteList'),
  workerLookupStatus: document.getElementById('workerLookupStatus'),
  workerNameValue: document.getElementById('workerNameValue'),
  workerLastActionValue: document.getElementById('workerLastActionValue'),
  workerLastPunchValue: document.getElementById('workerLastPunchValue'),
  workerStatusValue: document.getElementById('workerStatusValue'),
  workerStatusMessage: document.getElementById('workerStatusMessage'),
  workerHistoryBody: document.getElementById('workerHistoryBody'),
  workerViewTimeBtn: document.getElementById('workerViewTimeBtn'),
  workerViewMoreTimeBtn: document.getElementById('workerViewMoreTimeBtn'),
  workerRequestFixBtn: document.getElementById('workerRequestFixBtn'),
  workerMyTimePanel: document.getElementById('workerMyTimePanel'),
  workerFixPanel: document.getElementById('workerFixPanel'),
  workerTimeRangeControls: document.getElementById('workerTimeRangeControls'),
  workerTimeFromInput: document.getElementById('workerTimeFromInput'),
  workerTimeToInput: document.getElementById('workerTimeToInput'),
  workerTimeLookupBtn: document.getElementById('workerTimeLookupBtn'),
  workerTimeRangeStatus: document.getElementById('workerTimeRangeStatus'),
  workerTimeRangeResults: document.getElementById('workerTimeRangeResults'),
  workerWeekHoursValue: document.getElementById('workerWeekHoursValue'),
  workerRegularHoursValue: document.getElementById('workerRegularHoursValue'),
  workerOvertimeHoursValue: document.getElementById('workerOvertimeHoursValue'),
  workerDaysWorkedValue: document.getElementById('workerDaysWorkedValue'),
  workerFixForm: document.getElementById('workerFixForm'),
  workerFixActionInput: document.getElementById('workerFixActionInput'),
  workerFixDateInput: document.getElementById('workerFixDateInput'),
  workerFixTimeInput: document.getElementById('workerFixTimeInput'),
  workerFixReasonInput: document.getElementById('workerFixReasonInput'),
  workerPinField: document.getElementById('workerPinField'),
  workerPinInput: document.getElementById('workerPinInput'),

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
  userListBody: document.getElementById('userListBody'),

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
  empPinInput: document.getElementById('empPinInput'),
  empCancelEditBtn: document.getElementById('empCancelEditBtn'),
  empFilterInput: document.getElementById('empFilterInput'),
  empRosterStatusFilter: document.getElementById('empRosterStatusFilter'),
  employeeListBody: document.getElementById('employeeListBody'),
  inactiveWorkerListBody: document.getElementById('inactiveWorkerListBody'),
  exportBackupBtn: document.getElementById('exportBackupBtn'),
  refreshDuplicatesBtn: document.getElementById('refreshDuplicatesBtn'),
  duplicateWorkersList: document.getElementById('duplicateWorkersList'),
  permissionDebugPanel: document.getElementById('permissionDebugPanel'),
  managerTimeWorkerSelect: document.getElementById('managerTimeWorkerSelect'),
  managerTimeFromInput: document.getElementById('managerTimeFromInput'),
  managerTimeToInput: document.getElementById('managerTimeToInput'),
  managerTimeLookupBtn: document.getElementById('managerTimeLookupBtn'),
  managerTimeExportBtn: document.getElementById('managerTimeExportBtn'),
  managerTimeTotalValue: document.getElementById('managerTimeTotalValue'),
  managerTimeRegularValue: document.getElementById('managerTimeRegularValue'),
  managerTimeOvertimeValue: document.getElementById('managerTimeOvertimeValue'),
  managerTimeDaysValue: document.getElementById('managerTimeDaysValue'),
  managerTimeRangeStatus: document.getElementById('managerTimeRangeStatus'),
  managerTimeRangeResults: document.getElementById('managerTimeRangeResults'),

  agencyWorkerSelect: document.getElementById('agencyWorkerSelect'),
  agencyPreviewBtn: document.getElementById('agencyPreviewBtn'),
  agencyPrintBtn: document.getElementById('agencyPrintBtn'),
  agencyPreview: document.getElementById('agencyPreview'),

  toast: document.getElementById('toast'),
};

init();

async function init() {
  wireEvents();

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

  if (els.workerFixDateInput) {
    els.workerFixDateInput.value = formatDateInput(new Date());
  }

  if (els.workerFixTimeInput) {
    els.workerFixTimeInput.value = formatTimeForInput(Date.now());
  }
  applyQuickDateRange('this_week', els.workerTimeFromInput, els.workerTimeToInput);
  applyQuickDateRange('this_week', els.managerTimeFromInput, els.managerTimeToInput);

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
      state.companyId = state.profile.companyId || null;
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
      renderPermissionDebug();
      if (canEditPunches()) {
        attachManagerLiveViews();
        attachTimesheetView();
        attachUsersViewIfAdmin();
        populateAgencyWorkerSelect();
        renderAgencyPreview();
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

  // Name input with autocomplete
  els.workerNameInput?.addEventListener('input', debounce(handleWorkerNameAutocomplete, 250));
  els.workerNameInput?.addEventListener('focus', () => handleWorkerNameAutocomplete());
  els.workerNameInput?.addEventListener('keydown', handleAutocompleteKeydown);
  els.workerViewTimeBtn?.addEventListener('click', showWorkerTimeThisWeek);
  els.workerViewMoreTimeBtn?.addEventListener('click', showWorkerMoreTime);
  els.workerTimeLookupBtn?.addEventListener('click', lookupPublicWorkerTimeRange);
  document.querySelectorAll('.worker-range-quick').forEach((button) => {
    button.addEventListener('click', () => {
      applyQuickDateRange(button.dataset.range, els.workerTimeFromInput, els.workerTimeToInput);
      lookupPublicWorkerTimeRange();
    });
  });
  els.workerRequestFixBtn?.addEventListener('click', () => toggleWorkerSelfService('fix'));
  els.workerFixForm?.addEventListener('submit', handlePublicTimeFixSubmit);

  // Close autocomplete on outside click
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.autocomplete-wrap')) {
      hideAutocomplete();
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
  els.empRosterStatusFilter?.addEventListener('change', () => {
    state.employeeStatusFilter = els.empRosterStatusFilter.value || 'active';
    renderEmployeeList(state.allEmployees || []);
  });
  els.exportBackupBtn?.addEventListener('click', exportBackup);
  els.refreshDuplicatesBtn?.addEventListener('click', () => renderDuplicateWorkers(true));
  els.managerTimeLookupBtn?.addEventListener('click', lookupManagerTimeRange);
  els.managerTimeExportBtn?.addEventListener('click', exportManagerTimeRangeCsv);
  document.querySelectorAll('.manager-range-quick').forEach((button) => {
    button.addEventListener('click', () => {
      applyQuickDateRange(button.dataset.range, els.managerTimeFromInput, els.managerTimeToInput);
      if (els.managerTimeWorkerSelect?.value) lookupManagerTimeRange();
    });
  });

  els.agencyPreviewBtn?.addEventListener('click', () => renderAgencyPreview());
  els.agencyPrintBtn?.addEventListener('click', () => printAgencyPreview());
  els.agencyWorkerSelect?.addEventListener('change', () => renderAgencyPreview());
}

// ─── Public employee loading & autocomplete ─────────────
async function loadPublicEmployees() {
  try {
    // Legacy employee records use a blank companyId. Status is the stable,
    // backwards-compatible public clock-in scope.
    const q = query(collection(db, 'employees'), where('status', '==', 'active'));
    const snap = await getDocs(q);
    const activeEmployees = snap.docs
      .map((d) => ({ id: d.id, ...d.data() }))
      .sort((a, b) => (a.name || '').localeCompare(b.name || ''));
    state.publicEmployeeRecords = activeEmployees;
    state.publicEmployees = collapseDuplicateEmployees(activeEmployees);
  } catch (error) {
    console.warn('Could not load employees for autocomplete:', error.message);
    state.publicEmployeeRecords = [];
    state.publicEmployees = [];
  }
}

function restoreWorkerFromName(name) {
  const nameKey = normalizeName(name);
  const matches = state.publicEmployees.filter((e) => normalizeName(e.name) === nameKey);
  if (matches.length === 1) {
    const match = matches[0];
    state.workerEmployee = match;
    updateWorkerPinField(match);
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
    updateWorkerPinField(null);
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
  const exactMatches = state.publicEmployees.filter(
    (e) => normalizeName(e.name) === normalizeName(typed)
  );

  if (exactMatches.length === 1) {
    const exactMatch = exactMatches[0];
    state.workerEmployee = exactMatch;
    updateWorkerPinField(exactMatch);
    if (els.workerLookupStatus) {
      els.workerLookupStatus.textContent = `✓ Found: ${exactMatch.name}. Ready to punch.`;
      els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
    }
    hideAutocomplete();
    localStorage.setItem('workerPunchName', exactMatch.name);
    attachWorkerLiveView(exactMatch.name);
    return;
  }

  if (exactMatches.length > 1) {
    state.workerEmployee = null;
    updateWorkerPinField(null);
    if (els.workerLookupStatus) {
      els.workerLookupStatus.textContent = 'More than one worker has that name. Select your name from the list.';
    }
    renderAutocomplete(exactMatches, typed, true);
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

function normalizeIdentityPart(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function employeeScopeKey(employee) {
  return [
    normalizeIdentityPart(employee?.name || employee?.nameKey),
    normalizeIdentityPart(employee?.agencyId),
    normalizeIdentityPart(employee?.assignedSiteId || employee?.siteId),
  ].join('|');
}

function employeeRecordScore(employee) {
  let score = 0;
  if (normalizeIdentityPart(employee?.agencyId)) score += 20;
  if (normalizeIdentityPart(employee?.assignedSiteId || employee?.siteId)) score += 20;
  const employeeNumber = normalizeIdentityPart(employee?.employeeNumber);
  if (employeeNumber && employeeNumber !== 'emp-1001') score += 10;
  if (normalizeIdentityPart(employee?.employeeId) === normalizeIdentityPart(employee?.id)) score += 2;
  return score;
}

function compareEmployeeRecords(left, right) {
  const scoreDifference = employeeRecordScore(right) - employeeRecordScore(left);
  if (scoreDifference) return scoreDifference;
  const leftCreated = Number(left.createdAt?.seconds || left.createdAtMs || 0);
  const rightCreated = Number(right.createdAt?.seconds || right.createdAtMs || 0);
  return leftCreated - rightCreated || String(left.id).localeCompare(String(right.id));
}

function preferEmployeeRecord(employees) {
  return [...employees].sort(compareEmployeeRecords)[0];
}

function collapseDuplicateEmployees(employees) {
  const kept = [];
  const nameGroups = new Map();

  employees.forEach((employee) => {
    const name = normalizeIdentityPart(employee.name || employee.nameKey);
    if (!nameGroups.has(name)) nameGroups.set(name, []);
    nameGroups.get(name).push(employee);
  });

  nameGroups.forEach((nameEmployees) => {
    const scopedEmployees = new Map();
    const nonBlankScopes = new Set();
    nameEmployees.forEach((employee) => {
      const agency = normalizeIdentityPart(employee.agencyId);
      const site = normalizeIdentityPart(employee.assignedSiteId || employee.siteId);
      const scope = `${agency}|${site}`;
      if (!scopedEmployees.has(scope)) scopedEmployees.set(scope, []);
      scopedEmployees.get(scope).push(employee);
      if (agency || site) nonBlankScopes.add(scope);
    });

    // Legacy blank copies belong to the sole configured agency/site record.
    if (nonBlankScopes.size <= 1) {
      kept.push(preferEmployeeRecord(nameEmployees));
      return;
    }

    scopedEmployees.forEach((scopeEmployees) => kept.push(preferEmployeeRecord(scopeEmployees)));
  });

  return kept.sort((left, right) => String(left.name || '').localeCompare(String(right.name || '')));
}

function findReusableEmployee({ name, employeeNumber = '', agencyId = '', assignedSiteId = '' }, employees) {
  const nameKey = normalizeIdentityPart(name);
  const numberKey = normalizeIdentityPart(employeeNumber);
  const scopeKey = employeeScopeKey({ name, agencyId, assignedSiteId });
  const activeEmployees = (employees || []).filter((employee) => {
    if (!isActiveEmployee(employee)) return false;
    return normalizeIdentityPart(employee.name || employee.nameKey) === nameKey;
  });
  const scopedMatch = activeEmployees.find((employee) => employeeScopeKey(employee) === scopeKey);
  if (scopedMatch) return scopedMatch;
  return activeEmployees.find((employee) => {
    const existingNumber = normalizeIdentityPart(employee.employeeNumber);
    return numberKey && existingNumber === numberKey;
  }) || null;
}

function renderAutocomplete(matches, typed, selectionRequired = false) {
  const list = els.workerAutocompleteList;
  if (!list) return;

  _acActiveIndex = -1;
  let html = '';

  matches.forEach((emp, i) => {
    html += `<li data-index="${i}" data-emp-id="${emp.id}">
      ${escapeHTML(emp.name)}
    </li>`;
  });

  // Always show "new worker" option at the bottom if typed name doesn't match
  if (typed.length >= 2 && !selectionRequired && matches.length === 0) {
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
  updateWorkerPinField(emp);
  if (els.workerNameInput) els.workerNameInput.value = emp.name;
  if (els.workerNameValue) els.workerNameValue.textContent = emp.name;
  if (els.workerLookupStatus) {
    els.workerLookupStatus.textContent = `✓ Found: ${emp.name}. Ready to punch.`;
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
  updateWorkerPinField(null);
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
  if (state.workerPunchSaving) return;
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

  if (['inactive', 'terminated', 'removed'].includes(String(emp.status || '').toLowerCase())) {
    toast('Your employee record is not active. Contact your manager.', true);
    return;
  }

  const enteredPin = String(els.workerPinInput?.value || '').trim();
  if (enteredPin && emp.pinHash) {
    const enteredHash = await hashWorkerPin(enteredPin, emp.id || emp.employeeId || '');
    if (enteredHash !== emp.pinHash) {
      toast('PIN incorrect. You can clear it and continue with the normal name-based flow.', true);
      return;
    }
  }

  const urlCompanyId = new URLSearchParams(window.location.search).get('company') || '';

  // Auto-create employee if new
  if (emp._isNew) {
    const reusable = findReusableEmployee({
      name: emp.name,
      agencyId: '',
      assignedSiteId: '',
    }, state.publicEmployees);
    if (reusable) {
      emp = reusable;
      state.workerEmployee = reusable;
    }
  }

  if (emp._isNew) {
    try {
      const identityHash = await hashIdentityKey(employeeScopeKey(emp));
      const empNumber = `AUTO-${identityHash.slice(0, 8).toUpperCase()}`;
      const newRef = doc(db, 'employees', `auto_${identityHash.slice(0, 24)}`);
      const existingSnap = await getDoc(newRef);
      if (existingSnap.exists()) {
        emp = { id: existingSnap.id, ...existingSnap.data() };
        state.workerEmployee = emp;
      } else {
        const newPayload = {
          name: emp.name,
          nameKey: normalizeName(emp.name),
          employeeNumber: empNumber,
          companyId: urlCompanyId || '',
          agencyId: '',
          assignedSiteId: '',
          status: 'active',
          source: 'auto_created',
          createdAt: serverTimestamp(),
          updatedAt: serverTimestamp(),
        };
        await setDoc(newRef, { ...newPayload, employeeId: newRef.id });
        emp = { id: newRef.id, employeeId: newRef.id, ...newPayload };
        state.workerEmployee = emp;
        state.publicEmployees = collapseDuplicateEmployees([...state.publicEmployees, emp]);
        if (els.workerLookupStatus) {
          els.workerLookupStatus.textContent = `Found: ${emp.name}. Punching...`;
          els.workerLookupStatus.style.borderColor = 'rgba(43,213,118,0.4)';
        }
      }
    } catch (error) {
      console.warn('Auto-create employee skipped:', error.message);
      const identityHash = await hashIdentityKey(employeeScopeKey(emp));
      const existingSnap = await getDoc(doc(db, 'employees', `auto_${identityHash.slice(0, 24)}`));
      emp = existingSnap.exists()
        ? { id: existingSnap.id, ...existingSnap.data() }
        : { name: emp.name, nameKey: normalizeName(emp.name), employeeId: '', employeeNumber: '' };
      state.workerEmployee = emp;
    }
  }

  const name = emp.name || '';
  const nameKey = normalizeName(name);
  const now = new Date();
  const nowMs = Date.now();
  const dateKey = formatDateKey(now);
  const weekKey = formatDateKey(getMondayDate(now));

  state.workerPunchSaving = true;
  setWorkerPunchBusy(true);
  try {
    const duplicateKey = `lastPunch:${emp.employeeId || emp.id || nameKey}:${action}`;
    const previousPunchMs = Number(localStorage.getItem(duplicateKey) || 0);
    if (previousPunchMs && nowMs - previousPunchMs < 10000) {
      throw new Error('That punch was already saved. Please wait a few seconds.');
    }
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
      companyId: emp.companyId || urlCompanyId || '',
      agencyId: emp.agencyId || '',
    });

    cacheWorkerPunch(emp, {
      name,
      nameKey,
      action,
      timestampMs: nowMs,
      dateKey,
      weekKey,
      employeeId: emp.employeeId || emp.id || '',
    });

    if (els.workerLastActionValue) els.workerLastActionValue.textContent = prettyAction(action);
    if (els.workerLastPunchValue) els.workerLastPunchValue.textContent = formatDateTime(nowMs);
    if (els.workerStatusValue) els.workerStatusValue.textContent = statusLabelForAction(action);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = `${prettyAction(action)} saved for ${name} at ${formatDateTime(nowMs)}.`;
    }

    attachWorkerLiveView(name);
    localStorage.setItem('workerPunchName', name);
    localStorage.setItem(duplicateKey, String(nowMs));
    toast(`${prettyAction(action)} saved.`);
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save punch.', true);
  } finally {
    state.workerPunchSaving = false;
    setWorkerPunchBusy(false);
  }
}

function setWorkerPunchBusy(busy) {
  document.querySelectorAll('.worker-action-btn').forEach((button) => {
    button.disabled = busy;
    button.setAttribute('aria-busy', busy ? 'true' : 'false');
  });
}

function updateWorkerPinField(employee) {
  const available = Boolean(employee?.pinHash);
  els.workerPinField?.classList.toggle('hidden', !available);
  if (!available && els.workerPinInput) els.workerPinInput.value = '';
}

function namesAreSimilar(left, right) {
  const a = normalizeName(left);
  const b = normalizeName(right);
  return a === b || (a.length >= 4 && b.length >= 4 && (a.includes(b) || b.includes(a)));
}

async function hashWorkerPin(pin, employeeId) {
  const bytes = new TextEncoder().encode(`${employeeId}:${String(pin).trim()}`);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

async function hashIdentityKey(value) {
  const bytes = new TextEncoder().encode(String(value || ''));
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, '0')).join('');
}

function getWorkerCacheKey(employee) {
  return `qrTimeclockWorkerPunches:name:${normalizeName(employee?.name || '')}`;
}

function getLegacyWorkerCacheKey(employee) {
  const identity = employee?.employeeId || employee?.id || '';
  return identity ? `qrTimeclockWorkerPunches:${identity}` : '';
}

function getCachedWorkerPunches(employee) {
  if (!employee?.name) return [];
  try {
    const nameRows = JSON.parse(localStorage.getItem(getWorkerCacheKey(employee)) || '[]');
    const legacyKey = getLegacyWorkerCacheKey(employee);
    const legacyRows = legacyKey ? JSON.parse(localStorage.getItem(legacyKey) || '[]') : [];
    const combined = [...(Array.isArray(legacyRows) ? legacyRows : []), ...(Array.isArray(nameRows) ? nameRows : [])];
    const unique = new Map();
    combined.forEach((row) => {
      const key = `${row.timestampMs || 0}:${row.action || ''}:${row.nameKey || normalizeName(row.name || '')}`;
      unique.set(key, row);
    });
    return [...unique.values()].sort((left, right) => Number(left.timestampMs || 0) - Number(right.timestampMs || 0));
  } catch (_) {
    return [];
  }
}

function cacheWorkerPunch(employee, punch) {
  const rows = getCachedWorkerPunches(employee);
  rows.push(punch);
  rows.sort((left, right) => Number(left.timestampMs || 0) - Number(right.timestampMs || 0));
  localStorage.setItem(getWorkerCacheKey(employee), JSON.stringify(rows.slice(-250)));
}

function getSelectedWorkerByName() {
  const typedName = prettifyHumanName(els.workerNameInput?.value.trim() || '');
  if (
    state.workerEmployee?.name &&
    !state.workerEmployee._isNew &&
    normalizeName(state.workerEmployee.name) === normalizeName(typedName)
  ) {
    return state.workerEmployee;
  }
  const exactMatches = state.publicEmployees.filter(
    (employee) => normalizeName(employee.name) === normalizeName(typedName)
  );
  if (exactMatches.length === 1) {
    state.workerEmployee = exactMatches[0];
    return exactMatches[0];
  }
  if (exactMatches.length > 1) {
    renderAutocomplete(exactMatches, typedName, true);
    toast('Select your name from the list first.', true);
  } else {
    toast('Select your existing worker name first.', true);
  }
  return null;
}

function toggleWorkerSelfService(panel) {
  els.workerMyTimePanel?.classList.toggle('hidden', panel !== 'time');
  els.workerFixPanel?.classList.toggle('hidden', panel !== 'fix');
}

async function showWorkerTimeThisWeek() {
  const employee = getSelectedWorkerByName();
  if (!employee) return;

  toggleWorkerSelfService('time');
  els.workerTimeRangeControls?.classList.add('hidden');
  applyQuickDateRange('this_week', els.workerTimeFromInput, els.workerTimeToInput);
  await lookupPublicWorkerTimeRange(employee);
}

async function showWorkerMoreTime() {
  const employee = getSelectedWorkerByName();
  if (!employee) return;
  toggleWorkerSelfService('time');
  els.workerTimeRangeControls?.classList.remove('hidden');
  if (!els.workerTimeFromInput?.value || !els.workerTimeToInput?.value) {
    applyQuickDateRange('last_2_weeks', els.workerTimeFromInput, els.workerTimeToInput);
  }
  await lookupPublicWorkerTimeRange(employee);
}

async function lookupPublicWorkerTimeRange(selectedEmployee = null) {
  const employee = selectedEmployee || getSelectedWorkerByName();
  if (!employee) return;
  const range = readDateRange(els.workerTimeFromInput, els.workerTimeToInput);
  if (!range) return;

  toggleWorkerSelfService('time');
  setTimeRangeStatus(els.workerTimeRangeStatus, `Loading time for ${employee.name}...`);
  try {
    const officialPunches = await loadPunchesForEmployeeRange(employee, range.fromMs, range.toMs, {
      allowLegacyNameFallback: true,
    });
    const cachedPunches = getCachedWorkerPunches(employee)
      .filter((punch) => Number(punch.timestampMs || 0) >= range.fromMs && Number(punch.timestampMs || 0) <= range.toMs);
    const totals = buildTimeRangeTotals(dedupePunches([...officialPunches, ...cachedPunches]));
    renderTimeRangeSummary({
      totals,
      totalElement: els.workerWeekHoursValue,
      regularElement: els.workerRegularHoursValue,
      overtimeElement: els.workerOvertimeHoursValue,
      daysElement: els.workerDaysWorkedValue,
      statusElement: els.workerTimeRangeStatus,
      resultsElement: els.workerTimeRangeResults,
      emptyMessage: 'No punches were found for this date range.',
    });
  } catch (error) {
    console.warn('Official worker time lookup unavailable:', error.message);
    const cachedPunches = getCachedWorkerPunches(employee)
      .filter((punch) => Number(punch.timestampMs || 0) >= range.fromMs && Number(punch.timestampMs || 0) <= range.toMs);
    const totals = buildTimeRangeTotals(cachedPunches);
    renderTimeRangeSummary({
      totals,
      totalElement: els.workerWeekHoursValue,
      regularElement: els.workerRegularHoursValue,
      overtimeElement: els.workerOvertimeHoursValue,
      daysElement: els.workerDaysWorkedValue,
      statusElement: els.workerTimeRangeStatus,
      resultsElement: els.workerTimeRangeResults,
      emptyMessage: 'Official history could not load. No saved punches from this device were found for the range.',
      statusPrefix: 'Official history is temporarily unavailable. Showing device history only. ',
    });
  }
}

function applyQuickDateRange(rangeName, fromInput, toInput) {
  if (!fromInput || !toInput) return;
  const today = startOfLocalDay(new Date());
  let fromDate = new Date(today);
  let toDate = new Date(today);

  if (rangeName === 'this_week') {
    fromDate = getMondayDate(today);
    toDate = addLocalDays(fromDate, 6);
  } else if (rangeName === 'last_week') {
    fromDate = addLocalDays(getMondayDate(today), -7);
    toDate = addLocalDays(fromDate, 6);
  } else if (rangeName === 'last_2_weeks') {
    fromDate = addLocalDays(today, -13);
  } else if (rangeName === 'this_month') {
    fromDate = new Date(today.getFullYear(), today.getMonth(), 1);
  }

  fromInput.value = formatDateInput(fromDate);
  toInput.value = formatDateInput(toDate);
}

function startOfLocalDay(date) {
  const result = new Date(date);
  result.setHours(0, 0, 0, 0);
  return result;
}

function addLocalDays(date, days) {
  const result = new Date(date);
  result.setDate(result.getDate() + days);
  return result;
}

function readDateRange(fromInput, toInput) {
  const fromValue = fromInput?.value || '';
  const toValue = toInput?.value || '';
  if (!fromValue || !toValue) {
    toast('Choose both From Date and To Date.', true);
    return null;
  }
  const fromDate = new Date(`${fromValue}T00:00:00`);
  const toDate = new Date(`${toValue}T23:59:59.999`);
  if (!Number.isFinite(fromDate.getTime()) || !Number.isFinite(toDate.getTime()) || fromDate > toDate) {
    toast('Choose a valid date range. From Date must be before To Date.', true);
    return null;
  }
  return {
    fromValue,
    toValue,
    fromMs: fromDate.getTime(),
    toMs: toDate.getTime(),
  };
}

async function loadPunchesForEmployeeRange(employee, fromMs, toMs, options = {}) {
  const primaryId = employee.id || employee.employeeId || '';
  const workerIds = new Set([primaryId, employee.employeeId].filter(Boolean));
  const directWorkerIds = new Set([primaryId, employee.employeeId].filter(Boolean));
  getCompatibleWorkerRecords(employee).forEach((record) => {
    if (record.id) workerIds.add(record.id);
    if (record.employeeId) workerIds.add(record.employeeId);
  });
  if (primaryId) {
    const mergedConstraints = state.me && isManager()
      ? [where('mergedInto', '==', primaryId)]
      : [where('status', '==', 'merged'), where('mergedInto', '==', primaryId)];
    const mergedSnapshot = await getDocs(query(collection(db, 'employees'), ...mergedConstraints));
    mergedSnapshot.docs.forEach((record) => {
      workerIds.add(record.id);
      directWorkerIds.add(record.id);
    });
  }

  const rows = [];
  let nameQuerySucceeded = false;
  if (employee.name) {
    try {
      const nameRows = await fetchPunchesWithRange(
        [where('nameKey', '==', normalizeName(employee.name))],
        fromMs,
        toMs
      );
      rows.push(...nameRows.filter((punch) => {
        const punchEmployeeId = String(punch.employeeId || '');
        return !punchEmployeeId || workerIds.has(punchEmployeeId);
      }));
      nameQuerySucceeded = true;
    } catch (error) {
      console.warn('Normalized-name history lookup unavailable:', error.message);
    }
  }

  const idsToQuery = nameQuerySucceeded ? [...directWorkerIds] : [...workerIds];
  const idRows = await Promise.all(idsToQuery.map((workerId) =>
    fetchPunchesWithRange([where('employeeId', '==', workerId)], fromMs, toMs)
  ));
  idRows.forEach((workerRows) => rows.push(...workerRows));

  if (options.allowLegacyNameFallback && employee.name) {
    const legacyConstraints = state.me && isManager()
      ? []
      : [where('employeeId', '==', '')];
    const legacyRows = await fetchPunchesWithRange(
      legacyConstraints,
      fromMs,
      toMs
    );
    const normalizedName = normalizeName(employee.name);
    rows.push(...legacyRows.filter((punch) =>
      normalizeName(punch.name || punch.employeeName || '') === normalizedName ||
      String(punch.nameKey || '').trim().toLowerCase() === normalizedName
    ));
  }

  return dedupePunches(rows);
}

function getCompatibleWorkerRecords(employee) {
  if (!isActiveEmployee(employee)) return [employee];
  const sourceMap = new Map();
  const sourceRecords = state.me && isManager()
    ? [...state.publicEmployeeRecords, ...state.allEmployees]
    : state.publicEmployeeRecords;
  sourceRecords.forEach((record) => sourceMap.set(record.id, record));
  const source = [...sourceMap.values()];
  const normalizedName = normalizeIdentityPart(employee.name || employee.nameKey);
  const sameName = source.filter((record) =>
    isActiveEmployee(record) &&
    normalizeIdentityPart(record.name || record.nameKey) === normalizedName
  );
  const nonBlankScopes = new Set(
    sameName.map((record) => {
      const agency = normalizeIdentityPart(record.agencyId);
      const site = normalizeIdentityPart(record.assignedSiteId || record.siteId);
      return agency || site ? `${agency}|${site}` : '';
    }).filter(Boolean)
  );
  if (nonBlankScopes.size <= 1) return sameName;
  const targetAgency = normalizeIdentityPart(employee.agencyId);
  const targetSite = normalizeIdentityPart(employee.assignedSiteId || employee.siteId);
  return sameName.filter((record) =>
    normalizeIdentityPart(record.agencyId) === targetAgency &&
    normalizeIdentityPart(record.assignedSiteId || record.siteId) === targetSite
  );
}

async function fetchPunchesWithRange(baseConstraints, fromMs, toMs) {
  try {
    const rangedQuery = query(
      collection(db, 'punches'),
      ...baseConstraints,
      where('timestampMs', '>=', fromMs),
      where('timestampMs', '<=', toMs)
    );
    const snapshot = await getDocs(rangedQuery);
    return snapshot.docs.map((record) => ({ id: record.id, ...record.data() }));
  } catch (error) {
    if (!['failed-precondition', 'permission-denied'].includes(error.code)) throw error;
    const snapshot = await getDocs(query(collection(db, 'punches'), ...baseConstraints));
    return snapshot.docs
      .map((record) => ({ id: record.id, ...record.data() }))
      .filter((punch) => Number(punch.timestampMs || 0) >= fromMs && Number(punch.timestampMs || 0) <= toMs);
  }
}

function dedupePunches(punches) {
  const unique = new Map();
  punches.forEach((punch) => {
    const key = punch.id || [
      punch.employeeId || '',
      punch.timestampMs || 0,
      punch.action || '',
      normalizeName(punch.name || punch.employeeName || ''),
    ].join('|');
    unique.set(key, punch);
  });
  return [...unique.values()].sort((left, right) => Number(left.timestampMs || 0) - Number(right.timestampMs || 0));
}

function buildTimeRangeTotals(punches) {
  const grouped = new Map();
  dedupePunches(punches).forEach((punch) => {
    const timestampMs = Number(punch.timestampMs || 0);
    if (!timestampMs) return;
    const dateKey = punch.dateKey || formatDateKey(new Date(timestampMs));
    if (!grouped.has(dateKey)) grouped.set(dateKey, []);
    grouped.get(dateKey).push({ ...punch, timestampMs });
  });

  const daily = {};
  const weeklyMinutes = new Map();
  let totalMinutes = 0;

  [...grouped.entries()].sort(([left], [right]) => left.localeCompare(right)).forEach(([dateKey, dayPunches]) => {
    const sorted = [...dayPunches].sort((left, right) => left.timestampMs - right.timestampMs);
    const actionTimes = {
      clock_in: [],
      start_lunch: [],
      end_lunch: [],
      clock_out: [],
    };
    let activeStart = null;
    let minutes = 0;

    sorted.forEach((punch) => {
      if (actionTimes[punch.action]) actionTimes[punch.action].push(punch.timestampMs);
      if (punch.action === 'clock_in') activeStart = punch.timestampMs;
      if (punch.action === 'start_lunch') {
        if (activeStart) minutes += Math.max(0, Math.round((punch.timestampMs - activeStart) / 60000));
        activeStart = null;
      }
      if (punch.action === 'end_lunch') activeStart = punch.timestampMs;
      if (punch.action === 'clock_out') {
        if (activeStart) minutes += Math.max(0, Math.round((punch.timestampMs - activeStart) / 60000));
        activeStart = null;
      }
    });

    const warnings = [];
    if (actionTimes.clock_in.length && !actionTimes.clock_out.length) warnings.push('Missing Clock Out');
    if (actionTimes.clock_out.length && !actionTimes.clock_in.length) warnings.push('Missing Clock In');
    if (actionTimes.start_lunch.length && !actionTimes.end_lunch.length) warnings.push('Missing Lunch In');
    if (actionTimes.end_lunch.length && !actionTimes.start_lunch.length) warnings.push('Missing Lunch Out');

    const weekKey = formatDateKey(getMondayDate(new Date(`${dateKey}T12:00:00`)));
    weeklyMinutes.set(weekKey, (weeklyMinutes.get(weekKey) || 0) + minutes);
    totalMinutes += minutes;
    daily[dateKey] = {
      actionTimes,
      minutes,
      hours: Number((minutes / 60).toFixed(2)),
      warnings,
    };
  });

  let regularMinutes = 0;
  let overtimeMinutes = 0;
  weeklyMinutes.forEach((minutes) => {
    regularMinutes += Math.min(minutes, 40 * 60);
    overtimeMinutes += Math.max(0, minutes - (40 * 60));
  });

  return {
    daily,
    punches: dedupePunches(punches),
    daysWorked: Object.keys(daily).length,
    totalHours: Number((totalMinutes / 60).toFixed(2)),
    regularHours: Number((regularMinutes / 60).toFixed(2)),
    overtimeHours: Number((overtimeMinutes / 60).toFixed(2)),
  };
}

function renderTimeRangeSummary({
  totals,
  totalElement,
  regularElement,
  overtimeElement,
  daysElement,
  statusElement,
  resultsElement,
  emptyMessage,
  statusPrefix = '',
}) {
  if (totalElement) totalElement.textContent = Number(totals.totalHours || 0).toFixed(2);
  if (regularElement) regularElement.textContent = Number(totals.regularHours || 0).toFixed(2);
  if (overtimeElement) overtimeElement.textContent = Number(totals.overtimeHours || 0).toFixed(2);
  if (daysElement) daysElement.textContent = String(totals.daysWorked || 0);

  const dateKeys = Object.keys(totals.daily || {}).sort().reverse();
  setTimeRangeStatus(
    statusElement,
    `${statusPrefix}Total Hours: ${Number(totals.totalHours || 0).toFixed(2)}`
  );
  if (!resultsElement) return;
  resultsElement.innerHTML = dateKeys.length
    ? dateKeys.map((dateKey) => renderTimeResultCard(dateKey, totals.daily[dateKey])).join('')
    : `<div class="empty-state">${escapeHtml(emptyMessage)}</div>`;
}

function renderTimeResultCard(dateKey, day) {
  const actionLabel = (action) => {
    const times = day.actionTimes[action] || [];
    return times.length ? times.map(formatTime).join(', ') : '-';
  };
  return `
    <article class="time-result-card">
      <div class="time-result-card-head">
        <strong>${escapeHtml(formatDateOnly(new Date(`${dateKey}T12:00:00`).getTime()))}</strong>
        <span class="pill">${Number(day.hours || 0).toFixed(2)} hours</span>
      </div>
      <div class="time-punch-grid">
        <div class="time-punch-item"><span>Clock In</span><strong>${escapeHtml(actionLabel('clock_in'))}</strong></div>
        <div class="time-punch-item"><span>Lunch Out</span><strong>${escapeHtml(actionLabel('start_lunch'))}</strong></div>
        <div class="time-punch-item"><span>Lunch In</span><strong>${escapeHtml(actionLabel('end_lunch'))}</strong></div>
        <div class="time-punch-item"><span>Clock Out</span><strong>${escapeHtml(actionLabel('clock_out'))}</strong></div>
      </div>
      ${day.warnings.length
        ? `<div class="time-warning">${day.warnings.map(escapeHtml).join(' · ')}</div>`
        : ''}
    </article>
  `;
}

function setTimeRangeStatus(element, message) {
  if (element) element.textContent = message;
}

async function handlePublicTimeFixSubmit(event) {
  event.preventDefault();
  const employee = getSelectedWorkerByName();
  if (!employee) return;

  const requestedAction = els.workerFixActionInput?.value || '';
  const requestedDate = els.workerFixDateInput?.value || '';
  const requestedTime = els.workerFixTimeInput?.value || '';
  const reason = String(els.workerFixReasonInput?.value || '').trim();
  const requestedTimestampMs = parseLocalDateAndTime(requestedDate, requestedTime);

  if (!requestedAction || !requestedDate || !requestedTime || !reason || !requestedTimestampMs) {
    toast('Fill out every time fix field.', true);
    return;
  }

  try {
    await addDoc(collection(db, 'missedPunchRequests'), {
      uid: '',
      employeeId: employee.employeeId || employee.id,
      employeeNumber: employee.employeeNumber || '',
      companyId: employee.companyId || new URLSearchParams(window.location.search).get('company') || '',
      agencyId: employee.agencyId || '',
      name: employee.name,
      nameKey: normalizeName(employee.name),
      requestedAction,
      requestedDate,
      requestedTime,
      requestedTimestampMs,
      reason,
      status: 'pending',
      source: 'public_worker',
      reviewedBy: '',
      reviewedAt: null,
      approvedBy: '',
      approvedAt: null,
      createdAt: serverTimestamp(),
      updatedAt: serverTimestamp(),
    });

    els.workerFixForm?.reset();
    if (els.workerFixDateInput) els.workerFixDateInput.value = formatDateInput(new Date());
    if (els.workerFixTimeInput) els.workerFixTimeInput.value = formatTimeForInput(Date.now());
    toggleWorkerSelfService(null);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = `Time fix request submitted for ${employee.name}. A manager will review it.`;
    }
    toast('Time fix request submitted.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not submit the time fix request.', true);
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
    toast('Your role does not allow manual punch corrections.', true);
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
      companyId: state.companyId || '',
      agencyId: state.agencyId || '',
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
      companyId: state.companyId || '',
    });

    els.manualPunchForm?.reset();
    if (els.manualPunchDateInput) els.manualPunchDateInput.value = formatDateInput(new Date());
    if (els.manualPunchTimeInput) els.manualPunchTimeInput.value = formatTimeForInput(Date.now());

    toast('Manual punch added.');
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
  state.companyId = null;
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
  const canEdit = canEditPunches();
  const canManage = canManageEmployees();

  // Employee-only tabs
  els.myTimecardTabBtn?.classList.toggle('hidden', !emp);
  els.missedPunchTabBtn?.classList.toggle('hidden', !emp);

  // Punch editors and employee managers
  els.managerTabBtn?.classList.toggle('hidden', !canEdit);
  els.timesheetsTabBtn?.classList.toggle('hidden', !canEdit);
  els.editPunchesTabBtn?.classList.toggle('hidden', !canEdit);
  els.approvalsTabBtn?.classList.toggle('hidden', !canEdit);
  els.employeesTabBtn?.classList.toggle('hidden', !canManage);
  els.adminTabBtn?.classList.toggle('hidden', !isAdmin());
  els.agencyTabBtn?.classList.toggle('hidden', !canEdit);

  if (emp) {
    if (els.myTimecardWeekPicker) {
      els.myTimecardWeekPicker.value = formatDateInput(state.selectedWeekStart);
    }
    if (els.mpDateInput) els.mpDateInput.value = formatDateInput(new Date());
    switchTab('myTimecardTab');
    attachMyTimecardView();
    attachMyMissedPunchView();
  } else if (canEdit) {
    switchTab('managerTab');
    if (canManage) {
      attachEmployeesView();
    }
    attachApprovalView();
  } else {
    document.querySelectorAll('.tab-panel').forEach((panel) => panel.classList.add('hidden'));
    toast(`Role "${normalizedRole() || 'unknown'}" has no dashboard permissions.`, true);
  }
}

function switchTab(tabId) {
  if (!canAccessTab(tabId)) {
    toast('Your role does not have access to that section.', true);
    return;
  }
  document.querySelectorAll('.tab').forEach((tab) => {
    tab.classList.toggle('active', tab.dataset.tab === tabId);
  });

  document.querySelectorAll('.tab-panel').forEach((panel) => {
    panel.classList.toggle('hidden', panel.id !== tabId);
  });
}

function canAccessTab(tabId) {
  if (['myTimecardTab', 'missedPunchTab'].includes(tabId)) return isEmployee();
  if (['managerTab', 'timesheetsTab', 'editPunchesTab', 'approvalsTab', 'agencyTab'].includes(tabId)) {
    return canEditPunches();
  }
  if (tabId === 'employeesTab') return canManageEmployees();
  if (tabId === 'adminTab') return isAdmin();
  return false;
}

function attachManagerLiveViews() {
  const constraints = [];
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
        const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
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
  if (!canEditPunches()) {
    els.editPunchesBody.innerHTML = '<tr><td colspan="8">Your role cannot edit punches.</td></tr>';
    return;
  }

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
          <button class="secondary-btn manager-edit-punch-btn" data-id="${row.id}" type="button">Edit</button>
          ${canDeletePunches()
            ? `<button class="danger-btn manager-delete-punch-btn" data-id="${row.id}" type="button">Delete</button>`
            : ''}
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
    toast('Your role does not allow punch editing.', true);
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
      editedAt: serverTimestamp()
    });

    await updateDoc(doc(db, 'punches', punchId), updatedPayload);

    toast('Punch updated.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not update punch.', true);
  }
}

async function deletePunchRecord(punchId) {
  if (!canDeletePunches()) {
    toast('Your role does not allow punch deletion.', true);
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
      editedAt: serverTimestamp()
    });

    await deleteDoc(doc(db, 'punches', punchId));
    toast('Punch deleted.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not delete punch.', true);
  }
}

function attachTimesheetView() {
  const weekKey = formatDateKey(state.selectedWeekStart);

  const punchConstraints = [where('weekKey', '==', weekKey)];
  if (isAgencyUser()) punchConstraints.push(where('agencyId', '==', state.agencyId));
  punchConstraints.push(orderBy('timestampMs', 'asc'));

  const punchesQuery = query(collection(db, 'punches'), ...punchConstraints);

  const tsConstraints = [where('weekKey', '==', weekKey)];
  if (isAgencyUser()) tsConstraints.push(where('agencyId', '==', state.agencyId));

  const timesheetsQuery = query(collection(db, 'timesheets'), ...tsConstraints);

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

function isEmployee() {
  return ['employee', 'worker'].includes(normalizedRole());
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
  constraints.push(orderBy('timestampMs', 'asc'));

  const q = query(collection(db, 'punches'), ...constraints);

  state._myTcUnsub = onSnapshot(q, (snap) => {
    const rows = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
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
      companyId: state.companyId || '',
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

  const constraints = [where('uid', '==', state.me.uid)];

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
  const constraints = [];
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
  if (!canEditPunches()) { toast('Your role cannot approve punch requests.', true); return; }

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
      companyId: req.companyId || '',
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
  if (!canEditPunches()) { toast('Your role cannot deny punch requests.', true); return; }

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
  const empConstraints = [];
  if (isAgencyUser()) empConstraints.push(where('agencyId', '==', state.agencyId));
  empConstraints.push(orderBy('name', 'asc'));

  const empQuery = query(collection(db, 'employees'), ...empConstraints);

  state.unsubscribers.push(
    onSnapshot(empQuery, (snap) => {
      state.allEmployees = snap.docs.map((d) => ({ id: d.id, ...d.data() }));
      renderEmployeeList(state.allEmployees);
      renderManagerTimeWorkerOptions();
      if (isAdmin()) renderDuplicateWorkers(false);
    }, (error) => {
      console.error(error);
      toast(error.message || 'Could not load employees.', true);
    })
  );
}

function renderEmployeeList(employees) {
  if (!els.employeeListBody || !els.inactiveWorkerListBody) return;

  const filter = String(els.empFilterInput?.value || '').trim().toLowerCase();
  const statusFilter = state.employeeStatusFilter || 'active';
  const filtered = employees.filter((e) => {
    const employeeStatus = String(e.status || 'active').toLowerCase();
    if (statusFilter !== 'all' && employeeStatus !== statusFilter) return false;
    if (!filter) return true;
    return (
      String(e.name || '').toLowerCase().includes(filter) ||
      String(e.employeeNumber || '').toLowerCase().includes(filter)
    );
  });
  const rosterEmployees = statusFilter === 'active'
    ? collapseDuplicateEmployees(filtered)
    : filtered;
  const inactiveEmployees = employees.filter((employee) => !isActiveEmployee(employee));

  els.inactiveWorkerListBody.innerHTML = inactiveEmployees.length
    ? inactiveEmployees.map((employee) => `
      <tr>
        <td>${escapeHtml(employee.employeeNumber || '-')}</td>
        <td>${escapeHtml(employee.name || '-')}</td>
        <td>${escapeHtml(employee.status || 'inactive')}</td>
        <td>${escapeHtml(formatRemovedAt(employee))}</td>
        <td>${escapeHtml(employee.removedBy || '-')}</td>
        <td>${escapeHtml(employee.removalReason || '-')}</td>
        <td>${String(employee.status).toLowerCase() === 'merged'
          ? `Merged into ${escapeHtml(employee.mergedInto || '-')}`
          : `<button class="primary-btn emp-reactivate-btn" data-id="${employee.id}" type="button">Reactivate</button>`}</td>
      </tr>
    `).join('')
    : '<tr><td colspan="7">No removed or inactive workers.</td></tr>';

  els.inactiveWorkerListBody.querySelectorAll('.emp-reactivate-btn').forEach((button) => {
    button.addEventListener('click', () => reactivateEmployee(button.dataset.id));
  });

  if (!rosterEmployees.length) {
    els.employeeListBody.innerHTML = '<tr><td colspan="6">No employees found.</td></tr>';
    return;
  }

  els.employeeListBody.innerHTML = rosterEmployees.map((emp) => `
    <tr>
      <td>${escapeHtml(emp.employeeNumber || '-')}</td>
      <td>${escapeHtml(emp.name || '-')}</td>
      <td>${escapeHtml(agencyLabel(emp.agencyId))}</td>
      <td>${escapeHtml(emp.assignedSiteId || '-')}</td>
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

function renderManagerTimeWorkerOptions() {
  if (!els.managerTimeWorkerSelect || !isManager()) return;
  const selectedId = els.managerTimeWorkerSelect.value;
  const activeEmployees = collapseDuplicateEmployees(state.allEmployees.filter(isActiveEmployee));
  const historicalEmployees = state.allEmployees.filter((employee) =>
    !isActiveEmployee(employee) && String(employee.status || '').toLowerCase() !== 'merged'
  );
  const employees = [...activeEmployees, ...historicalEmployees]
    .sort((left, right) => String(left.name || '').localeCompare(String(right.name || '')));

  els.managerTimeWorkerSelect.innerHTML = [
    '<option value="">Select a worker</option>',
    ...employees.map((employee) => `
      <option value="${escapeHtml(employee.id)}">
        ${escapeHtml(employee.name || '-')} · ${escapeHtml(employee.employeeNumber || '-')} · ${escapeHtml(agencyLabel(employee.agencyId))}
      </option>
    `),
  ].join('');
  if (employees.some((employee) => employee.id === selectedId)) {
    els.managerTimeWorkerSelect.value = selectedId;
  }
}

async function lookupManagerTimeRange() {
  if (!isManager()) return;
  const employeeId = els.managerTimeWorkerSelect?.value || '';
  const employee = state.allEmployees.find((row) => row.id === employeeId);
  if (!employee) {
    toast('Choose a worker first.', true);
    return;
  }
  if (isAgencyUser() && employee.agencyId !== state.agencyId) {
    toast('That worker is outside your agency.', true);
    return;
  }
  const range = readDateRange(els.managerTimeFromInput, els.managerTimeToInput);
  if (!range) return;

  setTimeRangeStatus(els.managerTimeRangeStatus, `Loading time for ${employee.name}...`);
  if (els.managerTimeExportBtn) els.managerTimeExportBtn.disabled = true;
  try {
    const punches = await loadPunchesForEmployeeRange(employee, range.fromMs, range.toMs, {
      allowLegacyNameFallback: true,
    });
    const totals = buildTimeRangeTotals(punches);
    state.managerTimeLookup = { employee, range, totals };
    renderTimeRangeSummary({
      totals,
      totalElement: els.managerTimeTotalValue,
      regularElement: els.managerTimeRegularValue,
      overtimeElement: els.managerTimeOvertimeValue,
      daysElement: els.managerTimeDaysValue,
      statusElement: els.managerTimeRangeStatus,
      resultsElement: els.managerTimeRangeResults,
      emptyMessage: 'No punches were found for this worker and date range.',
      statusPrefix: `${employee.name} · `,
    });
    if (els.managerTimeExportBtn) els.managerTimeExportBtn.disabled = !totals.daysWorked;
  } catch (error) {
    console.error(error);
    state.managerTimeLookup = null;
    setTimeRangeStatus(els.managerTimeRangeStatus, error.message || 'Could not load worker time.');
    if (els.managerTimeRangeResults) {
      els.managerTimeRangeResults.innerHTML = '<div class="empty-state">The lookup failed without changing any punch data.</div>';
    }
    toast(error.message || 'Could not load worker time.', true);
  }
}

function exportManagerTimeRangeCsv() {
  if (!isManager() || !state.managerTimeLookup) return;
  const { employee, range, totals } = state.managerTimeLookup;
  if (isAgencyUser() && employee.agencyId !== state.agencyId) return;

  const headers = [
    'Worker Name',
    'Employee Number',
    'Date',
    'Clock In',
    'Lunch Out',
    'Lunch In',
    'Clock Out',
    'Total Hours',
    'Warnings',
    'Agency',
    'Site',
  ];
  const rows = Object.keys(totals.daily).sort().map((dateKey) => {
    const day = totals.daily[dateKey];
    const times = (action) => (day.actionTimes[action] || []).map(formatTime).join(' | ');
    return [
      employee.name || '',
      employee.employeeNumber || '',
      dateKey,
      times('clock_in'),
      times('start_lunch'),
      times('end_lunch'),
      times('clock_out'),
      Number(day.hours || 0).toFixed(2),
      day.warnings.join(' | '),
      agencyLabel(employee.agencyId),
      employee.assignedSiteId || employee.siteId || '',
    ];
  });
  rows.push([
    employee.name || '',
    employee.employeeNumber || '',
    `${range.fromValue} to ${range.toValue}`,
    '',
    '',
    '',
    '',
    Number(totals.totalHours || 0).toFixed(2),
    `Regular: ${Number(totals.regularHours || 0).toFixed(2)} | Overtime: ${Number(totals.overtimeHours || 0).toFixed(2)}`,
    agencyLabel(employee.agencyId),
    employee.assignedSiteId || employee.siteId || '',
  ]);

  downloadCsv(
    `worker-time-${normalizeName(employee.name)}-${range.fromValue}-to-${range.toValue}.csv`,
    [headers, ...rows]
  );
}

function downloadCsv(filename, rows) {
  const csv = rows
    .map((row) => row.map((value) => `"${String(value ?? '').replaceAll('"', '""')}"`).join(','))
    .join('\r\n');
  const url = URL.createObjectURL(new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8' }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function isActiveEmployee(employee) {
  return !['inactive', 'terminated', 'removed', 'merged'].includes(String(employee?.status || 'active').toLowerCase());
}

function formatRemovedAt(employee) {
  const milliseconds = Number(employee?.removedAtMs || 0)
    || Number(employee?.removedAt?.seconds || 0) * 1000;
  return milliseconds ? formatDateTime(milliseconds) : '-';
}

function loadEmployeeForEdit(empId) {
  const emp = (state.allEmployees || []).find((e) => e.id === empId);
  if (!emp) { toast('Employee not found.', true); return; }

  if (els.employeeDocId) els.employeeDocId.value = empId;
  if (els.empNameInput) els.empNameInput.value = emp.name || '';
  if (els.empNumberInput) els.empNumberInput.value = emp.employeeNumber || '';
  if (els.empAgencySelect) els.empAgencySelect.value = emp.agencyId || '';
  if (els.empSiteInput) els.empSiteInput.value = emp.assignedSiteId || '';
  if (els.empStatusSelect) els.empStatusSelect.value = emp.status || 'active';
  if (els.empPinInput) els.empPinInput.value = '';
  els.empCancelEditBtn?.classList.remove('hidden');
}

function cancelEmployeeEdit() {
  els.employeeForm?.reset();
  if (els.employeeDocId) els.employeeDocId.value = '';
  els.empCancelEditBtn?.classList.add('hidden');
}

async function handleSaveEmployee(event) {
  event.preventDefault();

  if (!canManageEmployees()) {
    toast('Your role does not allow employee management.', true);
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
  const assignedSiteId = els.empSiteInput?.value.trim() || '';
  const status = els.empStatusSelect?.value || 'active';
  const existingId = els.employeeDocId?.value || '';
  const existingEmployee = state.allEmployees.find((employee) => employee.id === existingId);

  // Auto-generate employee number if blank
  if (!employeeNumber) {
    employeeNumber = await generateNextEmployeeNumber();
  }

  if (!existingId) {
    const reusable = findReusableEmployee({
      name,
      employeeNumber,
      agencyId,
      assignedSiteId,
    }, state.allEmployees);
    if (reusable) {
      await logAudit('duplicate_detected', 'employee', reusable.id, {}, {
        attemptedName: name,
        attemptedEmployeeNumber: employeeNumber,
        agencyId,
        assignedSiteId,
      });
      loadEmployeeForEdit(reusable.id);
      toast(`Existing worker reused: ${reusable.name}.`);
      return;
    }
  }

  const payload = {
    name,
    nameKey,
    employeeNumber,
    companyId: state.companyId || '',
    agencyId,
    assignedSiteId,
    status,
    updatedAt: serverTimestamp(),
  };

  if (existingId && status !== 'active' && isActiveEmployee(existingEmployee)) {
    const removalReason = String(prompt(`Reason for marking ${name} ${status}:`) || '').trim();
    if (!removalReason) {
      toast('A removal reason is required.', true);
      return;
    }
    payload.removedAt = serverTimestamp();
    payload.removedAtMs = Date.now();
    payload.removedBy = state.profile?.name || state.me?.email || 'Manager';
    payload.removedById = state.me?.uid || '';
    payload.removalReason = removalReason;
  }

  const pin = String(els.empPinInput?.value || '').trim();
  if (pin && !/^\d{4,12}$/.test(pin)) {
    toast('Optional PIN must be 4 to 12 digits.', true);
    return;
  }

  try {
    if (existingId) {
      // Update existing employee
      if (pin) {
        payload.pinHash = await hashWorkerPin(pin, existingId);
        payload.pinUpdatedAt = serverTimestamp();
      }
      await updateDoc(doc(db, 'employees', existingId), payload);
      const auditAction = existingEmployee?.status !== status
        ? 'worker_status_changed'
        : 'employee_updated';
      await logAudit(auditAction, 'employee', existingId, existingEmployee || {}, payload);
      toast('Employee updated.');
    } else {
      // Create new employee
      payload.createdAt = serverTimestamp();
      const newRef = await addDoc(collection(db, 'employees'), payload);
      // Write employeeId field = doc ID
      const identityPayload = { employeeId: newRef.id };
      if (pin) {
        identityPayload.pinHash = await hashWorkerPin(pin, newRef.id);
        identityPayload.pinUpdatedAt = serverTimestamp();
      }
      await updateDoc(newRef, identityPayload);
      await logAudit('employee_created', 'employee', newRef.id, {}, payload);
      toast('Employee created: ' + employeeNumber);
    }

    cancelEmployeeEdit();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save employee.', true);
  }
}

async function reactivateEmployee(employeeId) {
  if (!isManager()) return;
  const employee = state.allEmployees.find((row) => row.id === employeeId);
  if (!employee) return;
  const payload = {
    status: 'active',
    reactivatedAt: serverTimestamp(),
    reactivatedAtMs: Date.now(),
    reactivatedBy: state.profile?.name || state.me?.email || 'Manager',
    reactivatedById: state.me?.uid || '',
    updatedAt: serverTimestamp(),
  };
  try {
    await updateDoc(doc(db, 'employees', employeeId), payload);
    await logAudit('worker_reactivated', 'employee', employeeId, employee, payload);
    toast('Employee reactivated.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not reactivate employee.', true);
  }
}

function buildDuplicateGroups(employees = state.allEmployees) {
  const candidates = employees.filter(isActiveEmployee);
  const buckets = new Map();
  const employeesByName = new Map();

  const addToBucket = (key, reason, employee) => {
    if (!key) return;
    if (!buckets.has(key)) buckets.set(key, { reason, employees: [] });
    buckets.get(key).employees.push(employee);
  };

  candidates.forEach((employee) => {
    const number = normalizeIdentityPart(employee.employeeNumber);
    const name = normalizeIdentityPart(employee.name || employee.nameKey);
    const agency = normalizeIdentityPart(employee.agencyId);
    const site = normalizeIdentityPart(employee.assignedSiteId || employee.siteId);
    if (!employeesByName.has(name)) employeesByName.set(name, []);
    employeesByName.get(name).push(employee);
    if (number) {
      addToBucket(
        `number:${number}|${name}|${agency}|${site}`,
        `Same employee number, normalized name, agency, and site (${employee.employeeNumber})`,
        employee
      );
    }
  });

  employeesByName.forEach((nameEmployees, name) => {
    const nonBlankScopes = new Set(
      nameEmployees
        .map((employee) => {
          const agency = normalizeIdentityPart(employee.agencyId);
          const site = normalizeIdentityPart(employee.assignedSiteId || employee.siteId);
          return { agency, site, key: `${agency}|${site}` };
        })
        .filter((scope) => scope.agency || scope.site)
        .map((scope) => scope.key)
    );

    if (nonBlankScopes.size <= 1) {
      nameEmployees.forEach((employee) => {
        addToBucket(`name:${name}`, 'Same normalized name with one compatible agency/site', employee);
      });
      return;
    }

    nameEmployees.forEach((employee) => {
      const agency = normalizeIdentityPart(employee.agencyId);
      const site = normalizeIdentityPart(employee.assignedSiteId || employee.siteId);
      addToBucket(`scope:${name}|${agency}|${site}`, 'Same normalized name, agency, and site', employee);
    });
  });

  const uniqueGroups = new Map();
  [...buckets.values()]
    .filter((group) => group.employees.length > 1)
    .forEach((group) => {
      const ids = [...new Set(group.employees.map((employee) => employee.id))].sort();
      if (ids.length < 2) return;
      const signature = ids.join('|');
      if (!uniqueGroups.has(signature)) {
        uniqueGroups.set(signature, {
          reasons: [],
          employees: ids
            .map((id) => candidates.find((row) => row.id === id))
            .sort(compareEmployeeRecords),
        });
      }
      uniqueGroups.get(signature).reasons.push(group.reason);
    });

  return [...uniqueGroups.values()]
    .map((group) => ({ ...group, reasons: [...new Set(group.reasons)] }))
    .sort((left, right) => right.employees.length - left.employees.length);
}

function renderDuplicateWorkers(logDetection = false) {
  if (!els.duplicateWorkersList || !isAdmin()) return;
  state.duplicateGroups = buildDuplicateGroups();

  if (!state.duplicateGroups.length) {
    els.duplicateWorkersList.classList.add('empty-state');
    els.duplicateWorkersList.innerHTML = 'No active duplicate groups detected.';
    return;
  }

  els.duplicateWorkersList.classList.remove('empty-state');
  els.duplicateWorkersList.innerHTML = state.duplicateGroups.map((group, groupIndex) => `
    <article class="duplicate-group" data-group-index="${groupIndex}">
      <div>
        <strong>${escapeHtml(group.reasons.join(' / '))}</strong>
        <span class="tiny">${group.employees.length} active records</span>
      </div>
      <div class="duplicate-options">
        ${group.employees.map((employee, employeeIndex) => `
          <label class="duplicate-option">
            <input type="radio" name="duplicate-primary-${groupIndex}" value="${employee.id}" ${employeeIndex === 0 ? 'checked' : ''} />
            <span>Primary: ${escapeHtml(employee.name || '-')} · ${escapeHtml(employee.employeeNumber || '-')} · ${escapeHtml(agencyLabel(employee.agencyId))} · ${escapeHtml(employee.assignedSiteId || '-')}</span>
            <input class="duplicate-merge-check" type="checkbox" value="${employee.id}" ${employeeIndex === 0 ? '' : 'checked'} />
            <span>Merge this record</span>
          </label>
        `).join('')}
      </div>
      <button class="danger-btn merge-workers-btn" data-group-index="${groupIndex}" type="button">Merge Selected Records</button>
    </article>
  `).join('');

  els.duplicateWorkersList.querySelectorAll('input[type="radio"]').forEach((radio) => {
    radio.addEventListener('change', () => syncDuplicatePrimary(Number(radio.closest('.duplicate-group').dataset.groupIndex)));
  });
  state.duplicateGroups.forEach((_, index) => syncDuplicatePrimary(index));
  els.duplicateWorkersList.querySelectorAll('.merge-workers-btn').forEach((button) => {
    button.addEventListener('click', () => mergeDuplicateGroup(Number(button.dataset.groupIndex)));
  });

  if (logDetection) {
    state.duplicateGroups.forEach((group) => {
      logAudit('duplicate_detected', 'employee_group', group.employees.map((employee) => employee.id).join(','), {}, {
        reasons: group.reasons,
        employeeIds: group.employees.map((employee) => employee.id),
      });
    });
  }
}

function syncDuplicatePrimary(groupIndex) {
  const container = els.duplicateWorkersList?.querySelector(`[data-group-index="${groupIndex}"]`);
  if (!container) return;
  const primaryId = container.querySelector('input[type="radio"]:checked')?.value || '';
  container.querySelectorAll('.duplicate-merge-check').forEach((checkbox) => {
    checkbox.disabled = checkbox.value === primaryId;
    if (checkbox.disabled) checkbox.checked = false;
  });
}

async function mergeDuplicateGroup(groupIndex) {
  if (!isAdmin()) {
    toast('Only admins can merge duplicate workers.', true);
    return;
  }
  const container = els.duplicateWorkersList?.querySelector(`[data-group-index="${groupIndex}"]`);
  const primaryId = container?.querySelector('input[type="radio"]:checked')?.value || '';
  const duplicateIds = [...(container?.querySelectorAll('.duplicate-merge-check:checked') || [])]
    .map((checkbox) => checkbox.value)
    .filter((id) => id && id !== primaryId);
  const primary = state.allEmployees.find((employee) => employee.id === primaryId);

  if (!primary || !duplicateIds.length) {
    toast('Choose a primary worker and at least one duplicate.', true);
    return;
  }
  if (!confirm(`Merge ${duplicateIds.length} duplicate record(s) into ${primary.name}? No records will be deleted.`)) return;

  try {
    const totals = { punches: 0, missedPunchRequests: 0, timesheets: 0 };
    for (const collectionName of Object.keys(totals)) {
      totals[collectionName] = await reassignWorkerReferences(collectionName, duplicateIds, primary);
      if (totals[collectionName]) {
        await logAudit(
          collectionName === 'punches' ? 'punch_reassigned' : `${collectionName}_reassigned`,
          collectionName,
          primary.id,
          { duplicateWorkerIds: duplicateIds },
          { primaryWorkerId: primary.id, count: totals[collectionName] }
        );
      }
    }

    const mergedAtMs = Date.now();
    const actor = state.profile?.name || state.me?.email || 'Admin';
    await commitDocumentUpdates(duplicateIds.map((duplicateId) => ({
      ref: doc(db, 'employees', duplicateId),
      data: {
        status: 'merged',
        mergedInto: primary.id,
        mergedAt: serverTimestamp(),
        mergedAtMs,
        mergedBy: actor,
        mergedById: state.me?.uid || '',
        updatedAt: serverTimestamp(),
      },
    })));

    await logAudit('worker_merged', 'employee', primary.id, {
      duplicateWorkerIds: duplicateIds,
    }, {
      primaryWorkerId: primary.id,
      totals,
      mergedBy: actor,
    });
    toast(`Merged ${duplicateIds.length} worker record(s). No data was deleted.`);
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not complete the safe merge.', true);
  }
}

async function reassignWorkerReferences(collectionName, duplicateIds, primary) {
  const documents = new Map();
  for (const duplicateId of duplicateIds) {
    for (const fieldName of ['employeeId', 'workerId']) {
      const snapshot = await getDocs(query(collection(db, collectionName), where(fieldName, '==', duplicateId)));
      snapshot.docs.forEach((record) => documents.set(record.id, { id: record.id, ...record.data() }));
    }
  }

  const updates = [...documents.values()].map((record) => {
    const data = {
      employeeId: primary.employeeId || primary.id,
      employeeNumber: primary.employeeNumber || '',
      name: primary.name || record.name || '',
      nameKey: normalizeName(primary.name || record.name || ''),
      updatedAt: serverTimestamp(),
      mergedFromWorkerId: record.employeeId || record.workerId || '',
    };
    if (record.workerId !== undefined) data.workerId = primary.id;
    if (record.workerName !== undefined) data.workerName = primary.name || '';
    return { ref: doc(db, collectionName, record.id), data };
  });
  await commitDocumentUpdates(updates);
  return updates.length;
}

async function commitDocumentUpdates(updates) {
  for (let index = 0; index < updates.length; index += 400) {
    const batch = writeBatch(db);
    updates.slice(index, index + 400).forEach((update) => batch.update(update.ref, update.data));
    await batch.commit();
  }
}

async function logAudit(action, entityType, entityId, oldValue, newValue) {
  if (!state.me || !isManager()) return;
  try {
    await addDoc(collection(db, 'auditLogs'), {
      action,
      entityType,
      entityId,
      oldValue: serializeForExport(oldValue),
      newValue: serializeForExport(newValue),
      actorId: state.me.uid,
      actorName: state.profile?.name || state.me.email || 'Manager',
      companyId: state.companyId || '',
      createdAt: serverTimestamp(),
      createdAtMs: Date.now(),
    });
  } catch (error) {
    console.warn('Audit log could not be saved:', error.message);
  }
}

async function exportBackup() {
  if (!isManager()) return;
  try {
    const collections = ['employees', 'punches', 'missedPunchRequests', 'agencies', 'users'];
    const backup = {
      exportedAt: new Date().toISOString(),
      companyId: state.companyId || '',
      note: 'Non-destructive QRTimeclock backup export',
    };
    for (const collectionName of collections) {
      const constraints = [];
      const snapshot = await getDocs(query(collection(db, collectionName), ...constraints));
      backup[collectionName] = snapshot.docs.map((record) => ({
        id: record.id,
        ...serializeForExport(record.data()),
      }));
    }
    downloadJson(`qrtimeclock-backup-${formatDateKey(new Date())}.json`, backup);
    toast('Backup exported.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not export backup.', true);
  }
}

function serializeForExport(value) {
  if (value?.toDate instanceof Function) return value.toDate().toISOString();
  if (Array.isArray(value)) return value.map(serializeForExport);
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([key, item]) => [key, serializeForExport(item)]));
  }
  return value;
}

function downloadJson(filename, value) {
  const url = URL.createObjectURL(new Blob([JSON.stringify(value, null, 2)], { type: 'application/json' }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
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
  if (!isAdmin()) return;
  attachUsersView();
}

function attachUsersView() {
  const userConstraints = [];
  if (state.companyId) userConstraints.push(where('companyId', '==', state.companyId));
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
    const profilePayload = {
      name: prettifyHumanName(els.userNameInput?.value.trim()),
      email: els.userEmailInput?.value.trim().toLowerCase(),
      role: els.userRoleInput?.value,
      active: els.userActiveInput?.value === 'true',
      updatedAt: serverTimestamp(),
    };
    // Auto-assign current user's companyId to new profiles
    if (state.companyId) profilePayload.companyId = state.companyId;

    await setDoc(doc(db, 'users', uid), profilePayload, { merge: true });

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
  }
}

const PUNCH_EDIT_ROLES = new Set([
  'admin',
  'manager',
  'supervisor',
  'super_admin',
  'owner',
  'agency_admin',
]);

const ADMIN_ROLES = new Set(['admin', 'super_admin', 'owner']);
const PUNCH_DELETE_ROLES = new Set(['admin', 'manager', 'super_admin', 'owner']);

function normalizeRole(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[\s-]+/g, '_');
}

function normalizedRole() {
  return normalizeRole(state.profile?.role);
}

function canEditPunches() {
  return PUNCH_EDIT_ROLES.has(normalizedRole());
}

function canManageEmployees() {
  return PUNCH_EDIT_ROLES.has(normalizedRole());
}

function canDeletePunches() {
  return PUNCH_DELETE_ROLES.has(normalizedRole());
}

function isManager() {
  return canEditPunches();
}

function isAdmin() {
  return ADMIN_ROLES.has(normalizedRole());
}

function renderPermissionDebug() {
  if (!els.permissionDebugPanel) return;
  els.permissionDebugPanel.classList.toggle('hidden', !isAdmin());
  if (!isAdmin()) return;
  const rows = [
    ['UID', state.me?.uid || '-'],
    ['Email', state.me?.email || state.profile?.email || '-'],
    ['Role', normalizedRole() || '-'],
    ['companyId', state.companyId || '(blank legacy scope)'],
    ['agencyId', state.agencyId || '(none)'],
    ['canEditPunches', String(canEditPunches())],
    ['canManageEmployees', String(canManageEmployees())],
  ];
  els.permissionDebugPanel.innerHTML = rows.map(([label, value]) => `
    <div class="permission-debug-row">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join('');
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
