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
  startAfter,
  onSnapshot,
  Timestamp
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// Branch deployment config. Change CURRENT_SITE_ID to "OHC" when deploying an OHC-specific manager app.
const CURRENT_COMPANY_ID = 'chadwell';
const CURRENT_SITE_ID = 'OH01';
const OWNER_ADMIN_EMAIL = 'brandon.evanshine@chadwellsupply.com';
const APP_VERSION = 'v2026.07.05-phase1';
const featureFlags = Object.freeze({
  newDashboard: false,
  payrollReports: false,
  employeeProfiles: false,
  editPunchRequests: false,
  aiSupervisor: false
});
const BRANCH_OPTIONS = [
  { siteId: 'OH01', label: 'OH01' },
  { siteId: 'OHC', label: 'OHC' },
];

window.QRTimeClockFeatureFlags = featureFlags;

const state = {
  me: null,
  profile: null,
  companyId: CURRENT_COMPANY_ID,
  siteId: CURRENT_SITE_ID,
  allowedBranches: [],
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
  selectedWeekPunchRowsLoaded: false,
  selectedWeekTimesheetDocs: {},
  allEmployees: [],
  allUsers: [],
  allMissedRequests: [],
  approvalFilter: 'pending',
  creatingPendingProfile: false,
  workerPunchSaving: false,
  employeeStatusFilter: 'active',
  duplicateGroups: [],
  loggedWorkerNameDiagnostics: new Set(),
  loggedPayrollCanonicalChecks: new Set(),
  firestoreReadCounters: {},
  employeeListCache: new Map(),
  weeklyDataCache: new Map(),
  loadedTabs: new Set(),
  managerTimeLookup: null,
  agencyReview: {
    page: 1,
    pageSize: 25,
    sortBy: 'name',
    sortDir: 'asc',
    selectedKey: '',
    activeTab: 'overview',
    filteredRows: [],
    rangePunchRows: null,
    auditRows: [],
    auditLoadingFor: '',
    visibleColumns: null,
    deletedPunchRows: [],
  },
  agencyEmployeeRows: [],
  agencyWorkerProfileRows: [],
  agencyEmployeeCursor: null,
  agencyWorkerCursor: null,
  agencyEmployeesExhausted: false,
  agencyWorkersExhausted: false,
  agencyEmployeesLoading: false,
  workerLocation: { locationStatus: 'not_requested' },
  siteContext: buildLegacySiteContext(),
};

const EMPLOYEE_CACHE_TTL_MS = 10 * 60 * 1000;
const WEEKLY_CACHE_TTL_MS = 2 * 60 * 1000;
const DEV_READ_COUNTERS_ENABLED = location.hostname === 'localhost'
  || location.hostname === '127.0.0.1'
  || location.hostname.endsWith('.github.io');

function resetFirestoreReadCounter(scope) {
  if (!DEV_READ_COUNTERS_ENABLED || !scope) return;
  state.firestoreReadCounters[scope] = 0;
  console.info(`[QRTimeclock Firestore reads] ${scope}: start`);
}

function countFirestoreReads(scope, snapshotOrCount) {
  if (!DEV_READ_COUNTERS_ENABLED || !scope) return;
  const count = typeof snapshotOrCount === 'number'
    ? snapshotOrCount
    : Number(snapshotOrCount?.docs?.length || 0);
  state.firestoreReadCounters[scope] = Number(state.firestoreReadCounters[scope] || 0) + count;
  console.info(`[QRTimeclock Firestore reads] ${scope}: +${count}, total ${state.firestoreReadCounters[scope]}`);
}

async function getDocsCounted(queryRef, scope) {
  const snap = await getDocs(queryRef);
  countFirestoreReads(scope, snap);
  return snap;
}

async function getDocCounted(docRef, scope) {
  const snap = await getDoc(docRef);
  countFirestoreReads(scope, snap.exists() ? 1 : 0);
  return snap;
}

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
  workerLocationBtn: document.getElementById('workerLocationBtn'),
  workerLocationStatus: document.getElementById('workerLocationStatus'),

  authCard: document.getElementById('authCard'),
  appShell: document.getElementById('appShell'),
  sessionChip: document.getElementById('sessionChip'),
  sessionName: document.getElementById('sessionName'),
  sessionRole: document.getElementById('sessionRole'),
  signOutBtn: document.getElementById('signOutBtn'),
  loginForm: document.getElementById('loginForm'),
  emailInput: document.getElementById('emailInput'),
  passwordInput: document.getElementById('passwordInput'),
  legalConsent: document.getElementById('legalConsent'),
  loginSubmitBtn: document.getElementById('loginSubmitBtn'),
  resetPasswordBtn: document.getElementById('resetPasswordBtn'),
  signupForm: document.getElementById('signupForm'),
  signupNameInput: document.getElementById('signupNameInput'),
  signupEmailInput: document.getElementById('signupEmailInput'),
  signupPasswordInput: document.getElementById('signupPasswordInput'),
  signupRequestedRoleInput: document.getElementById('signupRequestedRoleInput'),
  signupSiteInput: document.getElementById('signupSiteInput'),
  signupAgencyNameInput: document.getElementById('signupAgencyNameInput'),
  signupTermsInput: document.getElementById('signupTermsInput'),
  signedInAccessRequestCard: document.getElementById('signedInAccessRequestCard'),
  signedInAccessRequestForm: document.getElementById('signedInAccessRequestForm'),
  signedInRequestNameInput: document.getElementById('signedInRequestNameInput'),
  signedInRequestEmailInput: document.getElementById('signedInRequestEmailInput'),
  signedInRequestRoleInput: document.getElementById('signedInRequestRoleInput'),
  signedInRequestSiteInput: document.getElementById('signedInRequestSiteInput'),
  signedInRequestAgencyNameInput: document.getElementById('signedInRequestAgencyNameInput'),
  signedInRequestTermsInput: document.getElementById('signedInRequestTermsInput'),

  livePunchBody: document.getElementById('livePunchBody'),
  activeNowList: document.getElementById('activeNowList'),
  gpsVerifiedCount: document.getElementById('gpsVerifiedCount'),
  gpsDeniedCount: document.getElementById('gpsDeniedCount'),
  gpsOutsideCount: document.getElementById('gpsOutsideCount'),
  gpsLowAccuracyCount: document.getElementById('gpsLowAccuracyCount'),
  timesheetBody: document.getElementById('timesheetBody'),
  weekPicker: document.getElementById('weekPicker'),
  managerTabBtn: document.getElementById('managerTabBtn'),
  timesheetsTabBtn: document.getElementById('timesheetsTabBtn'),
  editPunchesTabBtn: document.getElementById('editPunchesTabBtn'),
  reportsTabBtn: document.getElementById('reportsTabBtn'),
  adminTabBtn: document.getElementById('adminTabBtn'),
  agencyTabBtn: document.getElementById('agencyTabBtn'),
  tabBar: document.getElementById('tabBar'),
  appVersion: document.getElementById('appVersion'),

  manualPunchForm: document.getElementById('manualPunchForm'),
  manualPunchNameInput: document.getElementById('manualPunchNameInput'),
  manualPunchActionInput: document.getElementById('manualPunchActionInput'),
  manualPunchDateInput: document.getElementById('manualPunchDateInput'),
  manualPunchTimeInput: document.getElementById('manualPunchTimeInput'),
  editWeekPicker: document.getElementById('editWeekPicker'),
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
  fixUserBranchDataBtn: document.getElementById('fixUserBranchDataBtn'),
  userBranchCleanupPreview: document.getElementById('userBranchCleanupPreview'),

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
  duplicateWorkerWarning: document.getElementById('duplicateWorkerWarning'),
  workerNameRepairStatus: document.getElementById('workerNameRepairStatus'),
  repairWorkerIdInput: document.getElementById('repairWorkerIdInput'),
  repairWorkerNameInput: document.getElementById('repairWorkerNameInput'),
  repairWeekStartInput: document.getElementById('repairWeekStartInput'),
  repairRenameWorkerBtn: document.getElementById('repairRenameWorkerBtn'),
  repairCopiedNamesBtn: document.getElementById('repairCopiedNamesBtn'),
  repairKnownErvinBtn: document.getElementById('repairKnownErvinBtn'),
  exportBackupBtn: document.getElementById('exportBackupBtn'),
  refreshDuplicatesBtn: document.getElementById('refreshDuplicatesBtn'),
  duplicateWorkersList: document.getElementById('duplicateWorkersList'),
  permissionDebugPanel: document.getElementById('permissionDebugPanel'),
  siteSettingsForm: document.getElementById('siteSettingsForm'),
  siteSettingsId: document.getElementById('siteSettingsId'),
  siteSettingsQrSlug: document.getElementById('siteSettingsQrSlug'),
  siteSettingsLatitude: document.getElementById('siteSettingsLatitude'),
  siteSettingsLongitude: document.getElementById('siteSettingsLongitude'),
  siteSettingsRadius: document.getElementById('siteSettingsRadius'),
  siteSettingsAccuracy: document.getElementById('siteSettingsAccuracy'),
  siteSettingsEnforce: document.getElementById('siteSettingsEnforce'),
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
  agencyLegacyWorkerSelect: document.getElementById('agencyLegacyWorkerSelect'),
  agencyPreviewBtn: document.getElementById('agencyPreviewBtn'),
  agencyPrintBtn: document.getElementById('agencyPrintBtn'),
  agencyLegacyPrintBtn: document.getElementById('agencyLegacyPrintBtn'),
  agencyPreview: document.getElementById('agencyPreview'),
  agencyWeekPicker: document.getElementById('agencyWeekPicker'),
  agencySearchInput: document.getElementById('agencySearchInput'),
  agencyDateFilter: document.getElementById('agencyDateFilter'),
  agencyFromDateFilter: document.getElementById('agencyFromDateFilter'),
  agencyToDateFilter: document.getElementById('agencyToDateFilter'),
  agencyAgencyFilter: document.getElementById('agencyAgencyFilter'),
  agencyBranchFilter: document.getElementById('agencyBranchFilter'),
  agencyDepartmentFilter: document.getElementById('agencyDepartmentFilter'),
  agencyStatusFilter: document.getElementById('agencyStatusFilter'),
  agencyMissingClockOutFilter: document.getElementById('agencyMissingClockOutFilter'),
  agencyMissingLunchFilter: document.getElementById('agencyMissingLunchFilter'),
  agencyOvertimeFilter: document.getElementById('agencyOvertimeFilter'),
  agencyRemovedFilter: document.getElementById('agencyRemovedFilter'),
  agencyPageSizeSelect: document.getElementById('agencyPageSizeSelect'),
  agencyColumnChooser: document.getElementById('agencyColumnChooser'),
  agencyReviewTable: document.getElementById('agencyReviewTable'),
  agencyReviewBody: document.getElementById('agencyReviewBody'),
  agencyPrevPageBtn: document.getElementById('agencyPrevPageBtn'),
  agencyNextPageBtn: document.getElementById('agencyNextPageBtn'),
  agencyPageInfo: document.getElementById('agencyPageInfo'),
  agencyLoadMoreEmployeesBtn: document.getElementById('agencyLoadMoreEmployeesBtn'),
  agencyRosterPageStatus: document.getElementById('agencyRosterPageStatus'),
  agencyEmployeePanel: document.getElementById('agencyEmployeePanel'),
  agencyPanelTitle: document.getElementById('agencyPanelTitle'),
  agencyPanelMeta: document.getElementById('agencyPanelMeta'),
  agencyPanelCloseBtn: document.getElementById('agencyPanelCloseBtn'),
  agencyPanelContent: document.getElementById('agencyPanelContent'),
  agencyExportCsvBtn: document.getElementById('agencyExportCsvBtn'),
  agencyExportExcelBtn: document.getElementById('agencyExportExcelBtn'),
  agencyStatsEmployees: document.getElementById('agencyStatsEmployees'),
  agencyStatsWeekHours: document.getElementById('agencyStatsWeekHours'),
  agencyStatsMissedClockOuts: document.getElementById('agencyStatsMissedClockOuts'),
  agencyStatsMissedLunches: document.getElementById('agencyStatsMissedLunches'),
  agencyStatsLateArrivals: document.getElementById('agencyStatsLateArrivals'),
  agencyStatsAttendance: document.getElementById('agencyStatsAttendance'),
  agencyStatsMonthHours: document.getElementById('agencyStatsMonthHours'),
  agencyStatsPayPeriodHours: document.getElementById('agencyStatsPayPeriodHours'),
  agencyCoverageStatus: document.getElementById('agencyCoverageStatus'),
  agencyRestoreDeletedBtn: document.getElementById('agencyRestoreDeletedBtn'),
  agencyRecoveryStatus: document.getElementById('agencyRecoveryStatus'),

  toast: document.getElementById('toast'),
};

init();

function applyPhase1ShellPolish() {
  if (els.appVersion) els.appVersion.textContent = APP_VERSION;
  const topbarCopy = document.querySelector('.topbar p');
  if (topbarCopy && !state.me) topbarCopy.textContent = 'Simple mobile time punches with live manager visibility.';
  const workerIntro = document.querySelector('#workerCard .card-head p');
  if (workerIntro) workerIntro.textContent = 'Choose your branch, enter your name, then tap one punch button.';
}

async function init() {
  resetFirestoreReadCounter('startup');
  applyPhase1ShellPolish();
  wireEvents();
  setupBranchSelectors();
  loadSiteContext();

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

  state.selectedWeekStart = normalizeWeekStartDate(state.selectedWeekStart);
  syncSelectedWeekInputs();

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
      state.loadedTabs.clear();
      showLoggedOut();
      return;
    }

    try {
      if (state.creatingPendingProfile) return;
      state.me = user;
      const profileSnap = await getDocCounted(doc(db, 'users', user.uid), 'startup');

      if (!profileSnap.exists()) {
        state.profile = normalizeUserProfile({
          uid: user.uid,
          email: user.email || '',
          name: user.displayName || user.email || '',
          role: '',
          active: true,
          companyId: CURRENT_COMPANY_ID,
          siteId: CURRENT_SITE_ID,
          siteIds: [CURRENT_SITE_ID],
          permissions: {},
        }, user);
        state.companyId = CURRENT_COMPANY_ID;
        state.siteId = CURRENT_SITE_ID;
        state.agencyId = null;
        showLoggedIn();
        document.querySelectorAll('.tab-panel').forEach((panel) => panel.classList.add('hidden'));
        renderSignedInAccessRequestCard();
        toast('Your login exists, but your profile has not been created. Please request access.', true);
        return;
      }

      const rawProfile = profileSnap.data();
      state.profile = normalizeUserProfile(rawProfile, user);
      if (typeof rawProfile.siteIds === 'string') {
        toast('Profile branch data needs cleanup: siteIds must be an array.', true);
      }
      state.companyId = CURRENT_COMPANY_ID;
      state.siteId = CURRENT_SITE_ID;
      state.allowedBranches = [];
      state.agencyId = state.profile.agencyId || null;
      if (state.profile.companyId && state.profile.companyId !== CURRENT_COMPANY_ID) {
        await signOut(auth);
        toast('You do not have access to this company.', true);
        return;
      }
      if (state.profile.approvalStatus === 'pending') {
        showAccessRequestOnly('Your access request is pending approval from Brandon.');
        return;
      }
      if (state.profile.approvalStatus === 'denied') {
        showAccessRequestOnly('Your access request was denied. Contact your admin.');
        return;
      }
      if (state.profile.active !== true) {
        showAccessRequestOnly('Your account is inactive. Contact your admin.');
        return;
      }

      if (!isAllowedDashboardRole(state.profile)) {
        showAccessRequestOnly(`Role "${normalizeRole(state.profile.role) || 'unknown'}" has no dashboard permissions.`);
        return;
      }

      const branchAccess = resolveUserBranchAccess(state.profile);
      logBranchDebug(state.me?.email || state.profile?.email, state.profile.role, branchAccess.allowedBranches, branchAccess.activeBranch, branchAccess.valid);
      if (!branchAccess.allowedBranches.length) {
        showAccessRequestOnly('Your account does not have a branch assigned. Contact an admin.');
        return;
      }
      if (!branchAccess.valid) {
        showAccessRequestOnly('Your account does not have a branch assigned. Contact an admin.');
        return;
      }
      state.allowedBranches = branchAccess.allowedBranches;
      state.siteId = branchAccess.activeBranch;

      // Load company doc if companyId exists
      if (state.companyId) {
        try {
          const compSnap = await getDocCounted(doc(db, 'companies', state.companyId), 'startup');
          state.companyDoc = compSnap.exists() ? compSnap.data() : null;
        } catch (_) {
          state.companyDoc = null;
        }
      }

      showLoggedIn();
      attachRoleViews();
      renderPermissionDebug();
      renderSiteSettingsForm();
      if (canEditPunches()) {
        attachManagerLiveViews();
      }
      if (canManageUsers()) {
        attachUsersViewIfAdmin();
        attachPendingUsersViewIfAdmin();
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
  els.workerLocationBtn?.addEventListener('click', requestOptionalWorkerLocation);
  els.workerFixForm?.addEventListener('submit', handlePublicTimeFixSubmit);

  // Close autocomplete on outside click
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.autocomplete-wrap')) {
      hideAutocomplete();
    }
  });

  els.loginForm?.addEventListener('submit', handleLogin);
  els.signupForm?.addEventListener('submit', handleSignupRequest);
  els.signedInAccessRequestForm?.addEventListener('submit', handleSignedInAccessRequest);
  els.legalConsent?.addEventListener('change', syncLoginConsent);
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

  syncLoginConsent();

  els.weekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = normalizeWeekStartDate(els.weekPicker.value);
    syncSelectedWeekInputs();
    if (state.me && isManager()) {
      loadSelectedWeekForManager({ force: true });
    }
  });

  els.tabBar?.addEventListener('click', (event) => {
    const btn = event.target.closest('.tab');
    if (!btn) return;
    switchTab(btn.dataset.tab);
  });

  els.manualPunchForm?.addEventListener('submit', handleManualPunchSubmit);

  els.editFilterNameInput?.addEventListener('input', () => {
    renderEditPunchesTable(getEditablePunchRows());
  });
  els.editWeekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = normalizeWeekStartDate(els.editWeekPicker.value);
    loadSelectedWeekForManager({ force: true });
  });

  els.userProfileForm?.addEventListener('submit', handleSaveProfile);
  els.siteSettingsForm?.addEventListener('submit', handleSaveSiteSettings);

  els.missedPunchForm?.addEventListener('submit', handleMissedPunchSubmit);

  els.myTimecardWeekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = normalizeWeekStartDate(els.myTimecardWeekPicker.value);
    syncSelectedWeekInputs();
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
  els.repairRenameWorkerBtn?.addEventListener('click', renameWorkerProfileFromRepairTool);
  els.repairCopiedNamesBtn?.addEventListener('click', repairCopiedNamesFromWorkerProfile);
  els.repairKnownErvinBtn?.addEventListener('click', prepareKnownErvinWilsonRepair);
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
  els.agencyPrintBtn?.addEventListener('click', () => exportAgencyReviewPdf());
  els.agencyLegacyPrintBtn?.addEventListener('click', () => printAgencyPreview());
  els.agencyLegacyWorkerSelect?.addEventListener('change', () => renderAgencyPreview());
  els.agencyWeekPicker?.addEventListener('change', () => {
    state.selectedWeekStart = normalizeWeekStartDate(els.agencyWeekPicker.value);
    state.agencyReview.page = 1;
    loadSelectedWeekForManager({ force: true });
  });
  els.agencySearchInput?.addEventListener('input', () => {
    state.agencyReview.page = 1;
    renderAgencyWorkbench();
  });
  [
    els.agencyDateFilter,
    els.agencyFromDateFilter,
    els.agencyToDateFilter,
    els.agencyWorkerSelect,
    els.agencyAgencyFilter,
    els.agencyBranchFilter,
    els.agencyDepartmentFilter,
    els.agencyStatusFilter,
    els.agencyMissingClockOutFilter,
    els.agencyMissingLunchFilter,
    els.agencyOvertimeFilter,
    els.agencyRemovedFilter,
  ].forEach((control) => {
    control?.addEventListener('change', () => {
      state.agencyReview.page = 1;
      handleAgencyDateRangeChange();
    });
  });
  els.agencyPageSizeSelect?.addEventListener('change', () => {
    state.agencyReview.pageSize = Number(els.agencyPageSizeSelect.value || 25);
    state.agencyReview.page = 1;
    renderAgencyWorkbench();
  });
  els.agencyPrevPageBtn?.addEventListener('click', () => {
    state.agencyReview.page = Math.max(1, state.agencyReview.page - 1);
    renderAgencyWorkbench();
  });
  els.agencyNextPageBtn?.addEventListener('click', () => {
    state.agencyReview.page += 1;
    renderAgencyWorkbench();
  });
  els.agencyLoadMoreEmployeesBtn?.addEventListener('click', () => loadAgencyEmployeePage());
  els.agencyReviewTable?.querySelectorAll('th[data-sort]').forEach((header) => {
    header.addEventListener('click', () => sortAgencyReview(header.dataset.sort));
  });
  els.agencyPanelCloseBtn?.addEventListener('click', closeAgencyPanel);
  document.querySelectorAll('[data-agency-panel-tab]').forEach((button) => {
    button.addEventListener('click', () => switchAgencyPanelTab(button.dataset.agencyPanelTab));
  });
  els.agencyExportCsvBtn?.addEventListener('click', () => exportAgencyReviewCsv());
  els.agencyExportExcelBtn?.addEventListener('click', () => exportAgencyReviewExcel());
  els.agencyRestoreDeletedBtn?.addEventListener('click', restoreAgencySoftDeletedPunches);
  document.addEventListener('keydown', handleAgencyKeyboardShortcuts);
  els.fixUserBranchDataBtn?.addEventListener('click', previewUserBranchCleanup);
}

// ─── Public employee loading & autocomplete ─────────────
function setupBranchSelectors() {
  populateBranchSelect(els.workerBranchSelect, localStorage.getItem('workerPunchSiteId') || CURRENT_SITE_ID);
  populateBranchSelect(els.signupSiteInput, CURRENT_SITE_ID);
  populateBranchSelect(els.signedInRequestSiteInput, CURRENT_SITE_ID);
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

function normalizeSiteId(value, fallback = CURRENT_SITE_ID) {
  const siteId = String(value || '').trim();
  if (BRANCH_OPTIONS.some((branch) => branch.siteId === siteId)) return siteId;
  return BRANCH_OPTIONS.some((branch) => branch.siteId === fallback) ? fallback : CURRENT_SITE_ID;
}

function firstPresent(...values) {
  return values.find((value) => String(value || '').trim()) || '';
}

function normalizeWeekStartDate(value = new Date()) {
  const date = value instanceof Date ? value : new Date(`${value}T00:00:00`);
  const validDate = Number.isFinite(date.getTime()) ? date : new Date();
  return getMondayDate(validDate);
}

function syncSelectedWeekInputs() {
  const weekValue = formatDateInput(state.selectedWeekStart);
  if (els.weekPicker) els.weekPicker.value = weekValue;
  if (els.editWeekPicker) els.editWeekPicker.value = weekValue;
  if (els.agencyWeekPicker) els.agencyWeekPicker.value = weekValue;
  if (els.repairWeekStartInput) els.repairWeekStartInput.value = weekValue;
  if (els.myTimecardWeekPicker) els.myTimecardWeekPicker.value = weekValue;
}

function loadSelectedWeekForManager({ force = true, updateAgencyRange = true } = {}) {
  state.selectedWeekStart = normalizeWeekStartDate(state.selectedWeekStart);
  syncSelectedWeekInputs();
  if (updateAgencyRange) {
    state.agencyReview.rangePunchRows = null;
    if (els.agencyDateFilter) els.agencyDateFilter.value = '';
    if (els.agencyFromDateFilter) els.agencyFromDateFilter.value = '';
    if (els.agencyToDateFilter) els.agencyToDateFilter.value = '';
  }
  attachTimesheetView({ force });
}

function getCurrentCompanyId() {
  return CURRENT_COMPANY_ID;
}

function getCurrentSiteId() {
  return state.siteId || CURRENT_SITE_ID;
}

function getAllowedSiteIds(profile = state.profile) {
  if (!profile) return [];
  if (Array.isArray(profile.branches) && profile.branches.length) {
    return parseSiteIds(profile.branches, false);
  }
  if (profile.branch) {
    return parseSiteIds([profile.branch], false);
  }
  if (Array.isArray(profile.siteIds) && profile.siteIds.length) {
    return parseSiteIds(profile.siteIds, false);
  }
  const singleSiteId = normalizeCleanupSiteId(profile.siteId);
  return singleSiteId ? [singleSiteId] : [];
}

function parseSiteIds(value, fallbackToCurrent = true) {
  const raw = Array.isArray(value) ? value : String(value || '').split(',');
  const sites = raw
    .map(normalizeCleanupSiteId)
    .filter(Boolean);
  if (!sites.length && !fallbackToCurrent) return [];
  return [...new Set(sites.length ? sites : [CURRENT_SITE_ID])];
}

function profileCanAccessSite(profile, siteId) {
  return String(profile?.companyId || '').trim() === CURRENT_COMPANY_ID
    && getAllowedSiteIds(profile).includes(siteId);
}

function isAllowedDashboardRole(profile = state.profile) {
  return [
    'manager',
    'admin',
    'supervisor',
    'agency_admin',
    'owner',
    'superadmin',
    'super_admin',
  ].includes(normalizeRole(profile?.role));
}

function resolveUserBranchAccess(profile) {
  const allowedBranches = getAllowedSiteIds(profile);
  const storedBranch = sessionStorage.getItem(`managerActiveBranch:${profile?.uid || ''}`);
  const requestedBranch = normalizeCleanupSiteId(storedBranch);
  const activeBranch = allowedBranches.includes(requestedBranch)
    ? requestedBranch
    : allowedBranches[0];
  return {
    allowedBranches,
    activeBranch,
    valid: !!activeBranch && allowedBranches.includes(activeBranch),
  };
}

function logBranchDebug(email, role, allowedBranches, activeBranch, valid) {
  console.info('[QRTimeclock branch access]', {
    signedInEmail: email || '',
    role: normalizeRole(role),
    allowedBranches,
    activeBranch,
    branchValidationResult: valid,
  });
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
    siteId: normalizeSiteId(siteId),
  };
}

function scopedPunchHistoryConstraints(siteId = getCurrentSiteId()) {
  const constraints = [...branchConstraints(siteId)];
  if (state.me && isAgencyUser()) constraints.push(where('agencyId', '==', agencyScopeId()));
  return constraints;
}

function validatePunchPayloadForSave(payload, { requireEmployeeId = true } = {}) {
  const errors = [];
  if (!String(payload?.name || payload?.workerName || payload?.employeeName || '').trim()) errors.push('worker name');
  if (!String(payload?.nameKey || '').trim()) errors.push('nameKey');
  if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(payload?.action)) errors.push('valid punch action');
  if (!Number.isFinite(Number(payload?.timestampMs)) || Number(payload?.timestampMs) <= 0) errors.push('valid timestamp');
  if (!String(payload?.dateKey || '').trim()) errors.push('dateKey');
  if (!String(payload?.weekKey || '').trim()) errors.push('weekKey');
  if (payload?.companyId !== getCurrentCompanyId()) errors.push('companyId');
  if (!BRANCH_OPTIONS.some((branch) => branch.siteId === payload?.siteId)) errors.push('siteId');
  if (requireEmployeeId && !String(payload?.employeeId || payload?.workerId || '').trim()) errors.push('employeeId');
  if (errors.length) {
    throw new Error(`Punch was not saved. Missing or invalid: ${errors.join(', ')}.`);
  }
}

function employeeCacheKey({ siteId = getCurrentSiteId(), agencyId = '', status = '', collectionName = 'employees' } = {}) {
  return [
    collectionName,
    getCurrentCompanyId(),
    normalizeSiteId(siteId),
    normalizeIdentityToken(agencyId),
    String(status || '').toLowerCase()
  ].join('|');
}

function readSessionEmployeeCache(key) {
  try {
    const cached = JSON.parse(sessionStorage.getItem(`qrTimeclock:employeeCache:${key}`) || 'null');
    if (!cached || Date.now() - Number(cached.savedAt || 0) > EMPLOYEE_CACHE_TTL_MS) return null;
    return Array.isArray(cached.rows) ? cached.rows : null;
  } catch (_) {
    return null;
  }
}

function writeSessionEmployeeCache(key, rows) {
  try {
    sessionStorage.setItem(`qrTimeclock:employeeCache:${key}`, JSON.stringify({
      savedAt: Date.now(),
      rows
    }));
  } catch (_) {}
}

function invalidateEmployeeCaches() {
  state.employeeListCache.clear();
  state.weeklyDataCache.clear();
  try {
    Object.keys(sessionStorage)
      .filter((key) => key.startsWith('qrTimeclock:employeeCache:'))
      .forEach((key) => sessionStorage.removeItem(key));
  } catch (_) {}
}

async function loadBranchEmployees({ siteId = getCurrentSiteId(), agencyId = isAgencyUser() ? agencyScopeId() : '', status = '', force = false, scope = 'employees' } = {}) {
  const cacheKey = employeeCacheKey({ siteId, agencyId, status });
  const cached = state.employeeListCache.get(cacheKey);
  if (!force && cached && Date.now() - cached.savedAt < EMPLOYEE_CACHE_TTL_MS) return cached.rows;

  if (!force) {
    const sessionRows = readSessionEmployeeCache(cacheKey);
    if (sessionRows) {
      state.employeeListCache.set(cacheKey, { savedAt: Date.now(), rows: sessionRows });
      return sessionRows;
    }
  }

  const constraints = [
    ...branchConstraints(siteId),
  ];
  if (status) constraints.push(where('status', '==', status));
  if (agencyId) constraints.push(where('agencyId', '==', agencyId));

  const snap = await getDocsCounted(query(collection(db, 'employees'), ...constraints), scope);
  const rows = snap.docs
    .map((d) => ({ id: d.id, ...d.data() }))
    .sort((a, b) => (a.name || '').localeCompare(b.name || ''));
  state.employeeListCache.set(cacheKey, { savedAt: Date.now(), rows });
  writeSessionEmployeeCache(cacheKey, rows);
  return rows;
}

async function loadPublicEmployees() {
  try {
    const activeEmployees = await loadBranchEmployees({
      siteId: getPublicSiteId(),
      status: 'active',
      agencyId: '',
      scope: 'startup'
    });
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

function buildLegacySiteContext() {
  const params = new URLSearchParams(window.location.search);
  const companyId = params.get('company') || params.get('companyId') || '';
  const siteId = params.get('site') || params.get('siteId') || '';
  const qrSlug = params.get('qr') || params.get('qrSlug') || '';
  return {
    companyId,
    siteId,
    qrSlug,
    legacyMode: !companyId && !siteId && !qrSlug,
    siteLatitude: null,
    siteLongitude: null,
    allowedRadiusMeters: 300,
    maxGpsAccuracyMeters: 100,
    enforceLocation: false,
  };
}

async function loadSiteContext() {
  const siteId = state.siteContext.siteId;
  if (!siteId) return;
  try {
    const snapshot = await getDocCounted(doc(db, 'sites', siteId), 'startup');
    if (!snapshot.exists()) return;
    const site = snapshot.data();
    state.siteContext = {
      ...state.siteContext,
      companyId: state.siteContext.companyId || site.companyId || '',
      qrSlug: state.siteContext.qrSlug || site.qrSlug || '',
      siteLatitude: finiteNumberOrNull(site.siteLatitude),
      siteLongitude: finiteNumberOrNull(site.siteLongitude),
      allowedRadiusMeters: positiveNumberOrDefault(site.allowedRadiusMeters, 300),
      maxGpsAccuracyMeters: positiveNumberOrDefault(site.maxGpsAccuracyMeters, 100),
      // Safe mode: location enforcement is deliberately disabled.
      enforceLocation: false,
    };
    renderSiteSettingsForm();
  } catch (error) {
    console.warn('Optional site settings could not be loaded:', error.message);
  }
}

function requestOptionalWorkerLocation() {
  if (!navigator.geolocation) {
    state.workerLocation = {
      locationStatus: 'unavailable',
      locationCapturedAtMs: Date.now(),
    };
    updateWorkerLocationStatus('GPS is unavailable. You can still punch normally.');
    return Promise.resolve(state.workerLocation);
  }

  if (els.workerLocationBtn) els.workerLocationBtn.disabled = true;
  updateWorkerLocationStatus('Requesting optional GPS location...');
  return new Promise((resolve) => {
    navigator.geolocation.getCurrentPosition(
      (position) => {
        state.workerLocation = buildVerifiedLocation(position.coords);
        updateWorkerLocationStatus(formatWorkerLocationMessage(state.workerLocation));
        if (els.workerLocationBtn) els.workerLocationBtn.textContent = 'Refresh GPS (Optional)';
        if (els.workerLocationBtn) els.workerLocationBtn.disabled = false;
        resolve(state.workerLocation);
      },
      (error) => {
        state.workerLocation = {
          locationStatus: error?.code === 1 ? 'denied' : 'unavailable',
          locationCapturedAtMs: Date.now(),
        };
        updateWorkerLocationStatus(
          error?.code === 1
            ? 'Location denied. Your punch will still save.'
            : 'Location could not be captured. Your punch will still save.'
        );
        if (els.workerLocationBtn) els.workerLocationBtn.disabled = false;
        resolve(state.workerLocation);
      },
      {
        enableHighAccuracy: true,
        timeout: 7000,
        maximumAge: 60000,
      }
    );
  });
}

function buildVerifiedLocation(coords) {
  const latitude = finiteNumberOrNull(coords?.latitude);
  const longitude = finiteNumberOrNull(coords?.longitude);
  const gpsAccuracyMeters = finiteNumberOrNull(coords?.accuracy);
  const hasSiteCoordinates = Number.isFinite(state.siteContext.siteLatitude) &&
    Number.isFinite(state.siteContext.siteLongitude);
  const distanceFromSiteMeters = hasSiteCoordinates && latitude !== null && longitude !== null
    ? haversineDistanceMeters(
      latitude,
      longitude,
      state.siteContext.siteLatitude,
      state.siteContext.siteLongitude
    )
    : null;
  return {
    latitude,
    longitude,
    gpsAccuracyMeters,
    locationStatus: 'verified',
    locationCapturedAtMs: Date.now(),
    distanceFromSiteMeters,
    withinAllowedRadius: distanceFromSiteMeters === null
      ? null
      : distanceFromSiteMeters <= state.siteContext.allowedRadiusMeters,
  };
}

function buildPunchLocationPayload(employee, preferredSiteId = '') {
  const location = state.workerLocation || { locationStatus: 'not_requested' };
  const siteId = normalizeSiteId(firstPresent(preferredSiteId, employee?.assignedSiteId, employee?.siteId, state.siteContext.siteId), getPublicSiteId());
  const employeeSiteIds = Array.isArray(employee?.siteIds) ? parseSiteIds(employee.siteIds, false) : [];
  const capturedAtMs = Number(location.locationCapturedAtMs || 0);
  return {
    latitude: finiteNumberOrNull(location.latitude),
    longitude: finiteNumberOrNull(location.longitude),
    accuracy: finiteNumberOrNull(location.gpsAccuracyMeters),
    gpsAccuracyMeters: finiteNumberOrNull(location.gpsAccuracyMeters),
    locationStatus: location.locationStatus || 'not_requested',
    locationCapturedAt: capturedAtMs ? Timestamp.fromMillis(capturedAtMs) : null,
    siteId,
    siteIds: employeeSiteIds.length ? employeeSiteIds : (siteId ? [siteId] : []),
    distanceFromSiteMeters: finiteNumberOrNull(location.distanceFromSiteMeters),
    withinAllowedRadius: typeof location.withinAllowedRadius === 'boolean'
      ? location.withinAllowedRadius
      : null,
    allowedRadiusMeters: state.siteContext.allowedRadiusMeters,
    maxGpsAccuracyMeters: state.siteContext.maxGpsAccuracyMeters,
    enforceLocation: false,
  };
}

function formatWorkerLocationMessage(location) {
  const parts = ['GPS verified'];
  if (Number.isFinite(location.gpsAccuracyMeters)) {
    parts.push(`accuracy ${Math.round(location.gpsAccuracyMeters)}m`);
  }
  if (Number.isFinite(location.distanceFromSiteMeters)) {
    parts.push(`${Math.round(location.distanceFromSiteMeters)}m from site`);
  }
  parts.push('Punching remains available.');
  return parts.join(' · ');
}

function updateWorkerLocationStatus(message) {
  if (els.workerLocationStatus) els.workerLocationStatus.textContent = message;
}

function haversineDistanceMeters(latitudeA, longitudeA, latitudeB, longitudeB) {
  const radians = (degrees) => degrees * Math.PI / 180;
  const earthRadiusMeters = 6371000;
  const deltaLatitude = radians(latitudeB - latitudeA);
  const deltaLongitude = radians(longitudeB - longitudeA);
  const a = Math.sin(deltaLatitude / 2) ** 2 +
    Math.cos(radians(latitudeA)) * Math.cos(radians(latitudeB)) *
    Math.sin(deltaLongitude / 2) ** 2;
  return earthRadiusMeters * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

function finiteNumberOrNull(value) {
  if (value === null || value === undefined || value === '') return null;
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function positiveNumberOrDefault(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) && number > 0 ? number : fallback;
}

async function handleWorkerPunch(action) {
  if (state.workerPunchSaving) return;
  resetFirestoreReadCounter('Put Away save');
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

  const publicSiteId = getPublicSiteId();
  const enteredPin = String(els.workerPinInput?.value || '').trim();
  if (enteredPin && emp.pinHash) {
    const enteredHash = await hashWorkerPin(enteredPin, emp.id || emp.employeeId || '');
    if (enteredHash !== emp.pinHash) {
      toast('PIN incorrect. You can clear it and continue with the normal name-based flow.', true);
      return;
    }
  }

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
      const empNumber = await generateNextPublicEmployeeNumber();
      const employeeSiteId = normalizeSiteId(publicSiteId, publicSiteId);
      const scope = {
        employeeNumber: empNumber,
        nameKey: normalizeName(emp.name),
        companyId: getCurrentCompanyId(),
        agencyId: '',
        siteId: employeeSiteId
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
        assignedSiteId: employeeSiteId,
        siteId: employeeSiteId,
        siteIds: [employeeSiteId],
        qrSlug: state.siteContext.qrSlug || '',
        status: 'active',
        active: true,
        employeeId,
        source: 'auto_created',
        updatedAt: serverTimestamp(),
      };
      if (!existingEmployee) {
        newPayload.createdAt = serverTimestamp();
        await setDoc(doc(db, 'employees', employeeId), newPayload, { merge: true });
        invalidateEmployeeCaches();
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
      console.warn('Auto-create employee skipped:', error.message);
      const identityHash = await hashIdentityKey(employeeScopeKey(emp));
      const existingSnap = await getDocCounted(doc(db, 'employees', `auto_${identityHash.slice(0, 24)}`), 'Put Away save');
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
  const punchSiteId = normalizeSiteId(firstPresent(emp.siteId, emp.assignedSiteId, publicSiteId), publicSiteId);
  const verifiedEmployeeId = emp.employeeId || emp.id || '';
  if (!verifiedEmployeeId) {
    toast('Your worker profile could not be verified. Ask a manager to add you before punching.', true);
    return;
  }

  state.workerPunchSaving = true;
  setWorkerPunchBusy(true);
  try {
    const duplicateKey = `lastPunch:${emp.employeeId || emp.id || nameKey}:${action}`;
    const previousPunchMs = Number(localStorage.getItem(duplicateKey) || 0);
    if (previousPunchMs && nowMs - previousPunchMs < 10000) {
      throw new Error('That punch was already saved. Please wait a few seconds.');
    }
    const payload = {
      ...buildPunchLocationPayload(emp, punchSiteId),
      ...branchPayload(punchSiteId),
      name,
      nameKey,
      action,
      timestamp: serverTimestamp(),
      timestampMs: nowMs,
      dateKey,
      weekKey,
      source: 'public_qr',
      createdAt: serverTimestamp(),
      employeeId: verifiedEmployeeId,
      employeeNumber: emp.employeeNumber || '',
      agencyId: emp.agencyId || '',
      assignedSiteId: normalizeSiteId(firstPresent(emp.assignedSiteId, emp.siteId, punchSiteId), punchSiteId),
      siteIds: [punchSiteId],
      qrSlug: state.siteContext.qrSlug || '',
    };
    validatePunchPayloadForSave(payload);
    await addDoc(collection(db, 'punches'), payload);

    cacheWorkerPunch(emp, {
      name,
      nameKey,
      action,
      timestampMs: nowMs,
      dateKey,
      weekKey,
      employeeId: verifiedEmployeeId,
    });

    const savedActionLabel = prettyAction(action);
    if (els.workerLastActionValue) els.workerLastActionValue.textContent = savedActionLabel;
    if (els.workerLastPunchValue) els.workerLastPunchValue.textContent = formatDateTime(nowMs);
    if (els.workerStatusValue) els.workerStatusValue.textContent = statusLabelForAction(action);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = `${savedActionLabel} saved for ${name} at ${formatDateTime(nowMs)}.`;
    }

    attachWorkerLiveView(name);
    localStorage.setItem('workerPunchName', name);
    localStorage.setItem(duplicateKey, String(nowMs));
    toast(`${savedActionLabel} saved successfully.`);
  } catch (error) {
    console.error(error);
    if (els.workerStatusMessage) {
      els.workerStatusMessage.textContent = formatWorkerPunchError(error);
    }
    toast(formatWorkerPunchError(error), true);
  } finally {
    state.workerPunchSaving = false;
    setWorkerPunchBusy(false);
  }
}

function formatWorkerPunchError(error) {
  const message = String(error?.message || '').trim();
  if (/already saved/i.test(message)) return 'That punch was already saved. Please wait a few seconds before trying again.';
  if (/select your/i.test(message)) return message;
  if (/permission|missing or insufficient/i.test(message)) return 'Punch could not be saved because access was denied. Ask a manager to check setup.';
  return message || 'Could not save punch. Check your connection and try again.';
}

function setWorkerPunchBusy(busy) {
  document.getElementById('workerCard')?.classList.toggle('is-saving', busy);
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
      const normalized = normalizePunchRecordForDisplay(row);
      const key = `${normalized.timestampMs || 0}:${normalized.action || ''}:${normalized.nameKey || normalizeName(normalized.name || '')}`;
      unique.set(key, normalized);
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
  resetFirestoreReadCounter('History search');
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
      emptyMessage: 'No punches found for selected range.',
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
      emptyMessage: 'No punches found for selected range.',
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
  const employeeSiteId = normalizeSiteId(firstPresent(employee.siteId, employee.assignedSiteId, getCurrentSiteId()), getCurrentSiteId());
  const historyScope = scopedPunchHistoryConstraints(employeeSiteId);
  getCompatibleWorkerRecords(employee).forEach((record) => {
    if (record.id) workerIds.add(record.id);
    if (record.employeeId) workerIds.add(record.employeeId);
  });
  if (primaryId) {
    const mergedConstraints = state.me && isManager()
      ? [where('mergedInto', '==', primaryId)]
      : [where('status', '==', 'merged'), where('mergedInto', '==', primaryId)];
    const mergedQueries = [
      query(collection(db, 'employees'), ...branchConstraints(employeeSiteId), ...mergedConstraints),
      query(
        collection(db, 'employees'),
        where('companyId', '==', getCurrentCompanyId()),
        where('assignedSiteId', '==', employeeSiteId),
        ...mergedConstraints
      ),
    ];
    const mergedResults = await Promise.allSettled(
      mergedQueries.map((mergedQuery) => getDocsCounted(mergedQuery, options.scope || 'History search'))
    );
    mergedResults.forEach((result) => {
      if (result.status !== 'fulfilled') {
        console.warn('Merged employee lookup skipped:', result.reason?.message || result.reason);
        return;
      }
      result.value.docs.forEach((record) => {
        workerIds.add(record.id);
        directWorkerIds.add(record.id);
      });
    });
  }

  const rows = [];
  let nameQuerySucceeded = false;
  if (employee.name) {
    try {
      const nameRows = await fetchPunchesWithRange(
        [...historyScope, where('nameKey', '==', normalizeName(employee.name))],
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
    fetchPunchesWithRange([...historyScope, where('employeeId', '==', workerId)], fromMs, toMs)
  ));
  idRows.forEach((workerRows) => rows.push(...workerRows));

  if (options.allowLegacyNameFallback && employee.name) {
    const legacyConstraints = state.me && isManager()
      ? []
      : [where('employeeId', '==', '')];
    const legacyRows = await fetchPunchesWithRange(
      [...historyScope, ...legacyConstraints],
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

async function fetchPunchesWithRange(baseConstraints, fromMs, toMs, options = {}) {
  const includeLegacyTimestampFallback = options.includeLegacyTimestampFallback === true;
  const scope = options.scope || 'History search';
  const rows = [];
  try {
    const rangedQuery = query(
      collection(db, 'punches'),
      ...baseConstraints,
      where('timestampMs', '>=', fromMs),
      where('timestampMs', '<=', toMs)
    );
    const snapshot = await getDocsCounted(rangedQuery, scope);
    rows.push(...snapshot.docs.map((record) => normalizePunchRecordForDisplay({ id: record.id, ...record.data() })));
  } catch (error) {
    if (!['failed-precondition', 'permission-denied'].includes(error.code)) throw error;
  }

  if (includeLegacyTimestampFallback && baseConstraints.length) {
    try {
      const snapshot = await getDocsCounted(query(collection(db, 'punches'), ...baseConstraints), scope);
      rows.push(...snapshot.docs
        .map((record) => normalizePunchRecordForDisplay({ id: record.id, ...record.data() }))
        .filter((punch) => Number(punch.timestampMs || 0) >= fromMs && Number(punch.timestampMs || 0) <= toMs));
    } catch (error) {
      console.warn('Legacy punch timestamp fallback failed:', error.message);
    }
  }

  return dedupePunches(rows);
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
    const requestSiteId = normalizeSiteId(firstPresent(employee.siteId, employee.assignedSiteId, getPublicSiteId()), getPublicSiteId());
    await addDoc(collection(db, 'missedPunchRequests'), {
      ...branchPayload(requestSiteId),
      uid: '',
      employeeId: employee.employeeId || employee.id,
      employeeNumber: employee.employeeNumber || '',
      agencyId: employee.agencyId || '',
      assignedSiteId: requestSiteId,
      siteIds: [requestSiteId],
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

function findActiveEmployeeForPunchName(nameKey, siteId = getCurrentSiteId(), agencyId = state.agencyId || '') {
  const normalizedAgencyId = String(agencyId || '').trim();
  const candidates = (state.allEmployees || []).filter((employee) => {
    if (!isActiveEmployee(employee)) return false;
    if (normalizeName(getWorkerProfileName(employee) || employee.nameKey || '') !== nameKey) return false;
    if (normalizeSiteId(firstPresent(employee.siteId, employee.assignedSiteId, siteId), siteId) !== siteId) return false;
    return String(employee.agencyId || '').trim() === normalizedAgencyId;
  });
  if (candidates.length === 1) return candidates[0];
  const collapsed = collapseDuplicateEmployees(candidates);
  return collapsed.length === 1 ? collapsed[0] : null;
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
  const employee = findActiveEmployeeForPunchName(nameKey);
  if (!employee) {
    toast('Create or select a single active employee profile before adding a manual punch.', true);
    return;
  }

  try {
    const payload = {
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
      agencyId: employee.agencyId || state.agencyId || '',
      employeeId: employee.employeeId || employee.id || '',
      workerId: employee.workerId || employee.id || employee.employeeId || '',
      employeeNumber: employee.employeeNumber || '',
    };
    validatePunchPayloadForSave(payload);
    await addDoc(collection(db, 'punches'), payload);

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
    const rows = snap.docs.map((d) => normalizePunchRecordForDisplay({ id: d.id, ...d.data() }));

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

  if (!els.legalConsent?.checked) {
    toast('You must agree to the Terms of Use and Privacy Policy.', true);
    return;
  }

  try {
    await signInWithEmailAndPassword(
      auth,
      els.emailInput?.value.trim(),
      els.passwordInput?.value
    );
    if (els.passwordInput) els.passwordInput.value = '';
    if (els.legalConsent) els.legalConsent.checked = false;
    syncLoginConsent();
  } catch (error) {
    console.error(error);
    toast(formatAuthError(error), true);
  }
}

async function handleSignupRequest(event) {
  event.preventDefault();

  const name = prettifyHumanName(els.signupNameInput?.value.trim());
  const email = String(els.signupEmailInput?.value || '').trim().toLowerCase();
  const password = els.signupPasswordInput?.value || '';
  const requestedRole = els.signupRequestedRoleInput?.value || 'manager';
  const siteId = els.signupSiteInput?.value || CURRENT_SITE_ID;
  const agencyName = els.signupAgencyNameInput?.value.trim() || '';
  const termsAccepted = els.signupTermsInput?.checked === true;

  if (!name || !email || password.length < 6 || !termsAccepted) {
    toast('Enter a name, email, password with at least 6 characters, and accept the terms.', true);
    return;
  }

  try {
    state.creatingPendingProfile = true;
    const credential = await createUserWithEmailAndPassword(auth, email, password);
    if (name) {
      await updateProfile(credential.user, { displayName: name });
    }

    await submitAccessRequestForUser(credential.user, {
      name,
      email,
      requestedRole,
      siteId,
      agencyName,
      termsAccepted,
      createIfMissing: true,
    });

    await signOut(auth);
    els.signupForm?.reset();
    populateBranchSelect(els.signupSiteInput, CURRENT_SITE_ID);
    toast('Access request sent to Brandon for approval.');
  } catch (error) {
    console.error(error);
    if (error.code === 'auth/email-already-in-use') {
      toast('This email already has an account. Please sign in, then submit your access request.', true);
    } else {
      toast(formatAuthError(error), true);
    }
  } finally {
    state.creatingPendingProfile = false;
  }
}

async function handleSignedInAccessRequest(event) {
  event.preventDefault();
  if (!state.me) {
    toast('Please sign in, then submit your access request.', true);
    return;
  }

  const name = prettifyHumanName(els.signedInRequestNameInput?.value.trim());
  const requestedRole = els.signedInRequestRoleInput?.value || 'manager';
  const siteId = els.signedInRequestSiteInput?.value || CURRENT_SITE_ID;
  const agencyName = els.signedInRequestAgencyNameInput?.value.trim() || '';
  const termsAccepted = els.signedInRequestTermsInput?.checked === true;

  if (!name || !termsAccepted) {
    toast('Enter your full name and accept the terms.', true);
    return;
  }

  try {
    await submitAccessRequestForUser(state.me, {
      name,
      email: state.me.email || state.profile?.email || '',
      requestedRole,
      siteId,
      agencyName,
      termsAccepted,
      createIfMissing: false,
    });
    if (els.signedInRequestTermsInput) els.signedInRequestTermsInput.checked = false;
    state.profile = {
      ...state.profile,
      active: false,
      approvalStatus: 'pending',
      requestedRole,
      role: 'worker',
      branch: siteId,
      branches: [siteId],
      siteId,
      siteIds: [siteId],
    };
    showAccessRequestOnly('Your access request is pending approval from Brandon.');
    toast('Access request sent to Brandon for approval.');
  } catch (error) {
    console.error(error);
    toast(formatFirestoreError(error), true);
  }
}

async function submitAccessRequestForUser(authUser, request) {
  const requestedRole = normalizeRequestedAccessRole(request.requestedRole);
  const siteId = normalizeSiteId(request.siteId);
  if (!['manager', 'supervisor', 'agency_admin'].includes(requestedRole)) {
    throw new Error('Choose manager, supervisor, or temp agency admin access.');
  }
  if (requestedRole === 'agency_admin' && !request.agencyName) {
    throw new Error('Enter an agency name for temp agency admin access.');
  }
  if (request.termsAccepted !== true) {
    throw new Error('You must accept the terms before requesting access.');
  }

  const userRef = doc(db, 'users', authUser.uid);
  const existingSnap = await getDoc(userRef);
  const existingProfile = existingSnap.exists() ? existingSnap.data() : {};
  const existingNormalizedRole = normalizeRole(existingProfile.role);
  const preserveApprovedElevatedAccess = existingProfile.active === true
    && existingProfile.approvalStatus === 'approved'
    && ['owner', 'super_admin', 'superadmin', 'admin', 'manager', 'supervisor', 'agency_admin'].includes(existingNormalizedRole);
  const agencyName = requestedRole === 'agency_admin' ? request.agencyName.trim() : '';

  const payload = {
    uid: existingProfile.uid || authUser.uid,
    name: request.name,
    displayName: request.name,
    email: String(authUser.email || request.email || existingProfile.email || '').trim().toLowerCase(),
    requestedRole,
    companyId: getCurrentCompanyId(),
    branch: siteId,
    branches: [siteId],
    siteId,
    siteIds: [siteId],
    requestedAdminEmail: OWNER_ADMIN_EMAIL,
    permissions: preserveApprovedElevatedAccess
      ? {
          canEditPunches: existingProfile.permissions?.canEditPunches === true,
          canDeletePunches: existingProfile.permissions?.canDeletePunches === true,
          canMergeWorkers: existingProfile.permissions?.canMergeWorkers === true,
          manageUsers: existingProfile.permissions?.manageUsers === true,
        }
      : {
          canEditPunches: false,
          canDeletePunches: false,
          canMergeWorkers: false,
          manageUsers: false,
        },
    termsAccepted: true,
    termsAcceptedAt: serverTimestamp(),
    accessRequestedAt: serverTimestamp(),
    updatedAt: serverTimestamp(),
  };

  if (agencyName) {
    payload.agencyName = agencyName;
    payload.agencyId = normalizeScopeId(agencyName);
  }

  if (preserveApprovedElevatedAccess) {
    payload.active = true;
    payload.approvalStatus = 'approved';
    payload.role = existingProfile.role;
  } else {
    payload.active = false;
    payload.approvalStatus = 'pending';
    payload.role = 'worker';
  }

  if (!existingSnap.exists()) {
    payload.createdAt = serverTimestamp();
  }

  await setDoc(userRef, payload, { merge: true });
}

function syncLoginConsent() {
  if (els.loginSubmitBtn) {
    els.loginSubmitBtn.disabled = !els.legalConsent?.checked;
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
  state.allowedBranches = [];
  state.agencyId = null;
  state.companyDoc = null;
  document.getElementById('managerBranchSelect')?.remove();
  els.authCard?.classList.remove('hidden');
  els.appShell?.classList.add('hidden');
  els.signedInAccessRequestCard?.classList.add('hidden');
  els.sessionChip?.classList.add('hidden');
  // Restore public worker card
  const workerCard = document.getElementById('workerCard');
  if (workerCard) workerCard.classList.remove('hidden');
  // Reset header
  const headerP = document.querySelector('.topbar p');
  if (headerP) headerP.textContent = 'Simple mobile time punches with live manager visibility.';
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
    roleParts.push(state.siteId || CURRENT_SITE_ID);
    if (state.agencyId) roleParts.push('agency');
    els.sessionRole.textContent = roleParts.join(' · ');
  }

  renderManagerBranchSwitcher();

  // Show company name in header
  const companyDisplayName = state.companyDoc?.name || (state.companyId ? state.companyId : appSettings.companyName);
  const headerP = document.querySelector('.topbar p');
  if (headerP) headerP.textContent = companyDisplayName + ' — TimeClock Pro';
}

function showAccessRequestOnly(message) {
  clearLiveListeners();
  showLoggedIn();
  document.querySelectorAll('.tab-panel').forEach((panel) => panel.classList.add('hidden'));
  document.querySelectorAll('.tab').forEach((tab) => tab.classList.remove('active'));
  renderSignedInAccessRequestCard();
  toast(message, true);
}

function renderManagerBranchSwitcher() {
  const chip = els.sessionChip;
  if (!chip) return;
  let selector = document.getElementById('managerBranchSelect');
  const branches = state.allowedBranches || [];
  if (branches.length <= 1 || !isAllowedDashboardRole(state.profile)) {
    selector?.remove();
    return;
  }
  if (!selector) {
    selector = document.createElement('select');
    selector.id = 'managerBranchSelect';
    selector.className = 'branch-switcher';
    selector.addEventListener('change', () => {
      const nextBranch = normalizeCleanupSiteId(selector.value);
      if (!branches.includes(nextBranch)) {
        toast('You do not have access to that branch.', true);
        selector.value = state.siteId || CURRENT_SITE_ID;
        return;
      }
      state.siteId = nextBranch;
      sessionStorage.setItem(`managerActiveBranch:${state.profile?.uid || ''}`, nextBranch);
      invalidateEmployeeCaches();
      state.loadedTabs.clear();
      refreshManagerDashboardForBranch();
    });
    chip.insertBefore(selector, els.signOutBtn || null);
  }
  selector.innerHTML = branches.map((branch) => `<option value="${branch}">${branch}</option>`).join('');
  selector.value = state.siteId || branches[0];
}

function refreshManagerDashboardForBranch() {
  clearLiveListeners();
  showLoggedIn();
  attachRoleViews();
  renderPermissionDebug();
  renderSiteSettingsForm();
  if (canEditPunches()) {
    attachManagerLiveViews();
  }
  if (canManageUsers()) {
    attachUsersViewIfAdmin();
    attachPendingUsersViewIfAdmin();
  }
}

function renderSignedInAccessRequestCard() {
  if (!els.signedInAccessRequestCard) return;
  const showCard = !!state.me && !canEditPunches() && !canManageUsers();
  els.signedInAccessRequestCard.classList.toggle('hidden', !showCard);
  if (!showCard) return;
  if (els.signedInRequestNameInput) {
    els.signedInRequestNameInput.value = state.profile?.name || state.me?.displayName || '';
  }
  if (els.signedInRequestEmailInput) {
    els.signedInRequestEmailInput.value = state.me?.email || state.profile?.email || '';
  }
  if (els.signedInRequestRoleInput && ['manager', 'supervisor', 'agency_admin'].includes(state.profile?.requestedRole)) {
    els.signedInRequestRoleInput.value = state.profile.requestedRole;
  }
  if (els.signedInRequestSiteInput) {
    els.signedInRequestSiteInput.value = getAllowedSiteIds(state.profile)[0] || CURRENT_SITE_ID;
  }
  if (els.signedInRequestAgencyNameInput) {
    els.signedInRequestAgencyNameInput.value = state.profile?.agencyName || '';
  }
}

function getCompanyName() {
  return state.companyDoc?.name || state.companyId || appSettings.companyName;
}

/** Returns true if current user is scoped to an agency */
function isAgencyUser() {
  return normalizedRole() === 'agency_admin' || !!state.agencyId;
}

function agencyScopeId() {
  return state.agencyId || '__missing_agency__';
}

const AGENCY_NAMES = {
  sterling_staffing: 'Sterling Staffing',
  excel_staffing: 'Excel Staffing',
};

function agencyLabel(agencyId) {
  if (!agencyId) return 'Direct';
  return AGENCY_NAMES[agencyId] || agencyId;
}

function isFeatureEnabled(flagName) {
  return featureFlags[flagName] === true;
}

function markFeatureTab(button, flagName, { visible = true, preserveCurrent = false } = {}) {
  if (!button) return;
  button.classList.toggle('hidden', !visible);
  const enabled = isFeatureEnabled(flagName);
  button.classList.toggle('coming-soon', visible && !enabled);
  button.dataset.badge = visible && !enabled ? 'Soon' : '';
  button.title = enabled
    ? ''
    : `${flagName} is off. This Phase 1 placeholder is not active yet.`;
  if (!preserveCurrent) {
    button.disabled = visible && !enabled;
  }
}

function applyFeatureFlagNavigation() {
  const adminVisible = isAdmin() || canManageUsers();
  markFeatureTab(els.reportsTabBtn, 'payrollReports', { visible: adminVisible });
  markFeatureTab(els.adminTabBtn, 'newDashboard', { visible: canManageUsers(), preserveCurrent: true });
  markFeatureTab(els.employeesTabBtn, 'employeeProfiles', { visible: canManageEmployees(), preserveCurrent: true });
  markFeatureTab(els.editPunchesTabBtn, 'editPunchRequests', { visible: canEditPunches(), preserveCurrent: true });
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
  els.adminTabBtn?.classList.toggle('hidden', !canManageUsers());
  els.agencyTabBtn?.classList.toggle('hidden', !canEdit);
  applyFeatureFlagNavigation();

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
  } else if (canManageUsers()) {
    switchTab('adminTab');
  } else {
    document.querySelectorAll('.tab-panel').forEach((panel) => panel.classList.add('hidden'));
    toast(`Role "${normalizedRole() || 'unknown'}" has no dashboard permissions.`, true);
  }
  renderSignedInAccessRequestCard();
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

  ensureTabDataLoaded(tabId);
}

function ensureTabDataLoaded(tabId) {
  if (tabId === 'timesheetsTab') {
    attachTimesheetView();
    return;
  }

  if (tabId === 'editPunchesTab') {
    loadSelectedWeekForManager({ force: !state.selectedWeekPunchRowsLoaded });
    return;
  }

  if (tabId === 'agencyTab') {
    loadSelectedWeekForManager({ force: !state.selectedWeekPunchRowsLoaded });
    if (!state.loadedTabs.has('agencyTab')) {
      state.loadedTabs.add('agencyTab');
      loadAgencyEmployeePage({ reset: true });
    } else {
      renderAgencyWorkbench();
    }
    return;
  }

  if (tabId === 'employeesTab' && !state.loadedTabs.has('employeesTab')) {
    state.loadedTabs.add('employeesTab');
    attachEmployeesView();
  }
}

function canAccessTab(tabId) {
  if (['myTimecardTab', 'missedPunchTab'].includes(tabId)) return isEmployee();
  if (['managerTab', 'timesheetsTab', 'editPunchesTab', 'approvalsTab', 'agencyTab'].includes(tabId)) {
    return canEditPunches();
  }
  if (tabId === 'reportsTab') return isAdmin() && isFeatureEnabled('payrollReports');
  if (tabId === 'employeesTab') return canManageEmployees();
  if (tabId === 'adminTab') return canManageUsers();
  return false;
}

function attachManagerLiveViews() {
  const todayStartMs = startOfLocalDay(new Date()).getTime();
  const constraints = [
    ...branchConstraints(),
    where('timestampMs', '>=', todayStartMs),
  ];
  if (isAgencyUser()) constraints.push(where('agencyId', '==', agencyScopeId()));
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
        const rows = snap.docs.map((d) => normalizePunchRecordForDisplay({ id: d.id, ...d.data() })).filter(isActivePunchRecord);
        state.allPunchRows = rows;
        renderLivePunches(rows);
        renderActiveNow(rows);
        renderEditPunchesTable(getEditablePunchRows());
        renderGpsSummary(rows);
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
    els.livePunchBody.innerHTML = '<tr><td colspan="5">No live data yet.</td></tr>';
    return;
  }

  els.livePunchBody.innerHTML = rows
    .map((row) => `
      <tr>
        <td>${formatDateTime(row.timestampMs)}</td>
        <td>${escapeHtml(row.name || '-')}</td>
        <td>${prettyAction(row.action)}</td>
        <td>${escapeHtml(row.source || '-')}</td>
        <td>${renderGpsBadge(row)}</td>
      </tr>
    `)
    .join('');
}

function renderGpsSummary(rows) {
  const summary = rows.reduce((counts, row) => {
    const status = String(row.locationStatus || 'not_requested').toLowerCase();
    if (status === 'verified') counts.verified += 1;
    if (status === 'denied') counts.denied += 1;
    if (row.withinAllowedRadius === false) counts.outside += 1;
    if (isGpsAccuracyTooLow(row)) counts.lowAccuracy += 1;
    return counts;
  }, { verified: 0, denied: 0, outside: 0, lowAccuracy: 0 });
  if (els.gpsVerifiedCount) els.gpsVerifiedCount.textContent = String(summary.verified);
  if (els.gpsDeniedCount) els.gpsDeniedCount.textContent = String(summary.denied);
  if (els.gpsOutsideCount) els.gpsOutsideCount.textContent = String(summary.outside);
  if (els.gpsLowAccuracyCount) els.gpsLowAccuracyCount.textContent = String(summary.lowAccuracy);
}

function renderGpsBadge(row) {
  if (row.withinAllowedRadius === false) {
    return '<span class="gps-badge warning">Outside Radius</span>';
  }
  if (isGpsAccuracyTooLow(row)) {
    return '<span class="gps-badge warning">Accuracy Too Low</span>';
  }
  const status = String(row.locationStatus || 'not_requested').toLowerCase();
  if (status === 'verified') return '<span class="gps-badge verified">GPS Verified</span>';
  if (status === 'denied') return '<span class="gps-badge denied">GPS Denied</span>';
  return `<span class="gps-badge">${escapeHtml(prettyAction(status))}</span>`;
}

function isGpsAccuracyTooLow(row) {
  const accuracy = finiteNumberOrNull(row.gpsAccuracyMeters);
  const threshold = positiveNumberOrDefault(row.maxGpsAccuracyMeters, 100);
  return accuracy !== null && accuracy > threshold;
}

function getEditablePunchRows() {
  return state.selectedWeekPunchRowsLoaded
    ? state.selectedWeekPunchRows
    : state.allPunchRows;
}

function findEditablePunch(punchId) {
  return getEditablePunchRows().find((row) => row.id === punchId)
    || state.selectedWeekPunchRows.find((row) => row.id === punchId)
    || state.allPunchRows.find((row) => row.id === punchId)
    || null;
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
    els.editPunchesBody.innerHTML = '<tr><td colspan="9">Your role cannot edit punches.</td></tr>';
    return;
  }

  const filter = String(els.editFilterNameInput?.value || '').trim().toLowerCase();
  const filtered = rows.filter((row) => {
    if (!filter) return true;
    return String(row.name || '').toLowerCase().includes(filter);
  });

  if (!filtered.length) {
    const weekText = formatDateKey(state.selectedWeekStart);
    els.editPunchesBody.innerHTML = `<tr><td colspan="9">No punches found for week ${escapeHtml(weekText)}.</td></tr>`;
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
        <td>${renderGpsBadge(row)}</td>
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

  const row = findEditablePunch(punchId);
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
      ...branchPayload(row.siteId || getCurrentSiteId())
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

  const row = findEditablePunch(punchId);
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

function getWeeklyDataCacheKey(weekStartDate = state.selectedWeekStart) {
  const normalizedWeekStart = normalizeWeekStartDate(weekStartDate);
  const weekKey = formatDateKey(normalizedWeekStart);
  return [
    getCurrentCompanyId(),
    getCurrentSiteId(),
    agencyScopeId(),
    weekKey
  ].join('|');
}

async function loadWeeklyWorkspaceData({ weekStartDate = state.selectedWeekStart, force = false, scope = 'Weekly Signoff' } = {}) {
  const normalizedWeekStart = normalizeWeekStartDate(weekStartDate);
  const weekKey = formatDateKey(normalizedWeekStart);
  const cacheKey = getWeeklyDataCacheKey(normalizedWeekStart);
  const cached = state.weeklyDataCache.get(cacheKey);
  if (!force && cached && Date.now() - cached.savedAt < WEEKLY_CACHE_TTL_MS) return cached.data;

  const pending = (async () => {
    resetFirestoreReadCounter(scope);
    const employeePromise = loadBranchEmployees({ force, scope });

    const punchConstraints = [
      ...branchConstraints(),
      where('weekKey', '==', weekKey),
    ];
    if (isAgencyUser()) punchConstraints.push(where('agencyId', '==', agencyScopeId()));
    punchConstraints.push(orderBy('timestampMs', 'asc'));

    const tsConstraints = [
      ...branchConstraints(),
      where('weekKey', '==', weekKey),
    ];
    if (isAgencyUser()) tsConstraints.push(where('agencyId', '==', agencyScopeId()));

    const [employees, punchSnap, timesheetSnap, compatibleRows] = await Promise.all([
      employeePromise,
      getDocsCounted(query(collection(db, 'punches'), ...punchConstraints), scope),
      getDocsCounted(query(collection(db, 'timesheets'), ...tsConstraints), scope),
      loadCompatibleWeeklyPunchRows(normalizedWeekStart, scope)
    ]);

    const snapshotRows = punchSnap.docs.map((d) => normalizePunchRecordForDisplay({ id: d.id, ...d.data() }));
    const mergedRows = dedupePunches([...snapshotRows, ...compatibleRows]);
    const timesheets = {};
    timesheetSnap.docs.forEach((d) => {
      timesheets[d.id] = { id: d.id, ...d.data() };
    });

    return { weekKey, employees, punches: mergedRows, timesheets };
  })();

  state.weeklyDataCache.set(cacheKey, { savedAt: Date.now(), data: pending });
  try {
    const data = await pending;
    state.weeklyDataCache.set(cacheKey, { savedAt: Date.now(), data });
    return data;
  } catch (error) {
    state.weeklyDataCache.delete(cacheKey);
    throw error;
  }
}

function applyWeeklyWorkspaceData(data, { resetAgencyRange = true } = {}) {
  state.allEmployees = data.employees || [];
  state.selectedWeekTimesheetDocs = data.timesheets || {};
  const weeklyPunches = data.punches || [];
  state.agencyReview.deletedPunchRows = weeklyPunches.filter((row) => !isActivePunchRecord(row));
  state.selectedWeekPunchRows = weeklyPunches.filter(isActivePunchRecord);
  state.selectedWeekPunchRowsLoaded = true;
  if (resetAgencyRange) state.agencyReview.rangePunchRows = null;
  renderEditPunchesTable(getEditablePunchRows());
  renderEmployeeList(state.allEmployees);
  renderManagerTimeWorkerOptions();
  if (isAdmin()) renderDuplicateWorkers(false);
  renderDerivedTimesheets();
  populateAgencyWorkerSelect();
  renderAgencyPreview();
  renderAgencyWorkbench();
}

function attachTimesheetView({ force = false } = {}) {
  state.selectedWeekStart = normalizeWeekStartDate(state.selectedWeekStart);
  syncSelectedWeekInputs();
  state.selectedWeekPunchRowsLoaded = false;
  state.selectedWeekPunchRows = [];
  renderEditPunchesTable(getEditablePunchRows());
  loadWeeklyWorkspaceData({ weekStartDate: state.selectedWeekStart, force, scope: 'Weekly Signoff' })
    .then((data) => applyWeeklyWorkspaceData(data))
    .catch((error) => {
      console.error(error);
      state.selectedWeekPunchRowsLoaded = true;
      state.selectedWeekPunchRows = [];
      renderEditPunchesTable(getEditablePunchRows());
      toast(error.message || 'Could not load weekly signoff data.', true);
    });
}

async function loadCompatibleWeeklyPunchRows(weekStartDate, scope = 'Weekly Signoff') {
  const weekStart = startOfLocalDay(normalizeWeekStartDate(weekStartDate));
  const weekEnd = addLocalDays(weekStart, 6);
  weekEnd.setHours(23, 59, 59, 999);
  const constraints = [...branchConstraints()];
  if (isAgencyUser()) constraints.push(where('agencyId', '==', agencyScopeId()));
  try {
    return await fetchPunchesWithRange(constraints, weekStart.getTime(), weekEnd.getTime(), { scope });
  } catch (error) {
    console.warn('Compatible weekly punch lookup failed:', error.message);
    return [];
  }
}

function renderDerivedTimesheets() {
  if (!els.timesheetBody) return;

  const rows = getCanonicalPayrollWorkers().map((worker) => worker.row);

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
  const profileIndex = buildWorkerProfileIndex();
  const canonicalDirectory = buildCanonicalWorkerDirectory([], profileIndex);

  state.selectedWeekPunchRows.forEach((p) => {
    const key = getWorkerIdentityKey(p, canonicalDirectory, profileIndex);
    if (!key) return;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(p);
  });

  const copiedNameKeyCounts = new Map();
  grouped.forEach((personPunches) => {
    const copiedKey = normalizeName(getCopiedWorkerName(personPunches[0]));
    if (copiedKey) copiedNameKeyCounts.set(copiedKey, (copiedNameKeyCounts.get(copiedKey) || 0) + 1);
  });

  const rows = [];

  grouped.forEach((personPunches, identityKey) => {
    const employee = findTimesheetEmployee(personPunches, identityKey, profileIndex);
    const firstPunch = personPunches[0] || {};
    const copiedPunchName = getCopiedWorkerName(firstPunch);
    const displayName = getWorkerProfileName(employee) || copiedPunchName || identityKey.replace(/^(worker|name|email|person):/, '').split('|')[0].replaceAll('_', ' ');
    const nameKey = normalizeName(displayName);
    const canonicalWorkerId = identityKey.startsWith('worker:') ? identityKey.slice('worker:'.length) : '';
    const workerIds = [...new Set(personPunches.flatMap((punch) => getRecordWorkerIds(punch)).concat(getWorkerProfileIds(employee)).filter(Boolean))];
    const stableWorkerId = canonicalWorkerId || workerIds[0] || employee?.id || employee?.employeeId || '';
    const totals = buildWeekTotals(personPunches);
    const timesheetId = `${weekKey}_${sanitizeTimesheetIdPart(stableWorkerId || nameKey)}`;
    const saved = findSavedTimesheetForGroup({
      fallbackTimesheetId: timesheetId,
      workerId: stableWorkerId,
      workerIds,
      nameKey,
      weekKey,
      allowLegacyNameFallback: (copiedNameKeyCounts.get(normalizeName(copiedPunchName)) || 0) <= 1
    });
    const copiedTimesheetName = getCopiedWorkerName(saved || {});
    const siteId = normalizeSiteId(firstPresent(
      firstPunch.siteId,
      firstPunch.assignedSiteId,
      employee?.siteId,
      employee?.assignedSiteId,
      getCurrentSiteId()
    ));
    logWorkerNameDiagnostic({
      workerId: stableWorkerId,
      profileName: displayName,
      copiedName: copiedPunchName || copiedTimesheetName,
      agencyId: firstPresent(firstPunch.agencyId, employee?.agencyId, saved?.agencyId, state.agencyId),
      agencyName: agencyLabel(firstPresent(firstPunch.agencyId, employee?.agencyId, saved?.agencyId, state.agencyId)),
      branchId: siteId,
      branchName: siteId,
      weekStart: weekKey,
      source: copiedPunchName ? 'punch' : 'timesheet'
    });
    if (copiedTimesheetName) {
      logWorkerNameDiagnostic({
        workerId: stableWorkerId,
        profileName: displayName,
        copiedName: copiedTimesheetName,
        agencyId: firstPresent(firstPunch.agencyId, employee?.agencyId, saved?.agencyId, state.agencyId),
        agencyName: agencyLabel(firstPresent(firstPunch.agencyId, employee?.agencyId, saved?.agencyId, state.agencyId)),
        branchId: siteId,
        branchName: siteId,
        weekStart: weekKey,
        source: 'timesheet'
      });
    }

    rows.push({
      id: saved?.id || timesheetId,
      identityKey,
      workerIds,
      punchCount: personPunches.length,
      name: displayName,
      workerName: displayName,
      nameKey,
      weekKey,
      companyId: firstPunch.companyId || getCurrentCompanyId(),
      agencyId: firstPresent(firstPunch.agencyId, employee?.agencyId, saved?.agencyId, state.agencyId),
      siteId,
      branchId: siteId,
      branchName: siteId,
      employeeId: employee?.employeeId || employee?.id || firstPunch.employeeId || stableWorkerId || '',
      workerId: employee?.workerId || employee?.id || firstPunch.workerId || stableWorkerId || '',
      copiedPunchName,
      copiedTimesheetName,
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

function findSavedTimesheetForGroup({ fallbackTimesheetId, workerId, workerIds = [], nameKey, weekKey, allowLegacyNameFallback = true }) {
  const docs = Object.values(state.selectedWeekTimesheetDocs || {});
  if (state.selectedWeekTimesheetDocs[fallbackTimesheetId]) return state.selectedWeekTimesheetDocs[fallbackTimesheetId];
  const workerIdSet = new Set([workerId, ...(workerIds || [])].filter(Boolean));
  if (workerIdSet.size) {
    const byId = docs.find((row) =>
      row.weekKey === weekKey
      && getRecordWorkerIds(row).some((id) => workerIdSet.has(id))
    );
    if (byId) return byId;
  }
  if (!allowLegacyNameFallback) return null;
  const legacyId = `${weekKey}_${nameKey}`;
  if (state.selectedWeekTimesheetDocs[legacyId]) return state.selectedWeekTimesheetDocs[legacyId];
  return docs.find((row) => row.weekKey === weekKey && normalizeName(row.nameKey || getCopiedWorkerName(row)) === nameKey) || null;
}

function findTimesheetEmployee(personPunches, identityKey, profileIndex = buildWorkerProfileIndex()) {
  const source = [...new Map([
    ...(state.allEmployees || []),
    ...(state.publicEmployeeRecords || []),
  ].map((employee) => [employee.id || employee.employeeId || `${employee.nameKey}:${employee.name}`, employee])).values()];
  const ids = new Set();
  personPunches.forEach((punch) => {
    getTrustedRecordWorkerIds(punch, profileIndex).forEach((id) => ids.add(id));
  });

  for (const id of ids) {
    if (profileIndex.has(id)) return profileIndex.get(id);
  }

  const byId = source.find((employee) =>
    ids.has(String(employee.id || '')) || ids.has(String(employee.employeeId || '')) || ids.has(String(employee.workerId || ''))
  );
  if (byId) return byId;

  const nameKey = identityKey.startsWith('name:')
    ? identityKey.slice(5)
    : identityKey.startsWith('person:')
      ? identityKey.slice(7).split('|')[0]
      : normalizeName(getCopiedWorkerName(personPunches[0]));
  const normalizedName = String(nameKey || '').trim().toLowerCase();
  const byName = source.filter((employee) =>
    normalizeName(employee.name || employee.nameKey || '') === normalizedName
  );
  if (!byName.length) return null;
  return byName.sort((left, right) => {
    const leftActive = isActiveEmployee(left) ? 1 : 0;
    const rightActive = isActiveEmployee(right) ? 1 : 0;
    return rightActive - leftActive || compareEmployeeRecords(left, right);
  })[0];
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
      workerName: row.name,
      employeeName: row.name,
      displayName: row.name,
      nameKey: row.nameKey,
      weekKey: row.weekKey,
      companyId: row.companyId || getCurrentCompanyId(),
      agencyId: row.agencyId || state.agencyId || '',
      siteId: row.siteId || getCurrentSiteId(),
      branchId: row.branchId || row.siteId || getCurrentSiteId(),
      branchName: row.branchName || row.siteId || getCurrentSiteId(),
      employeeId: row.employeeId || '',
      workerId: row.workerId || row.employeeId || '',
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
    loadWeeklyWorkspaceData({ force: true, scope: 'Weekly Signoff' })
      .then((data) => applyWeeklyWorkspaceData(data))
      .catch((error) => console.warn('Could not refresh weekly signoff cache:', error.message));
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
      workerName: row.name,
      employeeName: row.name,
      displayName: row.name,
      nameKey: row.nameKey,
      weekKey: row.weekKey,
      companyId: row.companyId || getCurrentCompanyId(),
      agencyId: row.agencyId || state.agencyId || '',
      siteId: row.siteId || getCurrentSiteId(),
      branchId: row.branchId || row.siteId || getCurrentSiteId(),
      branchName: row.branchName || row.siteId || getCurrentSiteId(),
      employeeId: row.employeeId || '',
      workerId: row.workerId || row.employeeId || '',
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
    loadWeeklyWorkspaceData({ force: true, scope: 'Weekly Signoff' })
      .then((data) => applyWeeklyWorkspaceData(data))
      .catch((error) => console.warn('Could not refresh weekly signoff cache:', error.message));
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not reopen timesheet.', true);
  }
}

function buildCurrentTimesheetRow(timesheetId, weekKey) {
  const rows = getCanonicalPayrollWorkers().map((worker) => worker.row);
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
    ? normalizeWeekStartDate(els.myTimecardWeekPicker.value)
    : normalizeWeekStartDate(state.selectedWeekStart);
  state.selectedWeekStart = weekStart;
  syncSelectedWeekInputs();

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
    const rows = snap.docs.map((d) => normalizePunchRecordForDisplay({ id: d.id, ...d.data() })).filter(isActivePunchRecord);
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
  if (isAgencyUser()) constraints.push(where('agencyId', '==', agencyScopeId()));

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
    const punchDate = new Date(req.requestedTimestampMs);
    const dateKey = formatDateKey(punchDate);
    const weekKey = formatDateKey(getMondayDate(punchDate));
    const punchPayload = {
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
    };
    validatePunchPayloadForSave(punchPayload);

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
    await addDoc(collection(db, 'punches'), punchPayload);

    toast('Request approved — punch created.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not approve request.', true);
  }
}

async function denyRequest(requestId) {
  if (!canEditPunches()) { toast('Your role cannot deny punch requests.', true); return; }

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
  state.loadedTabs.add('employeesTab');
  loadBranchEmployees({ scope: 'employees' })
    .then((rows) => {
      state.allEmployees = rows;
      renderEmployeeList(state.allEmployees);
      renderManagerTimeWorkerOptions();
      if (isAdmin()) renderDuplicateWorkers(false);
    })
    .catch((error) => {
      console.error(error);
      toast(error.message || 'Could not load employees.', true);
    });
}

function renderEmployeeList(employees) {
  if (!els.employeeListBody || !els.inactiveWorkerListBody) return;

  const filter = String(els.empFilterInput?.value || '').trim().toLowerCase();
  const statusFilter = state.employeeStatusFilter || 'active';
  const filtered = employees.filter((e) => {
    const employeeStatus = String(e.status || 'active').toLowerCase();
    if (statusFilter !== 'all' && employeeStatus !== statusFilter) return false;
    if (!filter) return true;
    const profileName = getWorkerProfileName(e);
    return (
      String(profileName).toLowerCase().includes(filter) ||
      String(e.employeeNumber || '').toLowerCase().includes(filter)
    );
  });
  const rosterEmployees = statusFilter === 'active'
    ? collapseDuplicateEmployees(filtered)
    : filtered;
  const inactiveEmployees = employees.filter((employee) => !isActiveEmployee(employee));
  renderDuplicateWorkerWarning(employees);

  els.inactiveWorkerListBody.innerHTML = inactiveEmployees.length
    ? inactiveEmployees.map((employee) => `
      <tr>
        <td>${escapeHtml(employee.employeeNumber || '-')}</td>
        <td>${escapeHtml(getWorkerProfileName(employee) || '-')}</td>
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
      <td>${escapeHtml(getWorkerProfileName(emp) || '-')}</td>
      <td>${escapeHtml(agencyLabel(emp.agencyId))}</td>
      <td>${escapeHtml(getWorkerBranchId(emp) || '-')}</td>
      <td><span class="tiny-flag">${escapeHtml(emp.status || 'active')}</span></td>
      <td>
        <button class="secondary-btn emp-edit-btn" data-id="${emp.id}" type="button">Edit</button>
      </td>
    </tr>
  `).join('');

  els.employeeListBody.querySelectorAll('.emp-edit-btn').forEach((btn) => {
    btn.addEventListener('click', () => loadEmployeeForEdit(btn.dataset.id));
  });
  els.duplicateWorkerWarning?.querySelectorAll('.repair-duplicate-worker-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      if (els.repairWorkerIdInput) els.repairWorkerIdInput.value = btn.dataset.workerId || '';
      if (els.repairWorkerNameInput) els.repairWorkerNameInput.value = btn.dataset.suggestedName || '';
      if (els.workerNameRepairStatus) {
        els.workerNameRepairStatus.textContent = `Prepared worker ${btn.dataset.workerId || ''}. Review the worker ID before renaming or repairing copied names.`;
      }
    });
  });
}

function getExactDuplicateWorkerGroups(workers = state.allEmployees || []) {
  const groups = new Map();
  (workers || []).filter(isActiveEmployee).forEach((worker) => {
    const nameKey = normalizeName(getWorkerProfileName(worker));
    if (!nameKey) return;
    const agencyId = String(worker.agencyId || '').trim();
    const branchId = getWorkerBranchId(worker);
    const key = [nameKey, agencyId, branchId].join('|');
    const group = groups.get(key) || {
      nameKey,
      name: getWorkerProfileName(worker),
      agencyId,
      agencyName: agencyLabel(agencyId),
      branchId,
      branchName: getWorkerBranchName(worker),
      workers: []
    };
    group.workers.push(worker);
    groups.set(key, group);
  });
  return Array.from(groups.values()).filter((group) => group.workers.length > 1);
}

function renderDuplicateWorkerWarning(workers = state.allEmployees || []) {
  if (!els.duplicateWorkerWarning) return;
  const duplicates = getExactDuplicateWorkerGroups(workers);
  if (!duplicates.length || !isAdmin()) {
    els.duplicateWorkerWarning.classList.add('hidden');
    els.duplicateWorkerWarning.innerHTML = '';
    return;
  }

  els.duplicateWorkerWarning.classList.remove('hidden');
  els.duplicateWorkerWarning.innerHTML = `
    <strong>Duplicate-looking worker profiles detected</strong>
    <p>These workers share the same normalized name, agency, and branch. Review manually; QRTimeclock will not merge or delete them automatically.</p>
    ${duplicates.map((group) => `
      <div style="margin-top:8px;">
        <strong>${escapeHtml(group.name)}</strong>
        <span class="tiny">${escapeHtml(group.agencyName || group.agencyId || 'Direct')} / ${escapeHtml(group.branchName || group.branchId || '-')}</span>
        <div class="tiny">Worker IDs: ${group.workers.map((worker) => escapeHtml(worker.id)).join(', ')}</div>
        <div class="form-actions" style="margin-top:6px;">
          ${group.workers.map((worker) => `
            <button class="ghost-btn repair-duplicate-worker-btn" type="button" data-worker-id="${escapeHtml(worker.id)}" data-suggested-name="${normalizeName(group.name) === 'bashir_ahmed' ? 'Ervin Wilson' : escapeHtml(getWorkerProfileName(worker))}">Use ${escapeHtml(worker.id)}</button>
          `).join('')}
        </div>
      </div>
    `).join('')}
  `;
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
  resetFirestoreReadCounter('History search');
  const employeeId = els.managerTimeWorkerSelect?.value || '';
  const employee = state.allEmployees.find((row) => row.id === employeeId);
  if (!employee) {
    toast('Choose a worker first.', true);
    return;
  }
  if (isAgencyUser() && employee.agencyId !== agencyScopeId()) {
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
  if (isAgencyUser() && employee.agencyId !== agencyScopeId()) return;

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
  if (els.empSiteInput) els.empSiteInput.value = emp.siteId || emp.assignedSiteId || CURRENT_SITE_ID;
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
  const requestedSiteId = String(els.empSiteInput?.value || CURRENT_SITE_ID).trim();
  const assignedSiteId = canUseSite(requestedSiteId) ? requestedSiteId : CURRENT_SITE_ID;
  if (requestedSiteId !== assignedSiteId) {
    toast('You do not have access to assign that branch.', true);
    return;
  }
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
    normalizedName: nameKey,
    employeeNumber,
    employeeNumberKey: normalizeWorkerNumber(employeeNumber),
    companyId: getCurrentCompanyId(),
    agencyId,
    assignedSiteId,
    siteId: assignedSiteId,
    siteIds: assignedSiteId ? [assignedSiteId] : [],
    qrSlug: existingEmployee?.qrSlug || state.siteContext.qrSlug || '',
    status,
    active: status === 'active',
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
      if (pin) {
        payload.pinHash = await hashWorkerPin(pin, existingId);
        payload.pinUpdatedAt = serverTimestamp();
      }
      await setDoc(doc(db, 'employees', existingId), payload, { merge: true });
      const auditAction = existingEmployee?.status !== status
        ? 'worker_status_changed'
        : 'employee_updated';
      await logAudit(auditAction, 'employee', existingId, existingEmployee || {}, payload, 'Admin employee update');
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
      if (pin) {
        payload.pinHash = await hashWorkerPin(pin, employeeId);
        payload.pinUpdatedAt = serverTimestamp();
      }
      if (!reusableEmployee) {
        payload.createdAt = serverTimestamp();
      }

      await setDoc(doc(db, 'employees', employeeId), payload, { merge: true });
      await logAudit(reusableEmployee ? 'employee_reused_existing' : 'employee_created', 'employee', employeeId, reusableEmployee || {}, payload, 'Admin employee save with stable ID');
      toast(reusableEmployee ? 'Existing active employee reused and updated.' : 'Employee created: ' + employeeNumber);
    }

    invalidateEmployeeCaches();
    cancelEmployeeEdit();
    attachEmployeesView();
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
    invalidateEmployeeCaches();
    attachEmployeesView();
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

async function exportBackup() {
  if (!isManager()) return;
  try {
    const collections = ['employees', 'punches', 'missedPunchRequests', 'agencies', 'sites', 'users'];
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

function setWorkerNameRepairStatus(message, isError = false) {
  if (!els.workerNameRepairStatus) return;
  els.workerNameRepairStatus.innerHTML = message;
  els.workerNameRepairStatus.style.borderColor = isError ? 'rgba(255,92,92,0.5)' : '';
}

function getWorkerForRepair(workerId) {
  const normalizedId = String(workerId || '').trim();
  if (!normalizedId) return null;
  return (state.allEmployees || []).find((employee) => getWorkerProfileIds(employee).includes(normalizedId)) || null;
}

function buildEmployeeNamePayload(name) {
  const cleanName = prettifyHumanName(name);
  const nameKey = normalizeName(cleanName);
  const [firstName, ...lastParts] = cleanName.split(' ');
  return {
    name: cleanName,
    workerName: cleanName,
    employeeName: cleanName,
    displayName: cleanName,
    nameKey,
    normalizedName: nameKey,
    firstName: firstName || '',
    lastName: lastParts.join(' '),
    updatedAt: serverTimestamp()
  };
}

async function renameWorkerProfileFromRepairTool() {
  if (!isAdmin()) {
    toast('Only admins can rename worker profiles.', true);
    return;
  }
  const workerId = String(els.repairWorkerIdInput?.value || '').trim();
  const newName = prettifyHumanName(els.repairWorkerNameInput?.value || '');
  if (!workerId || !newName || normalizeName(newName).length < 2) {
    setWorkerNameRepairStatus('Enter a worker ID and a valid new profile name.', true);
    return;
  }
  const employee = getWorkerForRepair(workerId);
  if (!employee) {
    setWorkerNameRepairStatus(`Worker ${escapeHtml(workerId)} was not found in the loaded roster.`, true);
    return;
  }
  const currentName = getWorkerProfileName(employee);
  const reason = String(prompt(`Rename worker profile ${employee.id} from "${currentName}" to "${newName}"? Reason:`, 'Correct stale worker profile name') || '').trim();
  if (!reason) {
    toast('A reason is required to rename a worker profile.', true);
    return;
  }

  const payload = buildEmployeeNamePayload(newName);
  try {
    await updateDoc(doc(db, 'employees', employee.id), payload);
    await logAudit('worker_profile_renamed', 'employee', employee.id, employee, {
      employeeId: employee.id,
      previousName: currentName,
      newName,
      agencyId: employee.agencyId || '',
      branchId: getWorkerBranchId(employee)
    }, reason);
    setWorkerNameRepairStatus(`Worker profile ${escapeHtml(employee.id)} renamed to ${escapeHtml(newName)}. Use Repair Copied Names to update punches/timesheets for the selected week.`);
    toast('Worker profile renamed.');
  } catch (error) {
    console.error(error);
    setWorkerNameRepairStatus(error.message || 'Could not rename worker profile.', true);
  }
}

async function getDocsForCopiedNameRepair(collectionName, employee, weekKey) {
  const seen = new Map();
  const ids = getWorkerProfileIds(employee);
  for (const fieldName of ['employeeId', 'workerId', 'userId']) {
    for (const employeeId of ids) {
      const constraints = [where(fieldName, '==', employeeId)];
      if (weekKey) constraints.push(where('weekKey', '==', weekKey));
      const snap = await getDocs(query(collection(db, collectionName), ...constraints));
      snap.docs.forEach((record) => seen.set(record.id, { ref: record.ref, data: { id: record.id, ...record.data() } }));
    }
  }
  return Array.from(seen.values());
}

async function repairCopiedNamesFromWorkerProfile() {
  if (!isAdmin()) {
    toast('Only admins can repair copied names.', true);
    return;
  }
  const workerId = String(els.repairWorkerIdInput?.value || '').trim();
  const employee = getWorkerForRepair(workerId);
  if (!employee) {
    setWorkerNameRepairStatus(`Worker ${escapeHtml(workerId)} was not found in the loaded roster.`, true);
    return;
  }
  const profileName = getWorkerProfileName(employee);
  if (!profileName) {
    setWorkerNameRepairStatus('That worker profile does not have a usable name.', true);
    return;
  }
  const weekValue = String(els.repairWeekStartInput?.value || '').trim();
  const weekKey = weekValue ? formatDateKey(getMondayDate(new Date(`${weekValue}T12:00:00`))) : '';
  const scopeText = weekKey ? `week ${weekKey}` : 'all weeks';
  const reason = String(prompt(`Repair copied punch/timesheet names for ${profileName} (${employee.id}) for ${scopeText}? Reason:`, 'Sync copied names from employee profile') || '').trim();
  if (!reason) {
    toast('A reason is required to repair copied names.', true);
    return;
  }

  const siteId = getWorkerBranchId(employee) || getCurrentSiteId();
  const copiedNamePayload = {
    name: profileName,
    workerName: profileName,
    employeeName: profileName,
    displayName: profileName,
    nameKey: normalizeName(profileName),
    employeeId: employee.id,
    workerId: employee.workerId || employee.employeeId || employee.id,
    employeeNumber: employee.employeeNumber || '',
    agencyId: employee.agencyId || '',
    siteId,
    assignedSiteId: siteId,
    branchId: siteId,
    branchName: getWorkerBranchName(employee) || siteId,
    updatedAt: serverTimestamp()
  };

  try {
    const [punchDocs, timesheetDocs] = await Promise.all([
      getDocsForCopiedNameRepair('punches', employee, weekKey),
      getDocsForCopiedNameRepair('timesheets', employee, weekKey)
    ]);
    await commitDocumentUpdates([
      ...punchDocs.map((record) => ({ ref: record.ref, data: copiedNamePayload })),
      ...timesheetDocs.map((record) => ({ ref: record.ref, data: copiedNamePayload }))
    ]);
    await logAudit('worker_copied_names_repaired', 'employee', employee.id, {
      employeeId: employee.id,
      weekKey,
      punchCount: punchDocs.length,
      timesheetCount: timesheetDocs.length
    }, {
      employeeId: employee.id,
      profileName,
      weekKey,
      punchCount: punchDocs.length,
      timesheetCount: timesheetDocs.length
    }, reason);
    setWorkerNameRepairStatus(`Repaired copied names for ${escapeHtml(profileName)}: ${punchDocs.length} punches and ${timesheetDocs.length} timesheets. No punch history, hours, or signatures were deleted.`);
    toast('Copied names repaired.');
  } catch (error) {
    console.error(error);
    setWorkerNameRepairStatus(error.message || 'Could not repair copied names.', true);
  }
}

function prepareKnownErvinWilsonRepair() {
  if (!isAdmin()) {
    toast('Only admins can use repair tools.', true);
    return;
  }
  if (els.repairWorkerNameInput) els.repairWorkerNameInput.value = 'Ervin Wilson';
  const candidates = (state.allEmployees || []).filter((employee) => {
    const nameMatch = normalizeName(getWorkerProfileName(employee)) === 'bashir_ahmed';
    const agencyKey = normalizeScopeId(employee.agencyId || 'direct');
    const agencyMatch = !agencyKey || agencyKey === 'direct';
    const branchMatch = normalizeScopeId(getWorkerBranchId(employee)) === 'oh01';
    return isActiveEmployee(employee) && nameMatch && agencyMatch && branchMatch;
  });

  if (!candidates.length) {
    setWorkerNameRepairStatus('No active Direct / OH01 Bashir Ahmed duplicate candidates are loaded. Check the roster filters/scope, then try again.', true);
    return;
  }

  setWorkerNameRepairStatus(`
    <strong>Ervin Wilson correction prepared</strong>
    <p>Choose the incorrect Bashir Ahmed worker ID, verify it against the worker details/hours, then click Rename Worker Profile. Nothing changes until you confirm.</p>
    <div class="form-actions" style="margin-top:6px;">
      ${candidates.map((employee) => `<button class="ghost-btn repair-duplicate-worker-btn" type="button" data-worker-id="${escapeHtml(employee.id)}" data-suggested-name="Ervin Wilson">Use ${escapeHtml(employee.id)}</button>`).join('')}
    </div>
  `);
  els.workerNameRepairStatus?.querySelectorAll('.repair-duplicate-worker-btn').forEach((btn) => {
    btn.addEventListener('click', () => {
      if (els.repairWorkerIdInput) els.repairWorkerIdInput.value = btn.dataset.workerId || '';
      if (els.repairWorkerNameInput) els.repairWorkerNameInput.value = 'Ervin Wilson';
    });
  });
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
    els.pendingUserListBody.innerHTML = '<tr><td colspan="9">No pending users.</td></tr>';
    return;
  }

  els.pendingUserListBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.name || row.displayName || '-')}</td>
      <td>${escapeHtml(row.email || '-')}</td>
      <td>${escapeHtml(row.requestedRole || '-')}</td>
      <td>${escapeHtml(row.companyId || '-')}</td>
      <td>${escapeHtml(getAllowedSiteIds(row).join(', ') || '-')}</td>
      <td>${escapeHtml(row.agencyName || '-')}</td>
      <td>${escapeHtml(row.requestedAdminEmail || '-')}</td>
      <td>${escapeHtml(formatTimestamp(row.accessRequestedAt) || '-')}</td>
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

  const role = normalizeRequestedAccessRole(row.requestedRole || 'manager');
  if (!['manager', 'supervisor', 'agency_admin'].includes(role)) {
    toast('Invalid requested role.', true);
    return;
  }

  const siteIds = parseSiteIds(row.siteIds || row.siteId || CURRENT_SITE_ID);
  const payload = {
    active: true,
    approvalStatus: 'approved',
    role,
    companyId: getCurrentCompanyId(),
    branch: siteIds[0],
    branches: siteIds,
    siteId: siteIds[0],
    siteIds,
    permissions: {
      canEditPunches: true,
      canDeletePunches: false,
      canMergeWorkers: false,
      manageUsers: false,
    },
    approvedBy: state.me.uid,
    approvedByEmail: state.me.email || state.profile?.email || '',
    approvedAt: serverTimestamp(),
    updatedAt: serverTimestamp(),
  };
  if (role === 'agency_admin') {
    payload.agencyName = row.agencyName || '';
    payload.agencyId = row.agencyId || normalizeScopeId(row.agencyName || '');
  }

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
    deniedByEmail: state.me.email || state.profile?.email || '',
    deniedAt: serverTimestamp(),
    updatedAt: serverTimestamp(),
  };
  await setDoc(doc(db, 'users', uid), payload, { merge: true });
  await logAudit('user_denied', 'user', uid, row || {}, payload, 'Pending user denied');
  toast('User denied/deactivated.');
}

function normalizeCleanupSiteId(value) {
  const cleaned = String(value || '').trim().toUpperCase();
  if (cleaned === 'HC') return 'OHC';
  if (cleaned === 'OH01' || cleaned === 'OHC') return cleaned;
  return '';
}

function normalizeCleanupSiteIds(value) {
  const raw = Array.isArray(value) ? value : String(value || '').split(',');
  return [...new Set(raw.map(normalizeCleanupSiteId).filter(Boolean))];
}

async function previewUserBranchCleanup() {
  if (!isOwnerOrSuperAdmin()) {
    toast('Only an owner can fix user branch data.', true);
    return;
  }
  if (!els.userBranchCleanupPreview) return;

  try {
    els.userBranchCleanupPreview.textContent = 'Scanning users...';
    const snap = await getDocs(query(collection(db, 'users'), where('companyId', '==', getCurrentCompanyId())));
    const fixes = [];
    snap.docs.forEach((record) => {
      const row = { id: record.id, ...record.data() };
      const normalizedSiteIds = normalizeCleanupSiteIds(row.siteIds);
      const fallbackSiteId = normalizeCleanupSiteId(row.siteId);
      const nextSiteIds = normalizedSiteIds.length ? normalizedSiteIds : (fallbackSiteId ? [fallbackSiteId] : []);
      const siteIdsChanged = JSON.stringify(Array.isArray(row.siteIds) ? row.siteIds : row.siteIds || '') !== JSON.stringify(nextSiteIds);
      const siteIdChanged = nextSiteIds[0] && row.siteId !== nextSiteIds[0];
      if (nextSiteIds.length && (siteIdsChanged || siteIdChanged)) {
        fixes.push({
          id: row.id,
          email: row.email || '',
          before: row.siteIds,
          beforeSiteId: row.siteId || '',
          siteIds: nextSiteIds,
          siteId: nextSiteIds[0],
        });
      }
    });

    renderUserBranchCleanupPreview(fixes);
  } catch (error) {
    console.error(error);
    toast(formatFirestoreError(error), true);
    els.userBranchCleanupPreview.textContent = 'Could not scan users.';
  }
}

function renderUserBranchCleanupPreview(fixes) {
  if (!els.userBranchCleanupPreview) return;
  if (!fixes.length) {
    els.userBranchCleanupPreview.innerHTML = '<div class="empty-state">No user branch cleanup needed.</div>';
    return;
  }

  els.userBranchCleanupPreview.innerHTML = `
    <div class="status-box">${fixes.length} user profile${fixes.length === 1 ? '' : 's'} need siteIds cleanup. Review before applying.</div>
    ${fixes.map((fix) => `
      <div class="person-row">
        <div class="person-meta">
          <strong>${escapeHtml(fix.email || fix.id)}</strong>
          <span>siteIds: ${escapeHtml(JSON.stringify(fix.before))} -> ${escapeHtml(JSON.stringify(fix.siteIds))}</span>
          <span>siteId: ${escapeHtml(fix.beforeSiteId || '-')} -> ${escapeHtml(fix.siteId)}</span>
        </div>
      </div>
    `).join('')}
    <button id="applyUserBranchCleanupBtn" class="danger-btn" type="button">Apply these user-only fixes</button>
  `;

  document.getElementById('applyUserBranchCleanupBtn')?.addEventListener('click', () => applyUserBranchCleanup(fixes));
}

async function applyUserBranchCleanup(fixes) {
  if (!isOwnerOrSuperAdmin()) {
    toast('Only an owner can apply user branch cleanup.', true);
    return;
  }
  if (!fixes.length) return;
  const okay = confirm(`Apply user branch cleanup to ${fixes.length} user profile(s)? This only updates users.siteIds/siteId.`);
  if (!okay) return;

  const batch = writeBatch(db);
  fixes.forEach((fix) => {
    batch.update(doc(db, 'users', fix.id), {
      siteId: fix.siteId,
      siteIds: fix.siteIds,
      updatedAt: serverTimestamp(),
    });
  });
  await batch.commit();
  toast('User branch data cleanup applied.');
  els.userBranchCleanupPreview.innerHTML = '<div class="empty-state">Cleanup applied. Run preview again to verify.</div>';
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
        const rows = snap.docs.map((d) => normalizeUserProfile({ id: d.id, uid: d.id, userId: d.id, ...d.data() }));
        state.allUsers = rows;

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
    const defaultPermissions = defaultPermissionsForRole(role);
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
      branch: requestedSiteIds[0] || CURRENT_SITE_ID,
      branches: requestedSiteIds,
      siteId: requestedSiteIds[0] || CURRENT_SITE_ID,
      siteIds: requestedSiteIds,
      approvalStatus: els.userActiveInput?.value === 'true' ? 'approved' : (existingProfile.approvalStatus || 'pending'),
      permissions: {
        canEditPunches: defaultPermissions.canEditPunches || existingProfile.permissions?.canEditPunches === true,
        canDeletePunches: defaultPermissions.canDeletePunches || existingProfile.permissions?.canDeletePunches === true,
        canMergeWorkers: defaultPermissions.canMergeWorkers || existingProfile.permissions?.canMergeWorkers === true,
        manageUsers: defaultPermissions.manageUsers || existingProfile.permissions?.manageUsers === true
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

const AGENCY_REVIEW_COLUMNS = [
  ['name', 'Employee'],
  ['employeeNumber', 'Emp #'],
  ['agency', 'Agency'],
  ['branch', 'Branch'],
  ['department', 'Department'],
  ['weeklyHours', 'Hours'],
  ['punchCount', 'Punches'],
  ['overtimeHours', 'OT'],
  ['status', 'Status'],
  ['lastPunch', 'Last punch'],
  ['warnings', 'Warnings'],
];
const AGENCY_ROSTER_PAGE_SIZE = 75;

async function loadAgencyEmployeePage({ reset = false } = {}) {
  if (!canEditPunches() || state.agencyEmployeesLoading) return;
  if (!reset && state.agencyEmployeesExhausted && state.agencyWorkersExhausted) return;
  state.agencyEmployeesLoading = true;
  if (els.agencyLoadMoreEmployeesBtn) els.agencyLoadMoreEmployeesBtn.disabled = true;
  if (els.agencyRosterPageStatus) els.agencyRosterPageStatus.textContent = reset ? 'Loading roster page...' : 'Loading more employees and workers...';

  try {
    if (reset) {
      state.agencyEmployeeRows = [];
      state.agencyWorkerProfileRows = [];
      state.agencyEmployeeCursor = { siteId: null, assignedSiteId: null };
      state.agencyWorkerCursor = { siteId: null, assignedSiteId: null };
      state.agencyEmployeesExhausted = false;
      state.agencyWorkersExhausted = false;
    }
    if (!state.agencyEmployeeCursor || !('siteId' in state.agencyEmployeeCursor)) {
      state.agencyEmployeeCursor = { siteId: null, assignedSiteId: null };
    }
    if (!state.agencyWorkerCursor || !('siteId' in state.agencyWorkerCursor)) {
      state.agencyWorkerCursor = { siteId: null, assignedSiteId: null };
    }

    const pageSize = Math.ceil(AGENCY_ROSTER_PAGE_SIZE / 4);
    const queries = [
      ['employees', 'siteId', state.agencyEmployeeCursor.siteId],
      ['employees', 'assignedSiteId', state.agencyEmployeeCursor.assignedSiteId],
      ['workers', 'siteId', state.agencyWorkerCursor.siteId],
      ['workers', 'assignedSiteId', state.agencyWorkerCursor.assignedSiteId],
    ];
    const results = await Promise.allSettled(
      queries.map(([collectionName, siteField, cursor]) => loadAgencyRosterPagePart(collectionName, siteField, cursor, pageSize))
    );
    const successful = results.filter((result) => result.status === 'fulfilled');
    if (!successful.length) throw results[0]?.reason || new Error('Roster page query failed.');

    const [employeeSite, employeeAssigned, workerSite, workerAssigned] = results.map((result, index) => {
      const [, , cursor] = queries[index];
      return result.status === 'fulfilled' ? result.value : { rows: [], cursor, exhausted: true };
    });

    const employeeRows = [...employeeSite.rows, ...employeeAssigned.rows];
    const existingEmployees = new Map(state.agencyEmployeeRows.map((row) => [agencyRosterRecordKey(row), row]));
    employeeRows.forEach((row) => existingEmployees.set(agencyRosterRecordKey(row), row));
    state.agencyEmployeeRows = [...existingEmployees.values()];
    state.agencyEmployeeCursor = { siteId: employeeSite.cursor, assignedSiteId: employeeAssigned.cursor };
    state.agencyEmployeesExhausted = employeeSite.exhausted && employeeAssigned.exhausted;

    const workerRows = [...workerSite.rows, ...workerAssigned.rows];
    const existingWorkers = new Map(state.agencyWorkerProfileRows.map((row) => [agencyRosterRecordKey(row), row]));
    workerRows.forEach((row) => existingWorkers.set(agencyRosterRecordKey(row), row));
    state.agencyWorkerProfileRows = [...existingWorkers.values()];
    state.agencyWorkerCursor = { siteId: workerSite.cursor, assignedSiteId: workerAssigned.cursor };
    state.agencyWorkersExhausted = workerSite.exhausted && workerAssigned.exhausted;

    if (els.agencyRosterPageStatus) {
      const partial = results.some((result) => result.status === 'rejected') ? ' Partial roster page loaded.' : '';
      const loadedCount = state.agencyEmployeeRows.length + state.agencyWorkerProfileRows.length;
      const exhausted = state.agencyEmployeesExhausted && state.agencyWorkersExhausted;
      els.agencyRosterPageStatus.textContent = exhausted
        ? `Loaded ${loadedCount} employee/worker record(s). End of roster.${partial}`
        : `Loaded ${loadedCount} employee/worker record(s). More available.${partial}`;
    }
    renderAgencyWorkbench();
  } catch (error) {
    console.error(error);
    if (els.agencyRosterPageStatus) els.agencyRosterPageStatus.textContent = 'Could not load roster page.';
    toast(error.message || 'Could not load Agency Export roster page.', true);
  } finally {
    state.agencyEmployeesLoading = false;
    if (els.agencyLoadMoreEmployeesBtn) {
      const exhausted = state.agencyEmployeesExhausted && state.agencyWorkersExhausted;
      els.agencyLoadMoreEmployeesBtn.disabled = exhausted;
      els.agencyLoadMoreEmployeesBtn.textContent = exhausted ? 'Roster Loaded' : 'Load More Roster';
    }
  }
}

async function loadAgencyRosterPagePart(collectionName, siteField, cursor, pageSize) {
  const constraints = [
    where('companyId', '==', getCurrentCompanyId()),
    where(siteField, '==', getCurrentSiteId()),
  ];
  if (isAgencyUser()) constraints.push(where('agencyId', '==', agencyScopeId()));
  constraints.push(orderBy('name', 'asc'));
  if (cursor) constraints.push(startAfter(cursor));
  constraints.push(limit(pageSize));

  const snap = await getDocsCounted(query(collection(db, collectionName), ...constraints), 'Agency Export');
  return {
    rows: snap.docs.map((record) => ({ id: record.id, sourceCollection: collectionName, ...record.data() })),
    cursor: snap.docs.at(-1) || cursor || null,
    exhausted: snap.docs.length < pageSize,
  };
}

function agencyRosterRecordKey(row) {
  return [
    row?.sourceCollection || 'employees',
    row?.id || row?.employeeId || row?.workerId || row?.uid || normalizeName(row?.name || row?.displayName || row?.nameKey || '')
  ].join(':');
}
function renderAgencyWorkbench() {
  if (!els.agencyReviewBody) return;
  ensureAgencyColumnChooser();
  populateAgencyFilterOptions();

  const allRows = buildAgencyReviewRows();
  logAgencyPunchCoverage(allRows, getAgencySourcePunches());
  const rows = sortAgencyRows(filterAgencyRows(allRows));
  state.agencyReview.filteredRows = rows;
  renderAgencyStats(rows);

  const pageSize = Number(state.agencyReview.pageSize || els.agencyPageSizeSelect?.value || 25);
  state.agencyReview.pageSize = pageSize;
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize));
  state.agencyReview.page = Math.min(Math.max(1, state.agencyReview.page || 1), totalPages);
  const start = (state.agencyReview.page - 1) * pageSize;
  const pageRows = rows.slice(start, start + pageSize);

  if (!pageRows.length) {
    els.agencyReviewBody.innerHTML = '<tr><td colspan="11">No employees match the current filters.</td></tr>';
  } else {
    els.agencyReviewBody.innerHTML = pageRows.map((row) => `
      <tr data-worker-key="${escapeHtml(row.identityKey)}" class="${row.identityKey === state.agencyReview.selectedKey ? 'selected-row' : ''}" tabindex="0">
        <td data-col="name"><strong>${escapeHtml(row.name || '-')}</strong><br><span class="tiny">${escapeHtml(row.pin ? `PIN ${row.pin}` : row.identityKey)}</span></td>
        <td data-col="employeeNumber">${escapeHtml(row.employeeNumber || '-')}</td>
        <td data-col="agency">${escapeHtml(row.agency || '-')}</td>
        <td data-col="branch">${escapeHtml(row.branch || '-')}</td>
        <td data-col="department">${escapeHtml(row.department || '-')}</td>
        <td data-col="weeklyHours">${Number(row.weeklyHours || 0).toFixed(2)}</td>
        <td data-col="punchCount">${Number(row.punchCount || 0)}</td>
        <td data-col="overtimeHours">${Number(row.overtimeHours || 0).toFixed(2)}</td>
        <td data-col="status"><span class="pill">${escapeHtml(row.statusLabel)}</span></td>
        <td data-col="lastPunch">${escapeHtml(row.lastPunchText || '-')}</td>
        <td data-col="warnings">${renderAgencyWarningBadges(row.warnings)}</td>
      </tr>
    `).join('');
  }

  els.agencyReviewBody.querySelectorAll('tr[data-worker-key]').forEach((tr) => {
    tr.addEventListener('click', () => openAgencyEmployeePanel(tr.dataset.workerKey));
    tr.addEventListener('dblclick', () => openAgencyQuickEdit(tr.dataset.workerKey));
    tr.addEventListener('keydown', (event) => {
      if (event.key === 'Enter') openAgencyEmployeePanel(tr.dataset.workerKey);
    });
  });

  if (els.agencyPageInfo) els.agencyPageInfo.textContent = `Page ${state.agencyReview.page} of ${totalPages} (${rows.length} employees)`;
  if (els.agencyPrevPageBtn) els.agencyPrevPageBtn.disabled = state.agencyReview.page <= 1;
  if (els.agencyNextPageBtn) els.agencyNextPageBtn.disabled = state.agencyReview.page >= totalPages;
  if (els.agencyLoadMoreEmployeesBtn) {
    els.agencyLoadMoreEmployeesBtn.disabled = state.agencyEmployeesLoading || state.agencyEmployeesExhausted;
    els.agencyLoadMoreEmployeesBtn.textContent = state.agencyEmployeesExhausted ? 'Roster Loaded' : 'Load More Roster';
  }
  if (els.agencyRosterPageStatus && !state.agencyEmployeesLoading && !state.agencyEmployeeRows.length) {
    els.agencyRosterPageStatus.textContent = 'Roster loads in safe pages. Punches load on demand.';
  }
  renderAgencyRecoveryStatus();
  applyAgencyColumnVisibility();
  if (state.agencyReview.selectedKey) renderAgencyPanel();
}

function renderAgencyRecoveryStatus() {
  const deletedRows = state.agencyReview.deletedPunchRows || [];
  if (els.agencyRestoreDeletedBtn) {
    els.agencyRestoreDeletedBtn.disabled = !deletedRows.length || !canEditPunches();
  }
  if (els.agencyRecoveryStatus) {
    els.agencyRecoveryStatus.textContent = deletedRows.length
      ? `${deletedRows.length} soft-deleted punch(es) found in this range. Review before restoring.`
      : 'No soft-deleted punches loaded for this range.';
  }
}

function buildAgencyReviewRows() {
  const punches = getAgencySourcePunches();
  const employeeRows = getAgencyEmployeeRowsForReview();
  const profileIndex = buildWorkerProfileIndex(employeeRows, state.allUsers || []);
  const directory = buildAgencyWorkerDirectory([...punches, ...employeeRows, ...(state.allUsers || [])], profileIndex);
  const grouped = new Map();

  punches.forEach((punch) => {
    const normalized = normalizePunchRecordForDisplay(punch);
    const key = getWorkerIdentityKey(normalized, directory, profileIndex);
    if (!key) return;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(normalized);
  });

  const rows = new Map();
  grouped.forEach((personPunches, identityKey) => {
    const employee = findAgencyReviewEmployee(personPunches, identityKey, profileIndex, employeeRows);
    rows.set(identityKey, buildAgencyReviewRow(identityKey, employee, personPunches));
  });

  employeeRows.forEach((employee) => {
    const rowAgency = getRecordAgencyIdentity(employee);
    const allowedAgency = normalizeIdentityToken(agencyScopeId());
    if (isAgencyUser() && rowAgency && rowAgency !== allowedAgency) return;
    const workerBranch = getWorkerBranchId(employee) || getCurrentSiteId();
    if (!canUseSite(workerBranch)) return;
    const key = getWorkerIdentityKey(employee, directory) || `employee:${employee.id || employee.employeeId || employee.workerId || normalizeName(employee.name || employee.displayName)}`;
    if (!key || rows.has(key)) return;
    rows.set(key, buildAgencyReviewRow(key, employee, []));
  });

  return [...rows.values()];
}

function getAgencyEmployeeRowsForReview() {
  const rows = [
    ...(state.allEmployees || []),
    ...(state.agencyEmployeeRows || []),
    ...(state.agencyWorkerProfileRows || []),
    ...(state.publicEmployeeRecords || []),
  ];
  const unique = new Map();
  rows.forEach((row) => {
    const key = agencyRosterRecordKey(row);
    if (key && !unique.has(key)) unique.set(key, row);
  });
  return [...unique.values()];
}

function findAgencyReviewEmployee(personPunches, identityKey, profileIndex, employeeRows) {
  const ids = new Set();
  personPunches.forEach((punch) => {
    getTrustedRecordWorkerIds(punch, profileIndex).forEach((id) => ids.add(id));
  });
  for (const id of ids) {
    if (profileIndex.has(id)) return profileIndex.get(id);
  }
  const normalizedName = normalizeName(getCopiedWorkerName(personPunches[0]) || identityKey.replace(/^(worker|email|person):/, ''));
  return employeeRows.find((employee) =>
    getWorkerProfileIds(employee).some((id) => ids.has(id))
    || normalizeName(getWorkerProfileName(employee) || employee.nameKey || '') === normalizedName
  ) || null;
}

function buildAgencyWorkerDirectory(sourceRows = [], profileIndex = buildWorkerProfileIndex()) {
  const signatureIds = new Map();
  const emailIds = new Map();

  sourceRows.forEach((row) => {
    const ids = getDirectoryWorkerIds(row, profileIndex);
    const email = getRecordEmail(row);
    const signature = getWorkerSignature(row);
    if (signature && ids.length) {
      if (!signatureIds.has(signature)) signatureIds.set(signature, new Set());
      ids.forEach((id) => signatureIds.get(signature).add(id));
    }
    if (email && ids.length) {
      if (!emailIds.has(email)) emailIds.set(email, new Set());
      ids.forEach((id) => emailIds.get(email).add(id));
    }
  });

  const primaryFor = (ids) => [...ids].sort((left, right) => String(left).localeCompare(String(right)))[0];
  const signaturePrimary = new Map();
  const emailPrimary = new Map();
  signatureIds.forEach((ids, signature) => signaturePrimary.set(signature, primaryFor(ids)));
  emailIds.forEach((ids, email) => emailPrimary.set(email, primaryFor(ids)));
  return { signaturePrimary, emailPrimary };
}

function buildAgencyReviewRow(identityKey, employee, punches) {
  const sortedPunches = dedupePunches(punches).sort((a, b) => Number(a.timestampMs || 0) - Number(b.timestampMs || 0));
  const firstPunch = sortedPunches[0] || {};
  const displayName = getWorkerProfileName(employee) || getCopiedWorkerName(firstPunch) || identityKey.replace(/^(worker|email|person|employee):/, '');
  const totals = buildAgencyPunchTotals(sortedPunches);
  const saved = findSavedTimesheetForGroup({
    fallbackTimesheetId: `${formatDateKey(state.selectedWeekStart)}_${sanitizeTimesheetIdPart(employee?.id || employee?.employeeId || normalizeName(displayName))}`,
    workerId: employee?.id || employee?.employeeId || employee?.workerId || firstPunch.employeeId || firstPunch.workerId || '',
    workerIds: [...new Set([...(getWorkerProfileIds(employee) || []), ...sortedPunches.flatMap(getRecordWorkerIds)].filter(Boolean))],
    nameKey: normalizeName(displayName),
    weekKey: formatDateKey(state.selectedWeekStart),
    allowLegacyNameFallback: true
  });
  const removed = !isActiveEmployeeRecord(employee || {}) && !!employee;
  const missing = totals.warnings.some((warning) => /^Missing/.test(warning));
  const statusLabel = removed
    ? 'Removed'
    : saved?.status === 'signed'
      ? 'Signed'
      : missing
        ? 'Missing'
        : sortedPunches.length
          ? 'Complete'
          : 'Open';

  return {
    identityKey,
    employee,
    punches: sortedPunches,
    name: prettifyHumanName(displayName),
    employeeNumber: firstPresent(employee?.employeeNumber, employee?.workerNumber, employee?.employeeNo, firstPunch.employeeNumber, firstPunch.workerNumber),
    agencyId: firstPresent(employee?.agencyId, firstPunch.agencyId, state.agencyId),
    agency: agencyLabel(firstPresent(employee?.agencyId, firstPunch.agencyId, state.agencyId)),
    branch: firstPresent(employee?.branchName, employee?.branchId, employee?.branchCode, employee?.assignedSiteId, employee?.siteId, firstPunch.branchName, firstPunch.branchId, firstPunch.branchCode, firstPunch.assignedSiteId, firstPunch.siteId, getCurrentSiteId()),
    department: firstPresent(employee?.department, employee?.dept, firstPunch.department, firstPunch.dept),
    pin: firstPresent(employee?.pin, employee?.workerPin, firstPunch.pin),
    notes: [employee?.notes, employee?.managerNotes, firstPunch.notes, firstPunch.managerNote].flat().filter(Boolean).join(' '),
    weeklyHours: totals.workedHours,
    punchCount: sortedPunches.length,
    regularHours: Math.min(40, totals.workedHours),
    overtimeHours: Math.max(0, totals.workedHours - 40),
    lunchMinutes: totals.lunchMinutes,
    warnings: totals.warnings,
    missingClockOut: totals.warnings.includes('Missing Clock Out'),
    missingLunch: totals.warnings.includes('Missing Lunch'),
    lateArrivals: totals.lateArrivals,
    daysWorked: totals.daysWorked,
    statusLabel,
    rawStatus: String(saved?.status || employee?.status || '').toLowerCase(),
    removed,
    lastPunch: sortedPunches.at(-1)?.timestampMs || 0,
    lastPunchText: sortedPunches.length ? `${prettyAction(sortedPunches.at(-1).action)} ${formatDateTime(sortedPunches.at(-1).timestampMs)}` : '',
    daySummaries: totals.daySummaries,
    searchable: '',
  };
}

function buildAgencyPunchTotals(punches) {
  const byDay = new Map();
  punches.forEach((punch) => {
    const dateKey = punch.dateKey || formatDateKey(new Date(Number(punch.timestampMs || 0)));
    if (!byDay.has(dateKey)) byDay.set(dateKey, []);
    byDay.get(dateKey).push(punch);
  });

  let workedMinutes = 0;
  let lunchMinutes = 0;
  let lateArrivals = 0;
  const warnings = new Set();
  const daySummaries = [];

  byDay.forEach((dayPunches, dateKey) => {
    const sorted = [...dayPunches].sort((a, b) => Number(a.timestampMs || 0) - Number(b.timestampMs || 0));
    let activeStart = null;
    let lunchStart = null;
    let dayWorked = 0;
    let dayLunch = 0;
    const actionTimes = { clock_in: [], start_lunch: [], end_lunch: [], clock_out: [] };

    sorted.forEach((punch) => {
      if (actionTimes[punch.action]) actionTimes[punch.action].push(punch.timestampMs);
      if (punch.action === 'clock_in') activeStart = punch.timestampMs;
      if (punch.action === 'start_lunch') {
        if (activeStart) {
          const diff = Math.max(0, Math.round((punch.timestampMs - activeStart) / 60000));
          dayWorked += diff;
          workedMinutes += diff;
        }
        activeStart = null;
        lunchStart = punch.timestampMs;
      }
      if (punch.action === 'end_lunch') {
        if (lunchStart) {
          const diff = Math.max(0, Math.round((punch.timestampMs - lunchStart) / 60000));
          dayLunch += diff;
          lunchMinutes += diff;
        }
        lunchStart = null;
        activeStart = punch.timestampMs;
      }
      if (punch.action === 'clock_out') {
        if (activeStart) {
          const diff = Math.max(0, Math.round((punch.timestampMs - activeStart) / 60000));
          dayWorked += diff;
          workedMinutes += diff;
        }
        activeStart = null;
      }
    });

    const dayWarnings = [];
    if (activeStart) dayWarnings.push('Missing Clock Out');
    if (lunchStart) dayWarnings.push('Missing Lunch In');
    if (actionTimes.clock_in.length && !actionTimes.clock_out.length) dayWarnings.push('Missing Clock Out');
    if (dayWorked >= 360 && (!actionTimes.start_lunch.length || !actionTimes.end_lunch.length)) dayWarnings.push('Missing Lunch');
    const firstClockIn = actionTimes.clock_in[0];
    if (firstClockIn) {
      const d = new Date(firstClockIn);
      if (d.getHours() > 8 || (d.getHours() === 8 && d.getMinutes() > 5)) lateArrivals += 1;
    }
    dayWarnings.forEach((warning) => warnings.add(warning));
    daySummaries.push({
      dateKey,
      punches: sorted,
      actionTimes,
      workedHours: Number((dayWorked / 60).toFixed(2)),
      lunchMinutes: dayLunch,
      overtimeHours: Number((Math.max(0, dayWorked - 480) / 60).toFixed(2)),
      warnings: [...new Set(dayWarnings)],
    });
  });

  if (workedMinutes > 2400) warnings.add('Overtime');
  return {
    workedHours: Number((workedMinutes / 60).toFixed(2)),
    lunchMinutes,
    daysWorked: byDay.size,
    lateArrivals,
    warnings: [...warnings],
    daySummaries: daySummaries.sort((a, b) => a.dateKey.localeCompare(b.dateKey)),
  };
}

function getAgencySourcePunches() {
  return Array.isArray(state.agencyReview.rangePunchRows)
    ? state.agencyReview.rangePunchRows
    : state.selectedWeekPunchRows;
}

function logAgencyPunchCoverage(rows, punches) {
  const sourcePunches = dedupePunches((punches || []).map(normalizePunchRecordForDisplay)).filter(isActivePunchRecord);
  const sourceKeys = new Set(sourcePunches.map(agencyPunchCoverageKey));
  const representedKeys = new Set(
    rows.flatMap((row) => row.punches || []).map(agencyPunchCoverageKey)
  );
  const missing = sourcePunches.filter((punch) => !representedKeys.has(agencyPunchCoverageKey(punch)));
  const logKey = [
    formatDateKey(state.selectedWeekStart),
    sourceKeys.size,
    representedKeys.size,
    missing.map((punch) => punch.id || agencyPunchCoverageKey(punch)).join(',')
  ].join('|');
  if (state.agencyReview.coverageLogKey === logKey) return;
  state.agencyReview.coverageLogKey = logKey;

  if (missing.length) {
    if (els.agencyCoverageStatus) {
      els.agencyCoverageStatus.textContent = `${missing.length} loaded punch(es) are not represented in the Agency Export rows. Do not use this export for payroll until reviewed.`;
      els.agencyCoverageStatus.style.borderColor = 'rgba(255,92,92,0.5)';
    }
    console.error('[QRTimeclock Agency Export coverage warning]', {
      message: 'Some loaded punches are not represented in Agency Export rows. Do not use this export for payroll until this is resolved.',
      loadedPunches: sourceKeys.size,
      representedPunches: representedKeys.size,
      missingPunches: missing.map((punch) => ({
        id: punch.id || '',
        name: punch.name || punch.workerName || '',
        employeeId: punch.employeeId || '',
        workerId: punch.workerId || '',
        agencyId: punch.agencyId || '',
        siteId: punch.siteId || '',
        action: punch.action || '',
        timestamp: formatDateTime(punch.timestampMs)
      }))
    });
  } else {
    if (els.agencyCoverageStatus) {
      els.agencyCoverageStatus.textContent = `Coverage OK: ${sourceKeys.size} loaded punch(es) represented across ${rows.filter((row) => Number(row.punchCount || 0) > 0).length} employee row(s).`;
      els.agencyCoverageStatus.style.borderColor = 'rgba(43,213,118,0.4)';
    }
    console.info('[QRTimeclock Agency Export coverage OK]', {
      loadedPunches: sourceKeys.size,
      representedPunches: representedKeys.size,
      employeesWithPunches: rows.filter((row) => Number(row.punchCount || 0) > 0).length
    });
  }
}

function agencyPunchCoverageKey(punch) {
  return punch.id || [
    punch.employeeId || punch.workerId || punch.nameKey || normalizeName(punch.name || punch.workerName || ''),
    punch.timestampMs || 0,
    punch.action || '',
    punch.siteId || '',
    punch.agencyId || ''
  ].join('|');
}

function filterAgencyRows(rows) {
  const filters = getAgencyFilters();
  return rows.filter((row) => {
    if (row.removed && !filters.removed && filters.status !== 'removed') return false;
    if (filters.employee && row.identityKey !== filters.employee) return false;
    if (filters.agency && normalizeIdentityToken(row.agencyId || row.agency) !== filters.agency) return false;
    if (filters.branch && normalizeIdentityToken(row.branch) !== filters.branch) return false;
    if (filters.department && normalizeIdentityToken(row.department) !== filters.department) return false;
    if (filters.status && !agencyStatusMatches(row, filters.status)) return false;
    if (filters.missingClockOut && !row.missingClockOut) return false;
    if (filters.missingLunch && !row.missingLunch) return false;
    if (filters.overtime && Number(row.overtimeHours || 0) <= 0) return false;
    if (filters.removed && !row.removed) return false;
    if (filters.date && !row.daySummaries.some((day) => day.dateKey === filters.date)) return false;
    if (filters.search && !agencySearchText(row).includes(filters.search)) return false;
    return true;
  });
}

function getAgencyFilters() {
  return {
    search: normalizeIdentityText(els.agencySearchInput?.value || ''),
    date: els.agencyDateFilter?.value || '',
    from: els.agencyFromDateFilter?.value || '',
    to: els.agencyToDateFilter?.value || '',
    employee: els.agencyWorkerSelect?.value || '',
    agency: normalizeIdentityToken(els.agencyAgencyFilter?.value || ''),
    branch: normalizeIdentityToken(els.agencyBranchFilter?.value || ''),
    department: normalizeIdentityToken(els.agencyDepartmentFilter?.value || ''),
    status: els.agencyStatusFilter?.value || '',
    missingClockOut: els.agencyMissingClockOutFilter?.checked === true,
    missingLunch: els.agencyMissingLunchFilter?.checked === true,
    overtime: els.agencyOvertimeFilter?.checked === true,
    removed: els.agencyRemovedFilter?.checked === true,
  };
}

function agencyStatusMatches(row, status) {
  if (status === 'complete') return row.statusLabel === 'Complete';
  if (status === 'missing') return row.warnings.length > 0;
  if (status === 'open') return row.statusLabel === 'Open' || row.rawStatus === 'open';
  if (status === 'signed') return row.statusLabel === 'Signed' || row.rawStatus === 'signed';
  if (status === 'removed') return row.removed;
  return true;
}

function agencySearchText(row) {
  return normalizeIdentityText([
    row.name,
    row.employeeNumber,
    row.agency,
    row.agencyId,
    row.branch,
    row.department,
    row.pin,
    row.notes,
    row.employee?.siteId,
    row.employee?.assignedSiteId,
    row.employee?.email,
    ...row.punches.flatMap((punch) => [punch.notes, punch.managerNote, punch.source, punch.siteId, punch.department])
  ].filter(Boolean).join(' '));
}

function sortAgencyRows(rows) {
  const { sortBy, sortDir } = state.agencyReview;
  const direction = sortDir === 'desc' ? -1 : 1;
  return [...rows].sort((left, right) => {
    const a = left[sortBy] ?? '';
    const b = right[sortBy] ?? '';
    if (typeof a === 'number' || typeof b === 'number') return (Number(a || 0) - Number(b || 0)) * direction;
    return String(a).localeCompare(String(b)) * direction;
  });
}

function sortAgencyReview(sortBy) {
  if (!sortBy) return;
  if (state.agencyReview.sortBy === sortBy) {
    state.agencyReview.sortDir = state.agencyReview.sortDir === 'asc' ? 'desc' : 'asc';
  } else {
    state.agencyReview.sortBy = sortBy;
    state.agencyReview.sortDir = 'asc';
  }
  renderAgencyWorkbench();
}

function populateAgencyFilterOptions() {
  const rows = buildAgencyReviewRows();
  setSelectOptions(els.agencyWorkerSelect, rows.map((row) => [row.identityKey, `${row.name || '-'}${row.employeeNumber ? ` | ${row.employeeNumber}` : ''}`]), 'All employees');
  setSelectOptions(els.agencyAgencyFilter, rows.map((row) => [normalizeIdentityToken(row.agencyId || row.agency), row.agency || row.agencyId || 'Direct']), 'All agencies');
  setSelectOptions(els.agencyBranchFilter, rows.map((row) => [normalizeIdentityToken(row.branch), row.branch || '-']), 'All branches');
  setSelectOptions(els.agencyDepartmentFilter, rows.map((row) => [normalizeIdentityToken(row.department), row.department || '-']).filter(([value]) => value), 'All departments');
}

function setSelectOptions(select, entries, allLabel) {
  if (!select) return;
  const current = select.value;
  const unique = new Map();
  entries.forEach(([value, label]) => {
    if (value && !unique.has(value)) unique.set(value, label);
  });
  select.innerHTML = `<option value="">${escapeHtml(allLabel)}</option>` +
    [...unique.entries()].sort((a, b) => String(a[1]).localeCompare(String(b[1])))
      .map(([value, label]) => `<option value="${escapeHtml(value)}">${escapeHtml(label)}</option>`).join('');
  if ([...unique.keys()].includes(current)) select.value = current;
}

function renderAgencyStats(rows) {
  const totalHours = rows.reduce((sum, row) => sum + Number(row.weeklyHours || 0), 0);
  const activeRows = rows.filter((row) => !row.removed);
  const attendedRows = activeRows.filter((row) => row.punches.length);
  if (els.agencyStatsEmployees) els.agencyStatsEmployees.textContent = String(rows.length);
  if (els.agencyStatsWeekHours) els.agencyStatsWeekHours.textContent = totalHours.toFixed(2);
  if (els.agencyStatsMissedClockOuts) els.agencyStatsMissedClockOuts.textContent = String(rows.filter((row) => row.missingClockOut).length);
  if (els.agencyStatsMissedLunches) els.agencyStatsMissedLunches.textContent = String(rows.filter((row) => row.missingLunch).length);
  if (els.agencyStatsLateArrivals) els.agencyStatsLateArrivals.textContent = String(rows.reduce((sum, row) => sum + Number(row.lateArrivals || 0), 0));
  if (els.agencyStatsAttendance) els.agencyStatsAttendance.textContent = activeRows.length ? `${Math.round((attendedRows.length / activeRows.length) * 100)}%` : '0%';
  if (els.agencyStatsMonthHours) els.agencyStatsMonthHours.textContent = totalHours.toFixed(2);
  if (els.agencyStatsPayPeriodHours) els.agencyStatsPayPeriodHours.textContent = totalHours.toFixed(2);
}

function renderAgencyWarningBadges(warnings = []) {
  if (!warnings.length) return '<span class="gps-badge verified">OK</span>';
  return warnings.map((warning) => `<span class="gps-badge warning">${escapeHtml(warning)}</span>`).join(' ');
}

function ensureAgencyColumnChooser() {
  if (!els.agencyColumnChooser) return;
  if (!state.agencyReview.visibleColumns) {
    state.agencyReview.visibleColumns = Object.fromEntries(AGENCY_REVIEW_COLUMNS.map(([key]) => [key, true]));
  }
  if (els.agencyColumnChooser.dataset.ready === 'true') return;
  els.agencyColumnChooser.innerHTML = AGENCY_REVIEW_COLUMNS.map(([key, label]) => `
    <label><input type="checkbox" data-agency-column="${escapeHtml(key)}" checked /> ${escapeHtml(label)}</label>
  `).join('');
  els.agencyColumnChooser.querySelectorAll('[data-agency-column]').forEach((input) => {
    input.addEventListener('change', () => {
      state.agencyReview.visibleColumns[input.dataset.agencyColumn] = input.checked;
      applyAgencyColumnVisibility();
    });
  });
  els.agencyColumnChooser.dataset.ready = 'true';
}

function applyAgencyColumnVisibility() {
  const visible = state.agencyReview.visibleColumns || {};
  AGENCY_REVIEW_COLUMNS.forEach(([key]) => {
    const hide = visible[key] === false;
    els.agencyReviewTable?.querySelectorAll(`[data-col="${CSS.escape(key)}"]`).forEach((cell) => {
      cell.classList.toggle('column-hidden', hide);
    });
  });
}

async function handleAgencyDateRangeChange() {
  const filters = getAgencyFilters();
  if (!filters.from && !filters.to && !filters.date) {
    state.agencyReview.rangePunchRows = null;
    renderAgencyWorkbench();
    return;
  }
  const fromValue = filters.date || filters.from;
  const toValue = filters.date || filters.to || filters.from;
  if (!fromValue || !toValue) {
    renderAgencyWorkbench();
    return;
  }
  const fromMs = new Date(`${fromValue}T00:00:00`).getTime();
  const toMs = new Date(`${toValue}T23:59:59`).getTime();
  if (!Number.isFinite(fromMs) || !Number.isFinite(toMs) || fromMs > toMs) {
    toast('Choose a valid Agency Export date range.', true);
    return;
  }
  try {
    resetFirestoreReadCounter('Agency Export');
    const rows = await loadAgencyRangePunches(fromValue, toValue, fromMs, toMs);
    state.agencyReview.deletedPunchRows = rows.filter((row) => !isActivePunchRecord(row));
    state.agencyReview.rangePunchRows = rows.filter(isActivePunchRecord);
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not load Agency Export date range.', true);
  }
  renderAgencyWorkbench();
}

async function loadAgencyRangePunches(fromValue, toValue, fromMs, toMs) {
  const selectedWeekStart = formatDateInput(state.selectedWeekStart);
  const selectedWeekEnd = formatDateInput(addLocalDays(state.selectedWeekStart, 6));
  if (fromValue === selectedWeekStart && toValue === selectedWeekEnd) {
    const weekly = await loadWeeklyWorkspaceData({ scope: 'Agency Export' });
    return weekly.punches || [];
  }

  const constraints = [...branchConstraints()];
  if (isAgencyUser()) constraints.push(where('agencyId', '==', agencyScopeId()));
  return fetchPunchesWithRange(constraints, fromMs, toMs, { scope: 'Agency Export' });
}

function openAgencyEmployeePanel(identityKey) {
  state.agencyReview.selectedKey = identityKey;
  state.agencyReview.activeTab = state.agencyReview.activeTab || 'overview';
  els.agencyEmployeePanel?.classList.remove('hidden');
  renderAgencyWorkbench();
}

function closeAgencyPanel() {
  state.agencyReview.selectedKey = '';
  els.agencyEmployeePanel?.classList.add('hidden');
  renderAgencyWorkbench();
}

function switchAgencyPanelTab(tabName) {
  state.agencyReview.activeTab = tabName || 'overview';
  document.querySelectorAll('[data-agency-panel-tab]').forEach((button) => {
    button.classList.toggle('active', button.dataset.agencyPanelTab === state.agencyReview.activeTab);
  });
  renderAgencyPanel();
}

function getSelectedAgencyRow() {
  const rows = state.agencyReview.filteredRows.length ? state.agencyReview.filteredRows : buildAgencyReviewRows();
  return rows.find((row) => row.identityKey === state.agencyReview.selectedKey) || null;
}

function renderAgencyPanel() {
  const row = getSelectedAgencyRow();
  if (!row || !els.agencyPanelContent) return;
  if (els.agencyPanelTitle) els.agencyPanelTitle.textContent = row.name || 'Employee';
  if (els.agencyPanelMeta) {
    els.agencyPanelMeta.textContent = `${row.employeeNumber || 'No employee #'} | ${row.agency || 'Direct'} | ${row.branch || '-'}`;
  }
  const tab = state.agencyReview.activeTab || 'overview';
  if (tab === 'timecard') els.agencyPanelContent.innerHTML = renderAgencyTimecard(row);
  else if (tab === 'punches') els.agencyPanelContent.innerHTML = renderAgencyTimeline(row);
  else if (tab === 'approvals') els.agencyPanelContent.innerHTML = renderAgencyApprovals(row);
  else if (tab === 'notes') els.agencyPanelContent.innerHTML = renderAgencyNotes(row);
  else if (tab === 'audit') renderAgencyAudit(row);
  else els.agencyPanelContent.innerHTML = renderAgencyOverview(row);
  wireAgencyPanelActions(row);
}

function renderAgencyOverview(row) {
  return `
    <div class="stats-grid">
      <div class="stat-card"><span>Total Hours</span><strong>${Number(row.weeklyHours || 0).toFixed(2)}</strong></div>
      <div class="stat-card"><span>Lunch Length</span><strong>${formatMinutes(row.lunchMinutes)}</strong></div>
      <div class="stat-card"><span>Worked Hours</span><strong>${Number(row.regularHours || 0).toFixed(2)}</strong></div>
      <div class="stat-card"><span>Overtime</span><strong>${Number(row.overtimeHours || 0).toFixed(2)}</strong></div>
    </div>
    <div class="status-box">${renderAgencyWarningBadges(row.warnings)}</div>
    ${renderAgencyTimeline(row)}
  `;
}

function renderAgencyTimecard(row) {
  const days = row.daySummaries.length ? row.daySummaries : buildEmptyWeekDays();
  return `
    <div class="mini-table-wrap">
      <table>
        <thead><tr><th>Date</th><th>Clock In</th><th>Lunch Out</th><th>Lunch In</th><th>Clock Out</th><th>Regular Hours</th><th>OT Hours</th><th>Status</th><th>Manager Notes</th></tr></thead>
        <tbody>
          ${days.map((day) => `
            <tr>
              <td>${escapeHtml(day.dateKey)}</td>
              <td>${formatAgencyActionTimes(day, 'clock_in')}</td>
              <td>${formatAgencyActionTimes(day, 'start_lunch')}</td>
              <td>${formatAgencyActionTimes(day, 'end_lunch')}</td>
              <td>${formatAgencyActionTimes(day, 'clock_out')}</td>
              <td>${Number(Math.min(day.workedHours || 0, 8)).toFixed(2)}</td>
              <td>${Number(day.overtimeHours || 0).toFixed(2)}</td>
              <td>${renderAgencyWarningBadges(day.warnings || [])}</td>
              <td>${escapeHtml(row.notes || '-')}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    </div>
  `;
}

function renderAgencyTimeline(row) {
  const groups = row.daySummaries.length ? row.daySummaries : [];
  return `
    <div class="form-actions">
      <button class="primary-btn agency-add-punch-btn" type="button">Add Missing Punch</button>
    </div>
    <div class="agency-timeline">
      ${groups.length ? groups.map((day) => `
        <section class="agency-day">
          <h4>${escapeHtml(formatLongDate(day.dateKey))}</h4>
          ${day.punches.map((punch) => `
            <div class="agency-punch-row">
              <div><strong>${formatTime(punch.timestampMs)}</strong><span>${prettyAction(punch.action)}${punch.inserted ? ' | Inserted' : ''}</span></div>
              <div class="form-actions">
                <button class="ghost-btn agency-note-punch-btn" data-punch-id="${escapeHtml(punch.id)}" type="button">Add Note</button>
                <button class="secondary-btn agency-edit-punch-btn" data-punch-id="${escapeHtml(punch.id)}" type="button">Edit</button>
                <button class="danger-btn agency-delete-punch-btn" data-punch-id="${escapeHtml(punch.id)}" type="button" ${canDeletePunches() ? '' : 'disabled'}>Delete</button>
              </div>
            </div>
          `).join('')}
          <div class="tiny">Worked ${Number(day.workedHours || 0).toFixed(2)} hrs | Lunch ${formatMinutes(day.lunchMinutes || 0)} | ${renderAgencyWarningBadges(day.warnings || [])}</div>
        </section>
      `).join('') : '<div class="empty-state">No punches for this employee in the selected range.</div>'}
    </div>
  `;
}

function renderAgencyApprovals(row) {
  const requests = (state.allMissedRequests || []).filter((request) => {
    const ids = new Set([row.employee?.id, row.employee?.employeeId, row.employee?.workerId, row.identityKey].filter(Boolean));
    return ids.has(request.employeeId) || normalizeName(request.name || request.employeeName || '') === normalizeName(row.name);
  });
  return requests.length
    ? requests.map((request) => `<div class="person-row"><div class="person-meta"><strong>${escapeHtml(prettyAction(request.requestedAction))}</strong><span>${escapeHtml(request.reason || '-')}</span></div><span class="pill">${escapeHtml(request.status || 'pending')}</span></div>`).join('')
    : '<div class="empty-state">No missed punch approvals found for this employee.</div>';
}

function renderAgencyNotes(row) {
  return `
    <div class="status-box">${escapeHtml(row.notes || 'No manager notes yet.')}</div>
    <div class="form-actions"><button class="primary-btn agency-add-employee-note-btn" type="button">Add Employee Note</button></div>
  `;
}

async function renderAgencyAudit(row) {
  els.agencyPanelContent.innerHTML = '<div class="empty-state">Loading audit history...</div>';
  const punchIds = row.punches.map((punch) => punch.id).filter(Boolean);
  const auditRows = [];
  try {
    for (const punchId of punchIds) {
      const snap = await getDocs(query(collection(db, 'punch_edits'), where('punchId', '==', punchId)));
      snap.docs.forEach((record) => auditRows.push({ id: record.id, ...record.data() }));
    }
    auditRows.sort((a, b) => Number(b.editedAt?.seconds || 0) - Number(a.editedAt?.seconds || 0));
    els.agencyPanelContent.innerHTML = auditRows.length
      ? auditRows.map((audit) => `<div class="person-row"><div class="person-meta"><strong>${escapeHtml(audit.type || audit.action || 'edit')}</strong><span>${escapeHtml(audit.reason || audit.editReason || '-')}</span></div><span class="tiny">${escapeHtml(audit.editedBy || audit.managerName || '-')}</span></div>`).join('')
      : '<div class="empty-state">No audit history found for this selected range.</div>';
  } catch (error) {
    console.error(error);
    els.agencyPanelContent.innerHTML = '<div class="empty-state">Could not load audit history.</div>';
  }
}

function wireAgencyPanelActions(row) {
  els.agencyPanelContent?.querySelectorAll('.agency-edit-punch-btn').forEach((button) => {
    button.addEventListener('click', () => editAgencyPunch(button.dataset.punchId));
  });
  els.agencyPanelContent?.querySelectorAll('.agency-delete-punch-btn').forEach((button) => {
    button.addEventListener('click', () => deleteAgencyPunch(button.dataset.punchId));
  });
  els.agencyPanelContent?.querySelectorAll('.agency-note-punch-btn').forEach((button) => {
    button.addEventListener('click', () => addAgencyPunchNote(button.dataset.punchId));
  });
  els.agencyPanelContent?.querySelector('.agency-add-punch-btn')?.addEventListener('click', () => addMissingAgencyPunch(row));
  els.agencyPanelContent?.querySelector('.agency-add-employee-note-btn')?.addEventListener('click', () => addAgencyEmployeeNote(row));
}

async function editAgencyPunch(punchId) {
  if (!canEditPunches()) return toast('You need edit-punch permission to edit punches.', true);
  const punch = findAgencyPunch(punchId);
  if (!punch) return toast('Punch not found.', true);
  const reason = prompt('Reason required for this punch edit:');
  if (!String(reason || '').trim()) return toast('A reason is required.', true);
  const newDateTime = prompt('Correct time (YYYY-MM-DD HH:MM):', toLocalEditString(punch.timestampMs));
  if (newDateTime === null) return;
  const parsedMs = parseLocalEditString(newDateTime);
  if (!parsedMs) return toast('Invalid date/time format.', true);
  const newAction = normalizePunchAction(prompt('Correct punch type (clock_in, start_lunch, end_lunch, clock_out):', punch.action) || '');
  if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(newAction)) return toast('Invalid punch type.', true);
  const date = new Date(parsedMs);
  const payload = {
    action: newAction,
    type: newAction,
    timestampMs: parsedMs,
    dateKey: formatDateKey(date),
    weekKey: formatDateKey(getMondayDate(date)),
    editedAt: serverTimestamp(),
    editedBy: state.profile?.name || state.me?.email || 'Manager',
    editedByUid: state.me?.uid || '',
    editReason: String(reason).trim(),
    updatedAt: serverTimestamp(),
  };
  await writeAgencyPunchChange(punch, payload, 'edit', reason);
}

async function addAgencyPunchNote(punchId) {
  if (!canEditPunches()) return toast('You need edit-punch permission to add notes.', true);
  const punch = findAgencyPunch(punchId);
  if (!punch) return toast('Punch not found.', true);
  const note = prompt('Manager note to add:');
  if (!String(note || '').trim()) return;
  const notes = Array.isArray(punch.managerNotes) ? punch.managerNotes : [];
  const payload = {
    managerNote: String(note).trim(),
    managerNotes: [...notes, { note: String(note).trim(), by: state.profile?.name || state.me?.email || 'Manager', atMs: Date.now() }],
    editedAt: serverTimestamp(),
    editedBy: state.profile?.name || state.me?.email || 'Manager',
    updatedAt: serverTimestamp(),
  };
  await writeAgencyPunchChange(punch, payload, 'note', note);
}

async function deleteAgencyPunch(punchId) {
  if (!canDeletePunches()) return toast('You need delete-punch permission to delete punches.', true);
  const punch = findAgencyPunch(punchId);
  if (!punch) return toast('Punch not found.', true);
  const reason = prompt('Reason required for deleting this punch:');
  if (!String(reason || '').trim()) return toast('A reason is required.', true);
  if (!confirm('This will mark the punch deleted without removing history. Continue?')) return;
  const payload = {
    status: 'deleted',
    active: false,
    deletedAt: serverTimestamp(),
    deletedBy: state.profile?.name || state.me?.email || 'Manager',
    deletedByUid: state.me?.uid || '',
    deleteReason: String(reason).trim(),
    updatedAt: serverTimestamp(),
  };
  await writeAgencyPunchChange(punch, payload, 'delete', reason);
}

async function restoreAgencySoftDeletedPunches() {
  const deletedRows = dedupePunches(state.agencyReview.deletedPunchRows || []);
  if (!deletedRows.length) return toast('No soft-deleted punches are loaded for this range.', true);
  if (!canEditPunches()) return toast('You need edit-punch permission to restore punches.', true);
  const reason = prompt(`Restore ${deletedRows.length} soft-deleted punch(es) for this Agency Export range? Reason required:`, 'Recover soft-deleted punch for payroll review');
  if (!String(reason || '').trim()) return toast('A reason is required to restore punches.', true);
  if (!confirm(`Restore ${deletedRows.length} soft-deleted punch(es)? This does not delete or recreate data.`)) return;

  let restored = 0;
  for (const punch of deletedRows) {
    const payload = {
      active: true,
      status: 'restored',
      restoredAt: serverTimestamp(),
      restoredBy: state.profile?.name || state.me?.email || 'Manager',
      restoredByUid: state.me?.uid || '',
      restoreReason: String(reason).trim(),
      updatedAt: serverTimestamp(),
    };
    try {
      await addDoc(collection(db, 'punch_edits'), {
        ...branchPayload(punch.siteId || getCurrentSiteId()),
        punchId: punch.id,
        type: 'restore',
        original: punch,
        updated: payload,
        reason: String(reason).trim(),
        editedBy: state.profile?.name || state.me?.email || 'Manager',
        managerName: state.profile?.name || state.me?.email || 'Manager',
        editedAt: serverTimestamp(),
      });
      await updateDoc(doc(db, 'punches', punch.id), payload);
      await logAudit('punch_restored', 'punch', punch.id, punch, payload, String(reason).trim());
      restored += 1;
    } catch (error) {
      console.error('Could not restore punch', punch.id, error);
    }
  }

  state.agencyReview.deletedPunchRows = [];
  toast(`Restored ${restored} soft-deleted punch(es).`);
  await refreshAgencyPunchRangeAfterRecovery();
}

async function refreshAgencyPunchRangeAfterRecovery() {
  const filters = getAgencyFilters();
  if (filters.date || filters.from || filters.to) {
    await handleAgencyDateRangeChange();
    return;
  }
  const rows = await loadCompatibleWeeklyPunchRows(state.selectedWeekStart);
  state.agencyReview.deletedPunchRows = rows.filter((row) => !isActivePunchRecord(row));
  state.selectedWeekPunchRows = rows.filter(isActivePunchRecord);
  renderDerivedTimesheets();
  populateAgencyWorkerSelect();
  renderAgencyPreview();
  renderAgencyWorkbench();
}

async function addMissingAgencyPunch(row) {
  if (!canEditPunches()) return toast('You need edit-punch permission to insert punches.', true);
  const action = normalizePunchAction(prompt('Missing punch type (clock_in, start_lunch, end_lunch, clock_out):', 'clock_in') || '');
  if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(action)) return toast('Invalid punch type.', true);
  const dateTime = prompt('Missing punch time (YYYY-MM-DD HH:MM):', toLocalEditString(Date.now()));
  if (dateTime === null) return;
  const timestampMs = parseLocalEditString(dateTime);
  if (!timestampMs) return toast('Invalid date/time format.', true);
  const reason = prompt('Reason required for inserted punch:', `Forgot ${prettyAction(action)}`);
  if (!String(reason || '').trim()) return toast('A reason is required.', true);
  const date = new Date(timestampMs);
  const employee = row.employee || {};
  const payload = {
    ...branchPayload(employee.assignedSiteId || employee.siteId || getCurrentSiteId()),
    companyId: getCurrentCompanyId(),
    agencyId: row.agencyId || employee.agencyId || state.agencyId || '',
    employeeId: employee.employeeId || employee.id || row.employeeId || '',
    workerId: employee.workerId || employee.id || row.workerId || '',
    employeeNumber: row.employeeNumber || employee.employeeNumber || '',
    name: row.name,
    workerName: row.name,
    employeeName: row.name,
    nameKey: normalizeName(row.name),
    action,
    type: action,
    timestampMs,
    dateKey: formatDateKey(date),
    weekKey: formatDateKey(getMondayDate(date)),
    source: 'manager_inserted',
    inserted: true,
    insertedReason: String(reason).trim(),
    managerName: state.profile?.name || state.me?.email || 'Manager',
    createdBy: state.me?.uid || '',
    createdAt: serverTimestamp(),
    updatedAt: serverTimestamp(),
  };
  try {
    validatePunchPayloadForSave(payload);
    const ref = await addDoc(collection(db, 'punches'), payload);
    await addDoc(collection(db, 'punch_edits'), {
      ...branchPayload(payload.siteId),
      punchId: ref.id,
      type: 'insert',
      original: null,
      updated: payload,
      reason: String(reason).trim(),
      editedBy: state.profile?.name || state.me?.email || 'Manager',
      managerName: state.profile?.name || state.me?.email || 'Manager',
      editedAt: serverTimestamp(),
    });
    await logAudit('punch_inserted', 'punch', ref.id, {}, payload, String(reason).trim());
    toast('Missing punch inserted.');
    if (Array.isArray(state.agencyReview.rangePunchRows)) {
      state.agencyReview.rangePunchRows.push(normalizePunchRecordForDisplay({ id: ref.id, ...payload, timestampMs }));
    }
    renderAgencyWorkbench();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not insert punch.', true);
  }
}

async function addAgencyEmployeeNote(row) {
  if (!canEditPunches()) return toast('You need edit-punch permission to add employee notes.', true);
  const employee = row.employee;
  if (!employee?.id) return toast('No employee profile found for this note.', true);
  const note = prompt('Employee manager note:');
  if (!String(note || '').trim()) return;
  const notes = Array.isArray(employee.managerNotes) ? employee.managerNotes : [];
  const payload = {
    managerNotes: [...notes, { note: String(note).trim(), by: state.profile?.name || state.me?.email || 'Manager', atMs: Date.now() }],
    notes: [employee.notes, String(note).trim()].filter(Boolean).join('\n'),
    updatedAt: serverTimestamp(),
  };
  try {
    await updateDoc(doc(db, 'employees', employee.id), payload);
    await logAudit('employee_note_added', 'employee', employee.id, employee, payload, String(note).trim());
    toast('Employee note added.');
    Object.assign(employee, payload);
    renderAgencyWorkbench();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not add employee note.', true);
  }
}

async function writeAgencyPunchChange(punch, payload, type, reason) {
  try {
    await addDoc(collection(db, 'punch_edits'), {
      ...branchPayload(punch.siteId || getCurrentSiteId()),
      punchId: punch.id,
      type,
      original: punch,
      updated: payload,
      reason: String(reason || '').trim(),
      editedBy: state.profile?.name || state.me?.email || 'Manager',
      managerName: state.profile?.name || state.me?.email || 'Manager',
      editedAt: serverTimestamp(),
    });
    await updateDoc(doc(db, 'punches', punch.id), payload);
    await logAudit(`punch_${type === 'delete' ? 'deleted' : type === 'note' ? 'note_added' : 'edited'}`, 'punch', punch.id, punch, payload, String(reason || '').trim());
    replaceAgencyRangePunch(punch.id, { ...punch, ...payload });
    toast(type === 'delete' ? 'Punch marked deleted.' : 'Punch updated.');
    renderAgencyWorkbench();
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not update punch.', true);
  }
}

function replaceAgencyRangePunch(punchId, updated) {
  if (!Array.isArray(state.agencyReview.rangePunchRows)) return;
  state.agencyReview.rangePunchRows = state.agencyReview.rangePunchRows
    .map((punch) => punch.id === punchId ? normalizePunchRecordForDisplay(updated) : punch)
    .filter(isActivePunchRecord);
}

function findAgencyPunch(punchId) {
  return getAgencySourcePunches().find((punch) => punch.id === punchId)
    || state.selectedWeekPunchRows.find((punch) => punch.id === punchId)
    || null;
}

function openAgencyQuickEdit(identityKey) {
  openAgencyEmployeePanel(identityKey);
  state.agencyReview.activeTab = 'punches';
  switchAgencyPanelTab('punches');
}

function exportAgencyReviewCsv() {
  const rows = agencyExportRows();
  downloadCsv(`agency-export-${formatDateKey(new Date())}.csv`, rows);
}

function exportAgencyReviewExcel() {
  const rows = agencyExportRows();
  const html = `<table>${rows.map((row) => `<tr>${row.map((value) => `<td>${escapeHtml(value)}</td>`).join('')}</tr>`).join('')}</table>`;
  downloadBlob(`agency-export-${formatDateKey(new Date())}.xls`, html, 'application/vnd.ms-excel;charset=utf-8');
}

function exportAgencyReviewPdf() {
  const rows = agencyExportRows();
  const win = window.open('', '_blank', 'width=1100,height=800');
  if (!win) return toast('Pop-up blocked. Allow pop-ups to export PDF.', true);
  win.document.write(`<!doctype html><html><head><title>Agency Export</title><style>body{font-family:Arial,sans-serif;margin:24px;color:#111}table{border-collapse:collapse;width:100%}td,th{border:1px solid #bbb;padding:7px;text-align:left}th{background:#f1f4f8}</style></head><body><h1>Agency Export</h1><table>${rows.map((row, index) => `<tr>${row.map((value) => index ? `<td>${escapeHtml(value)}</td>` : `<th>${escapeHtml(value)}</th>`).join('')}</tr>`).join('')}</table><script>window.print();</script></body></html>`);
  win.document.close();
}

function agencyExportRows() {
  const headers = ['Employee', 'Employee Number', 'Agency', 'Branch', 'Department', 'Punches', 'Regular Hours', 'OT Hours', 'Status', 'Warnings', 'Last Punch'];
  return [headers, ...state.agencyReview.filteredRows.map((row) => [
    row.name,
    row.employeeNumber,
    row.agency,
    row.branch,
    row.department,
    Number(row.punchCount || 0),
    Number(row.regularHours || 0).toFixed(2),
    Number(row.overtimeHours || 0).toFixed(2),
    row.statusLabel,
    row.warnings.join(' | '),
    row.lastPunchText,
  ])];
}

function downloadBlob(filename, content, type) {
  const url = URL.createObjectURL(new Blob([`\uFEFF${content}`], { type }));
  const anchor = document.createElement('a');
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

function handleAgencyKeyboardShortcuts(event) {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'f' && !document.querySelector('#agencyTab.hidden')) {
    event.preventDefault();
    els.agencySearchInput?.focus();
  }
  if (event.key === 'Escape' && !els.agencyEmployeePanel?.classList.contains('hidden')) {
    closeAgencyPanel();
  }
}

function buildEmptyWeekDays() {
  return Array.from({ length: 7 }, (_, index) => {
    const date = new Date(state.selectedWeekStart);
    date.setDate(date.getDate() + index);
    return { dateKey: formatDateKey(date), actionTimes: {}, workedHours: 0, overtimeHours: 0, lunchMinutes: 0, warnings: [] };
  });
}

function formatAgencyActionTimes(day, action) {
  const values = day.actionTimes?.[action] || [];
  return values.length ? values.map(formatTime).join('<br>') : '-';
}

function formatMinutes(minutes) {
  const total = Number(minutes || 0);
  const hours = Math.floor(total / 60);
  const mins = total % 60;
  return hours ? `${hours}h ${mins}m` : `${mins}m`;
}

function formatLongDate(dateKey) {
  const date = new Date(`${dateKey}T12:00:00`);
  return date.toLocaleDateString(undefined, { weekday: 'long', month: 'long', day: 'numeric' });
}

function getAgencySelectableWorkers() {
  const merged = new Map();
  const addWorker = (worker) => {
    if (!worker?.identityKey) return;
    const existing = merged.get(worker.identityKey);
    if (!existing) {
      merged.set(worker.identityKey, worker);
      return;
    }
    existing.row = {
      ...worker.row,
      ...existing.row,
      punchCount: Number(existing.row?.punchCount || 0) + Number(worker.row?.punchCount || 0),
      weeklyHours: Math.max(Number(existing.row?.weeklyHours || 0), Number(worker.row?.weeklyHours || 0)),
      dailyTotals: existing.row?.dailyTotals || worker.row?.dailyTotals || {},
    };
  };

  getCanonicalPayrollWorkers().forEach(addWorker);
  buildAgencyReviewRows().forEach((row) => {
    addWorker({
      identityKey: row.identityKey,
      name: row.name || row.workerName || '-',
      agencyId: row.agencyId || '',
      branchId: row.branch || row.branchId || row.siteId || '',
      row: agencyReviewRowToLegacySheetRow(row),
    });
  });

  return [...merged.values()].sort((left, right) => String(left.name || '').localeCompare(String(right.name || '')));
}

function agencyReviewRowToLegacySheetRow(row) {
  const dailyTotals = {};
  (row.daySummaries || []).forEach((day) => {
    dailyTotals[day.dateKey] = {
      clock_in: formatAgencyActionTimesText(day, 'clock_in'),
      start_lunch: formatAgencyActionTimesText(day, 'start_lunch'),
      end_lunch: formatAgencyActionTimesText(day, 'end_lunch'),
      clock_out: formatAgencyActionTimesText(day, 'clock_out'),
      hours: Number(day.workedHours || 0),
    };
  });
  return {
    id: row.identityKey,
    identityKey: row.identityKey,
    name: row.name || '-',
    workerName: row.name || '-',
    agencyId: row.agencyId || '',
    siteId: row.branch || getCurrentSiteId(),
    branchId: row.branch || getCurrentSiteId(),
    weekKey: formatDateKey(state.selectedWeekStart),
    status: row.statusLabel || 'open',
    weeklyHours: Number(row.weeklyHours || 0),
    daysWorked: Number(row.daysWorked || 0),
    dailyTotals,
    punchCount: Number(row.punches?.length || 0),
    managerSignedBy: '',
    managerSignedAt: null,
  };
}

function formatAgencyActionTimesText(day, action) {
  const values = day.actionTimes?.[action] || [];
  return values.length ? values.map(formatTime).join(' | ') : '';
}

function populateAgencyWorkerSelect() {
  if (!els.agencyWorkerSelect && !els.agencyLegacyWorkerSelect) return;

  const current = els.agencyWorkerSelect?.value || '';
  const legacyCurrent = els.agencyLegacyWorkerSelect?.value || '';
  const workers = getAgencySelectableWorkers();
  const options = workers.map((worker) => {
    const row = worker.row;
    const details = [agencyLabel(row.agencyId), row.siteId].filter(Boolean).join(' · ');
    return `<option value="${escapeHtml(worker.identityKey)}">${escapeHtml(worker.name)}${details ? ` (${escapeHtml(details)})` : ''}</option>`;
  }).join('');

  if (els.agencyWorkerSelect) {
    els.agencyWorkerSelect.innerHTML = '<option value="">All employees</option>' + options;
    if (workers.some((worker) => worker.identityKey === current)) els.agencyWorkerSelect.value = current;
  }
  if (els.agencyLegacyWorkerSelect) {
    els.agencyLegacyWorkerSelect.innerHTML = '<option value="">Select a worker</option>' + options;
    if (workers.some((worker) => worker.identityKey === legacyCurrent)) els.agencyLegacyWorkerSelect.value = legacyCurrent;
  }

  runPayrollCanonicalConsoleChecks(workers);
}

function renderAgencyPreview() {
  if (!els.agencyPreview || !els.agencyLegacyWorkerSelect) return;

  const selectedWorkerKey = els.agencyLegacyWorkerSelect.value;
  if (!selectedWorkerKey) {
    els.agencyPreview.innerHTML = '<div class="empty-state">Choose a worker and click Preview Sheet.</div>';
    return;
  }

  const weekKey = formatDateKey(state.selectedWeekStart);
  const worker = getAgencySelectableWorkers().find((item) => item.identityKey === selectedWorkerKey);
  const row = worker?.row && worker.row.weekKey === weekKey ? worker.row : null;

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
            <div><strong>Agency:</strong> ${escapeHtml(agencyLabel(row.agencyId))}</div>
            <div><strong>Branch:</strong> ${escapeHtml(row.siteId || '-')}</div>
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

function mergeDuplicatePayrollRowsInMemory(rows) {
  const merged = new Map();
  rows.forEach((row) => {
    const identityKey = row.identityKey || row.workerId || row.employeeId || row.nameKey || normalizeName(row.name || '');
    if (!identityKey) return;
    const existing = merged.get(identityKey);
    if (!existing) {
      merged.set(identityKey, { ...row, duplicateDisplayRows: 1 });
      return;
    }
    existing.duplicateDisplayRows = Number(existing.duplicateDisplayRows || 1) + 1;
    existing.workerIds = [...new Set([...(existing.workerIds || []), ...(row.workerIds || []), row.workerId, row.employeeId].filter(Boolean))];
    existing.punchCount = Number(existing.punchCount || 0) + Number(row.punchCount || 0);
    existing.canonicalMergedInMemory = true;
    if (!existing.name && row.name) existing.name = row.name;
  });
  return [...merged.values()].sort((a, b) => String(a.name || '').localeCompare(String(b.name || '')));
}

function getCanonicalPayrollWorkers(rows = getDerivedTimesheetRows()) {
  return mergeDuplicatePayrollRowsInMemory(rows).map((row) => ({
    identityKey: row.identityKey || row.workerId || row.employeeId || row.nameKey,
    name: row.name || row.workerName || '-',
    agencyId: row.agencyId || '',
    branchId: row.branchId || row.siteId || '',
    row
  })).filter((worker) => worker.identityKey);
}

function runPayrollCanonicalConsoleChecks(workers = getCanonicalPayrollWorkers()) {
  const weekKey = formatDateKey(state.selectedWeekStart);
  const logKey = `${weekKey}|${workers.map((worker) => worker.identityKey).join(',')}`;
  if (state.loggedPayrollCanonicalChecks.has(logKey)) return;
  state.loggedPayrollCanonicalChecks.add(logKey);

  const namesToCheck = ['Emanuel Palmer', 'Al-Lee Mayo', 'Ervin Wilson'];
  const checks = namesToCheck.map((name) => {
    const nameKey = normalizeName(name);
    const matches = workers.filter((worker) => normalizeName(worker.name) === nameKey);
    return {
      name,
      dropdownOptions: matches.length,
      selectedPunchesIncluded: matches.reduce((sum, worker) => sum + Number(worker.row?.punchCount || 0), 0),
      canonicalKeys: matches.map((worker) => worker.identityKey)
    };
  });

  console.info('[QR TimeClock Pro payroll canonical checks]', {
    weekKey,
    checks,
    expected: 'Each listed worker should appear once when present; selectedPunchesIncluded should include all weekly punches for that canonical worker.'
  });
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
    attachUsersViewIfAdmin();
    attachPendingUsersViewIfAdmin();
  }
}

const PUNCH_EDIT_ROLES = new Set([
  'admin',
  'manager',
  'supervisor',
  'super_admin',
  'superadmin',
  'owner',
  'agency_admin',
]);

const ADMIN_ROLES = new Set(['admin', 'super_admin', 'superadmin', 'owner']);
const PUNCH_DELETE_ROLES = new Set(['admin', 'manager', 'super_admin', 'superadmin', 'owner']);

function normalizeRole(value) {
  return String(value || '')
    .trim()
    .replace(/([a-z0-9])([A-Z])/g, '$1_$2')
    .toLowerCase()
    .replace(/[\s-]+/g, '_');
}

function normalizeRequestedAccessRole(value) {
  const role = normalizeRole(value);
  if (role === 'agency_admin') return 'agency_admin';
  if (role === 'supervisor') return 'supervisor';
  return 'manager';
}

function formatAuthError(error) {
  if (error?.code === 'auth/email-already-in-use') {
    return 'This email already has an account. Please sign in, then submit your access request.';
  }
  if (error?.code === 'auth/wrong-password' || error?.code === 'auth/invalid-credential') {
    return 'Incorrect password. Try again or send a password reset email.';
  }
  if (error?.code === 'auth/user-not-found') {
    return 'No account was found for that email. Request access first.';
  }
  if (error?.code === 'permission-denied') {
    return 'Permission denied. Your account may not have access to this action.';
  }
  return error?.message || 'The request could not be completed.';
}

function formatFirestoreError(error) {
  if (error?.code === 'permission-denied') {
    return 'Permission denied. Your account may not have access to this action.';
  }
  return formatAuthError(error);
}

function normalizedRole() {
  return normalizeRole(state.profile?.role);
}

function currentRole() {
  return normalizedRole();
}

function hasAnyRole(roles) {
  const role = normalizedRole();
  return roles.map(normalizeRole).includes(role);
}

function hasPermission(permissionName) {
  return state.profile?.permissions?.[permissionName] === true;
}

function defaultPermissionsForRole(role) {
  const normalized = normalizeRole(role);
  const fullAccess = isOwnerOrSuperAdminRole(normalized);
  return {
    canEditPunches: fullAccess || ['admin', 'agency_admin', 'manager', 'supervisor'].includes(normalized),
    canDeletePunches: fullAccess || normalized === 'admin',
    canMergeWorkers: fullAccess || normalized === 'admin',
    manageUsers: fullAccess || normalized === 'admin',
  };
}

function isOwnerOrSuperAdminRole(role) {
  return ['owner', 'super_admin', 'superadmin'].includes(normalizeRole(role));
}

function isOwnerOrSuperAdmin() {
  return isOwnerOrSuperAdminRole(state.profile?.role);
}

function canEditPunches() {
  if (isOwnerOrSuperAdmin()) return true;
  return PUNCH_EDIT_ROLES.has(normalizedRole()) && hasPermission('canEditPunches');
}

function canManageEmployees() {
  return canEditPunches();
}

function canDeletePunches() {
  if (isOwnerOrSuperAdmin()) return true;
  return PUNCH_DELETE_ROLES.has(normalizedRole()) && hasPermission('canDeletePunches');
}

function canMergeWorkers() {
  if (isOwnerOrSuperAdmin()) return true;
  return isAdmin() && hasPermission('canMergeWorkers');
}

function canManageUsers() {
  if (isOwnerOrSuperAdmin()) return true;
  return isAdmin() || hasPermission('manageUsers');
}

function isManager() {
  return PUNCH_EDIT_ROLES.has(normalizedRole());
}

function isAdmin() {
  return ADMIN_ROLES.has(normalizedRole());
}

function renderPermissionDebug() {
  if (!els.permissionDebugPanel) return;
  els.permissionDebugPanel.classList.toggle('hidden', !canManageUsers());
  if (!canManageUsers()) return;
  const rows = [
    ['Auth UID', state.me?.uid || '-'],
    ['Auth Email', state.me?.email || '-'],
    ['Loaded Profile', JSON.stringify(state.profile || {}, null, 2)],
    ['Normalized Role', normalizedRole() || '-'],
    ['Active', String(state.profile?.active === true)],
    ['approvalStatus', state.profile?.approvalStatus || '-'],
    ['companyId', state.profile?.companyId || state.companyId || '(blank legacy scope)'],
    ['siteId', state.profile?.siteId || '-'],
    ['siteIds', JSON.stringify(getAllowedSiteIds(state.profile))],
    ['permissions', JSON.stringify(state.profile?.permissions || {}, null, 2)],
    ['manageUsers', String(canManageUsers())],
    ['agencyId', state.agencyId || '(none)'],
    ['canEditPunches', String(canEditPunches())],
  ];
  els.permissionDebugPanel.innerHTML = rows.map(([label, value]) => `
    <div class="permission-debug-row">
      <span>${escapeHtml(label)}</span>
      <strong>${escapeHtml(value)}</strong>
    </div>
  `).join('');
}

function renderSiteSettingsForm() {
  if (!isAdmin() || !els.siteSettingsForm) return;
  if (els.siteSettingsId) els.siteSettingsId.value = state.siteContext.siteId || '';
  if (els.siteSettingsQrSlug) els.siteSettingsQrSlug.value = state.siteContext.qrSlug || '';
  if (els.siteSettingsLatitude) els.siteSettingsLatitude.value = state.siteContext.siteLatitude ?? '';
  if (els.siteSettingsLongitude) els.siteSettingsLongitude.value = state.siteContext.siteLongitude ?? '';
  if (els.siteSettingsRadius) els.siteSettingsRadius.value = String(state.siteContext.allowedRadiusMeters || 300);
  if (els.siteSettingsAccuracy) els.siteSettingsAccuracy.value = String(state.siteContext.maxGpsAccuracyMeters || 100);
  if (els.siteSettingsEnforce) els.siteSettingsEnforce.value = 'false';
}

async function handleSaveSiteSettings(event) {
  event.preventDefault();
  if (!isAdmin()) {
    toast('Only admins can save site settings.', true);
    return;
  }
  const siteId = String(els.siteSettingsId?.value || '').trim();
  if (!siteId) {
    toast('Enter a Site ID.', true);
    return;
  }
  const siteLatitude = finiteNumberOrNull(els.siteSettingsLatitude?.value);
  const siteLongitude = finiteNumberOrNull(els.siteSettingsLongitude?.value);
  if ((siteLatitude === null) !== (siteLongitude === null)) {
    toast('Enter both site latitude and longitude, or leave both blank.', true);
    return;
  }
  const payload = {
    companyId: state.companyId || '',
    siteId,
    qrSlug: String(els.siteSettingsQrSlug?.value || '').trim(),
    siteLatitude,
    siteLongitude,
    allowedRadiusMeters: positiveNumberOrDefault(els.siteSettingsRadius?.value, 300),
    maxGpsAccuracyMeters: positiveNumberOrDefault(els.siteSettingsAccuracy?.value, 100),
    enforceLocation: false,
    updatedAt: serverTimestamp(),
    updatedBy: state.me?.uid || '',
  };
  try {
    await setDoc(doc(db, 'sites', siteId), payload, { merge: true });
    if (state.siteContext.siteId === siteId) {
      state.siteContext = { ...state.siteContext, ...payload, enforceLocation: false };
    }
    toast('Site settings saved. Location remains optional.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not save site settings.', true);
  }
}

function punchTimestampMs(row) {
  const explicit = Number(row?.timestampMs || 0);
  if (explicit) return explicit;
  if (row?.timestamp?.toMillis instanceof Function) return row.timestamp.toMillis();
  if (row?.createdAt?.toMillis instanceof Function) return row.createdAt.toMillis();
  const parsed = Date.parse(String(row?.timestamp || row?.createdAt || ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizePunchAction(action) {
  const value = String(action || '').trim();
  const map = {
    clockIn: 'clock_in',
    clock_in: 'clock_in',
    startLunch: 'start_lunch',
    lunchStart: 'start_lunch',
    start_lunch: 'start_lunch',
    endLunch: 'end_lunch',
    lunchEnd: 'end_lunch',
    end_lunch: 'end_lunch',
    clockOut: 'clock_out',
    clock_out: 'clock_out'
  };
  return map[value] || value;
}

function normalizePunchRecordForDisplay(row = {}) {
  const timestampMs = punchTimestampMs(row);
  const name = row.name || row.workerName || row.employeeName || row.displayName || '';
  const action = normalizePunchAction(row.action || row.type || row.punchType || '');
  const dateKey = row.dateKey || row.localDate || (timestampMs ? formatDateKey(new Date(timestampMs)) : '');
  return {
    ...row,
    name,
    workerName: row.workerName || name,
    employeeName: row.employeeName || name,
    nameKey: row.nameKey || normalizeName(name),
    action,
    type: normalizePunchAction(row.type || action),
    timestampMs,
    dateKey,
    weekKey: row.weekKey || (dateKey ? formatDateKey(getMondayDate(new Date(`${dateKey}T12:00:00`))) : '')
  };
}

function prettyAction(action) {
  return String(normalizePunchAction(action) || '-')
    .replaceAll('_', ' ')
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function statusLabelForAction(action) {
  const normalized = normalizePunchAction(action);
  const map = {
    clock_in: 'Clocked In',
    start_lunch: 'On Lunch',
    end_lunch: 'Back From Lunch',
    clock_out: 'Clocked Out'
  };
  return map[normalized] || 'Saved';
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

function getWorkerProfileName(worker) {
  if (!worker) return '';
  const full = prettifyHumanName([worker.firstName, worker.lastName].filter(Boolean).join(' '));
  return prettifyHumanName(worker.name || worker.displayName || worker.employeeName || worker.workerName || full || '');
}

function getWorkerBranchId(workerOrRow) {
  return String(workerOrRow?.branchId || workerOrRow?.branchCode || workerOrRow?.siteId || workerOrRow?.assignedSiteId || '').trim();
}

function getWorkerBranchName(workerOrRow) {
  return String(workerOrRow?.branchName || workerOrRow?.siteName || getWorkerBranchId(workerOrRow) || '').trim();
}

function getWorkerProfileIds(worker) {
  return [
    worker?.id,
    worker?.employeeId,
    worker?.workerId,
    worker?.userId,
    worker?.uid
  ].map((value) => String(value || '').trim()).filter(Boolean);
}

function getRecordWorkerIds(row) {
  return [
    row?.employeeId,
    row?.workerId,
    row?.userId,
    row?.uid
  ].map((value) => String(value || '').trim()).filter(Boolean);
}

function getTrustedRecordWorkerIds(row, profileIndex = buildWorkerProfileIndex()) {
  const copiedName = getCopiedWorkerName(row);
  return getRecordWorkerIds(row).filter((id) => {
    const profile = profileIndex.get(id);
    const profileName = getWorkerProfileName(profile);
    return !copiedName || !profileName || !namesDiffer(profileName, copiedName);
  });
}

function getDirectoryWorkerIds(row, profileIndex = buildWorkerProfileIndex()) {
  return row?.action || row?.timestampMs || row?.dateKey || row?.weekKey
    ? getTrustedRecordWorkerIds(row, profileIndex)
    : [...getWorkerProfileIds(row), ...getTrustedRecordWorkerIds(row, profileIndex)].filter(Boolean);
}

function buildWorkerProfileIndex(workers = state.allEmployees || [], users = state.allUsers || []) {
  const byId = new Map();
  (workers || []).forEach((worker) => {
    getWorkerProfileIds(worker).forEach((id) => byId.set(id, worker));
  });
  (users || []).forEach((user) => {
    getWorkerProfileIds(user).forEach((id) => {
      if (!byId.has(id)) byId.set(id, user);
    });
  });
  return byId;
}

function getCopiedWorkerName(row) {
  return prettifyHumanName(row?.employeeName || row?.workerName || row?.displayName || row?.name || '');
}

function normalizeIdentityText(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ');
}

function normalizeIdentityToken(value) {
  return normalizeIdentityText(value)
    .replace(/[^a-z0-9@._ -]/g, '')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function normalizeEmail(value) {
  const email = normalizeIdentityText(value);
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email) ? email : '';
}

function getRecordEmail(row) {
  return normalizeEmail(row?.email || row?.workerEmail || row?.employeeEmail || row?.userEmail);
}

function getRecordAgencyIdentity(row) {
  return normalizeIdentityToken(row?.agencyId || row?.agencyName || row?.staffingAgency || state.agencyId || '');
}

function getRecordBranchIdentity(row) {
  return normalizeIdentityToken(
    row?.branchId
    || row?.branchCode
    || row?.branchName
    || row?.siteId
    || row?.assignedSiteId
    || row?.siteName
    || ''
  );
}

function getWorkerSignature(row) {
  const name = normalizeIdentityToken(getCopiedWorkerName(row));
  const agency = getRecordAgencyIdentity(row);
  const branch = getRecordBranchIdentity(row);
  return name && (agency || branch) ? `${name}|${agency}|${branch}` : '';
}

function buildCanonicalWorkerDirectory(records = [], profileIndex = buildWorkerProfileIndex()) {
  const sourceRows = [
    ...(state.allEmployees || []),
    ...(state.publicEmployeeRecords || []),
    ...(state.allUsers || []),
    ...(state.selectedWeekPunchRows || []),
    ...Object.values(state.selectedWeekTimesheetDocs || {}),
    ...(records || [])
  ];
  const signatureIds = new Map();
  const emailIds = new Map();

  sourceRows.forEach((row) => {
    const ids = getDirectoryWorkerIds(row, profileIndex);
    const email = getRecordEmail(row);
    const signature = getWorkerSignature(row);
    if (signature && ids.length) {
      if (!signatureIds.has(signature)) signatureIds.set(signature, new Set());
      ids.forEach((id) => signatureIds.get(signature).add(id));
    }
    if (email && ids.length) {
      if (!emailIds.has(email)) emailIds.set(email, new Set());
      ids.forEach((id) => emailIds.get(email).add(id));
    }
  });

  const primaryFor = (ids) => [...ids].sort((left, right) => String(left).localeCompare(String(right)))[0];
  const signaturePrimary = new Map();
  signatureIds.forEach((ids, signature) => signaturePrimary.set(signature, primaryFor(ids)));
  const emailPrimary = new Map();
  emailIds.forEach((ids, email) => emailPrimary.set(email, primaryFor(ids)));
  return { signaturePrimary, emailPrimary };
}

function getWorkerIdentityKey(row, directory = buildCanonicalWorkerDirectory(), profileIndex = buildWorkerProfileIndex()) {
  const stableId = getDirectoryWorkerIds(row, profileIndex)[0] || '';
  const email = getRecordEmail(row);
  const signature = getWorkerSignature(row);
  const signaturePrimaryId = signature ? directory.signaturePrimary.get(signature) : '';
  const emailPrimaryId = email ? directory.emailPrimary.get(email) : '';

  if (signaturePrimaryId) return `worker:${signaturePrimaryId}`;
  if (stableId) return `worker:${stableId}`;
  if (emailPrimaryId) return `worker:${emailPrimaryId}`;
  if (email) return `email:${email}`;
  return signature ? `person:${signature}` : '';
}

function sanitizeTimesheetIdPart(value) {
  return String(value || '')
    .trim()
    .replace(/[^a-zA-Z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || 'worker';
}

function namesDiffer(left, right) {
  const leftKey = normalizeName(left);
  const rightKey = normalizeName(right);
  return !!leftKey && !!rightKey && leftKey !== rightKey;
}

function logWorkerNameDiagnostic({ workerId, profileName, copiedName, agencyId, agencyName, branchId, branchName, weekStart, source }) {
  if (!isAdmin() && !/^(localhost|127\.0\.0\.1)$/i.test(window.location.hostname)) return;
  if (!namesDiffer(profileName, copiedName)) return;
  const key = [source, workerId, profileName, copiedName, agencyId, branchId, weekStart].join('|');
  if (state.loggedWorkerNameDiagnostics.has(key)) return;
  state.loggedWorkerNameDiagnostics.add(key);
  console.warn('[QRTimeclock worker name mismatch]', {
    workerId,
    profileName,
    copiedPunchOrTimesheetName: copiedName,
    agencyId,
    agencyName,
    branchId,
    branchName,
    weekStart,
    source
  });
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
  const role = profile.role === 'worker' ? 'worker' : (profile.role || 'worker');
  const defaultPermissions = defaultPermissionsForRole(role);
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
    companyId: profile.companyId || CURRENT_COMPANY_ID,
    agencyId: profile.agencyId || '',
    branch: profile.branch || profile.siteId || profile.assignedSiteId || '',
    branches: Array.isArray(profile.branches) ? parseSiteIds(profile.branches, false) : [],
    siteId: profile.siteId || profile.branch || profile.assignedSiteId || '',
    siteIds: parseSiteIds(profile.branches || profile.branch || profile.siteIds || profile.siteId || profile.assignedSiteId || '', false),
    permissions: {
      canEditPunches: defaultPermissions.canEditPunches || profile.permissions?.canEditPunches === true,
      canDeletePunches: defaultPermissions.canDeletePunches || profile.permissions?.canDeletePunches === true,
      canMergeWorkers: defaultPermissions.canMergeWorkers || profile.permissions?.canMergeWorkers === true,
      manageUsers: defaultPermissions.manageUsers || profile.permissions?.manageUsers === true
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
    addMatches(await getDocsCounted(query(
      employeesRef,
      where('companyId', '==', normalizedCompanyId),
      where('agencyId', '==', normalizedAgencyId),
      where('assignedSiteId', '==', normalizedSiteId),
      where('status', '==', 'active'),
      where('employeeNumberKey', '==', employeeNumberKey)
    ), 'Put Away save'));
    if (!matches.length) {
      addMatches(await getDocsCounted(query(
        employeesRef,
        where('companyId', '==', normalizedCompanyId),
        where('agencyId', '==', normalizedAgencyId),
        where('assignedSiteId', '==', normalizedSiteId),
        where('status', '==', 'active'),
        where('employeeNumber', '==', employeeNumber)
      ), 'Put Away save'));
    }
  }

  if (!matches.length && nameKey) {
    addMatches(await getDocsCounted(query(
      employeesRef,
      where('companyId', '==', normalizedCompanyId),
      where('agencyId', '==', normalizedAgencyId),
      where('assignedSiteId', '==', normalizedSiteId),
      where('status', '==', 'active'),
      where('nameKey', '==', nameKey)
    ), 'Put Away save'));
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

  invalidateEmployeeCaches();
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

function formatTimestamp(value) {
  if (!value) return '';
  if (typeof value === 'number') return formatDateTime(value);
  if (value.seconds) return formatDateTime(value.seconds * 1000);
  if (value.toDate) return formatDateTime(value.toDate().getTime());
  return '';
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
