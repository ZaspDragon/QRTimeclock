import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { collection, getDocs, getFirestore, limit, query, where } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

// Read-only guard for employee self-service time lookup.
// Agency Export worker options are built from active employee records using
// name + agency + branch. This guard checks that same registration contract
// before allowing My Hours / Past Time requests to continue.
// It never creates/updates employees and never touches punch buttons or punches.

const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const AGENCY_NAMES = {
  sterling_staffing: 'Sterling Staffing',
  excel_staffing: 'Excel Staffing',
  lifestyle_staffing: 'Lifestyle Staffing',
};
const LOOKUP_SELECTOR = '#workerViewTimeBtn, #workerViewMoreTimeBtn, #workerTimeLookupBtn, .worker-range-quick';
let checking = false;
let bypassNextLookup = false;
let checkGeneration = 0;

function dbInstance() {
  const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
  return getFirestore(app);
}

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function nameKey(value) {
  return normalizeName(value).replaceAll(' ', '_');
}

function enteredName() {
  return String(document.getElementById('workerNameInput')?.value || '').trim().replace(/\s+/g, ' ');
}

function selectedSite() {
  const querySite = String(new URLSearchParams(location.search).get('site') || '').trim().toUpperCase();
  if (VALID_SITES.has(querySite)) return querySite;
  const selected = String(document.getElementById('workerBranchSelect')?.value || '').trim().toUpperCase();
  return VALID_SITES.has(selected) ? selected : '';
}

function selectedAgency() {
  const selected = String(document.getElementById('workerAgencySelect')?.value || '').trim();
  return Object.hasOwn(AGENCY_NAMES, selected) ? selected : '';
}

function workerName(row) {
  return String(row?.name || row?.employeeName || row?.displayName || row?.fullName || row?.nameKey || '').trim();
}

function workerSite(row) {
  return String(row?.siteId || row?.assignedSiteId || row?.branch || row?.branchId || '').trim().toUpperCase();
}

function workerAgency(row) {
  return String(row?.agencyId || row?.staffingAgencyId || '').trim();
}

function activeWorker(row) {
  return row?.active === true || String(row?.status || '').toLowerCase() === 'active';
}

function uniqueRows(results) {
  const rows = new Map();
  results.forEach((result) => {
    if (result.status !== 'fulfilled') return;
    result.value.docs.forEach((snapshot) => rows.set(snapshot.id, { id: snapshot.id, ...snapshot.data() }));
  });
  return [...rows.values()];
}

async function loadActiveExactNameMatches(name) {
  const db = dbInstance();
  const key = nameKey(name);
  const normalized = normalizeName(name);
  const employees = collection(db, 'employees');
  const searches = [
    query(employees, where('active', '==', true), where('nameKey', '==', key), limit(30)),
    query(employees, where('status', '==', 'active'), where('nameKey', '==', key), limit(30)),
    query(employees, where('active', '==', true), limit(500)),
    query(employees, where('status', '==', 'active'), limit(500)),
  ];
  const results = await Promise.allSettled(searches.map((entry) => getDocs(entry)));
  return uniqueRows(results).filter((row) =>
    activeWorker(row)
    && String(row.companyId || COMPANY_ID).trim() === COMPANY_ID
    && normalizeName(workerName(row)) === normalized
  );
}

function registrationResult(matches, siteId, agencyId) {
  if (!matches.length) {
    return {
      registered: false,
      message: `Not registered in Agency Export yet. Missing: an active employee registration for ${siteId}. Ask a manager or agency admin to add or activate your worker profile.`,
    };
  }

  const siteMatches = matches.filter((row) => workerSite(row) === siteId);
  const blankSite = matches.filter((row) => !workerSite(row));
  if (!siteMatches.length) {
    if (blankSite.length) {
      return {
        registered: false,
        message: `Agency Export registration is incomplete. Missing: branch assignment (${siteId}). Ask a manager or agency admin to assign your branch.`,
      };
    }
    const sites = [...new Set(matches.map(workerSite).filter(Boolean))].join(' / ');
    return {
      registered: false,
      message: `Agency Export has you registered under ${sites || 'another branch'}, not ${siteId}. Choose the registered branch or ask a manager to correct your branch assignment.`,
    };
  }

  const exactAgency = siteMatches.filter((row) => workerAgency(row) === agencyId);
  if (exactAgency.length === 1) {
    return { registered: true, workerId: exactAgency[0].id };
  }
  if (exactAgency.length > 1) {
    return {
      registered: false,
      message: 'Agency Export has duplicate active registrations for this name, branch, and agency. Ask a manager to merge the duplicate profiles before viewing time.',
    };
  }

  const blankAgency = siteMatches.filter((row) => !workerAgency(row));
  if (blankAgency.length) {
    return {
      registered: false,
      message: `Agency Export registration is incomplete. Missing: staffing agency (${AGENCY_NAMES[agencyId]}). Ask a manager or agency admin to assign your staffing agency.`,
    };
  }

  const agencies = [...new Set(siteMatches.map(workerAgency).filter(Boolean))];
  const labels = agencies.map((id) => AGENCY_NAMES[id] || id).join(' / ');
  return {
    registered: false,
    message: `Agency Export has you registered under ${labels || 'another staffing agency'}, not ${AGENCY_NAMES[agencyId]}. Choose the registered agency or ask a manager to correct your agency assignment.`,
  };
}

async function checkRegistration() {
  const name = enteredName();
  const siteId = selectedSite();
  const agencyId = selectedAgency();
  if (normalizeName(name).length < 2) {
    return { registered: false, message: 'Enter your first and last name before looking up time.' };
  }
  if (!siteId) {
    return { registered: false, message: 'Choose OH01 or OHC before looking up time.' };
  }
  if (!agencyId) {
    return { registered: false, message: 'Choose Sterling, Excel, or Lifestyle Staffing before looking up time.' };
  }
  const matches = await loadActiveExactNameMatches(name);
  return registrationResult(matches, siteId, agencyId);
}

function setHoursStatus(message) {
  document.getElementById('workerMyTimePanel')?.classList.remove('hidden');
  const status = document.getElementById('workerTimeRangeStatus');
  if (status) status.textContent = message;
}

function clearHoursSummary() {
  ['workerWeekHoursValue', 'workerRegularHoursValue', 'workerOvertimeHoursValue', 'workerDaysWorkedValue'].forEach((id) => {
    const node = document.getElementById(id);
    if (node) node.textContent = '—';
  });
  document.getElementById('workerTimeRangeResults')?.replaceChildren();
}

function setLookupControlsDisabled(disabled) {
  ['workerViewTimeBtn', 'workerViewMoreTimeBtn', 'workerTimeLookupBtn'].forEach((id) => {
    const button = document.getElementById(id);
    if (button) button.disabled = disabled;
  });
}

window.addEventListener('click', async (event) => {
  const target = event.target instanceof Element ? event.target.closest(LOOKUP_SELECTOR) : null;
  if (!target || bypassNextLookup) return;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
  if (checking) return;

  checking = true;
  const generation = ++checkGeneration;
  setLookupControlsDisabled(true);
  clearHoursSummary();
  setHoursStatus('Checking Agency Export registration...');

  try {
    const result = await checkRegistration();
    if (generation !== checkGeneration) return;
    if (!result.registered) {
      setHoursStatus(result.message);
      return;
    }

    setHoursStatus('Agency Export registration confirmed. Loading saved time...');
    checking = false;
    setLookupControlsDisabled(false);
    bypassNextLookup = true;
    try {
      target.click();
    } finally {
      bypassNextLookup = false;
    }
  } catch (error) {
    if (generation !== checkGeneration) return;
    console.error('[agency-registration-guard]', error);
    setHoursStatus('Could not verify your Agency Export registration. Please try again or ask a manager. You can still clock in and out.');
  } finally {
    if (generation === checkGeneration && checking) {
      checking = false;
      setLookupControlsDisabled(false);
    }
  }
}, true);

function invalidateRegistration(event) {
  if (!['workerNameInput', 'workerBranchSelect', 'workerAgencySelect'].includes(event.target?.id)) return;
  checkGeneration += 1;
  checking = false;
  setLookupControlsDisabled(false);
}
document.addEventListener('input', invalidateRegistration);
document.addEventListener('change', invalidateRegistration);

console.info('[QRTimeclock] Agency Export registration guard installed for employee time lookup.');
