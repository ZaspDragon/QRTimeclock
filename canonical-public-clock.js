import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  addDoc,
  collection,
  doc,
  getDoc,
  getDocs,
  getFirestore,
  limit,
  query,
  serverTimestamp,
  setDoc,
  where,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

// Public-clock identity guard.
// One person keeps one employee ID even if their staffing agency changes.
// Historical punches are never deleted or rewritten by this module.

const FIREBASE_CONFIG = {
  apiKey: 'AIzaSyB4xdaxbkXDRILPe2nGZuGCS-PXf35bk3o',
  authDomain: 'qrtimeclock-42764.firebaseapp.com',
  projectId: 'qrtimeclock-42764',
  storageBucket: 'qrtimeclock-42764.appspot.com',
  messagingSenderId: '232535382723',
  appId: '1:232535382723:web:9fe08f4961d87ba4062076',
};

const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_AGENCIES = new Map([
  ['sterling_staffing', 'Sterling Staffing'],
  ['excel_staffing', 'Excel Staffing'],
  ['lifestyle_staffing', 'Lifestyle Staffing'],
]);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ACTION_LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Start Lunch',
  end_lunch: 'End Lunch',
  clock_out: 'Clock Out',
};

// Known canonical identities are used only to choose among existing duplicate
// profiles. No employee or punch record is rewritten by this mapping.
const KNOWN_CANONICAL_WORKERS = new Map([
  ['donald gibson|OH01|sterling_staffing', 'EMP-1058'],
]);

let saving = false;

function dbInstance() {
  const app = getApps().length ? getApp() : initializeApp(FIREBASE_CONFIG);
  return getFirestore(app);
}

function prettyName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase());
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

function safeIdPart(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 48) || 'worker';
}

function localDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function mondayKey(date) {
  const monday = new Date(date);
  const day = monday.getDay();
  monday.setDate(monday.getDate() - (day === 0 ? 6 : day - 1));
  monday.setHours(0, 0, 0, 0);
  return localDateKey(monday);
}

function selectedSite() {
  const querySite = String(new URLSearchParams(location.search).get('site') || '').toUpperCase();
  if (VALID_SITES.has(querySite)) return querySite;
  const selected = String(document.getElementById('workerBranchSelect')?.value || '').toUpperCase();
  return VALID_SITES.has(selected) ? selected : '';
}

function selectedAgency() {
  const value = String(document.getElementById('workerAgencySelect')?.value || '').trim();
  return VALID_AGENCIES.has(value) ? value : '';
}

function employeeName(row) {
  return String(row?.name || row?.employeeName || row?.displayName || row?.fullName || '').trim();
}

function employeeSite(row) {
  return String(row?.siteId || row?.assignedSiteId || row?.branch || row?.branchId || '').trim().toUpperCase();
}

function employeeAgency(row) {
  return String(row?.agencyId || row?.staffingAgencyId || '').trim();
}

function employeeIdentityValues(row) {
  return [
    row?.id,
    row?.employeeId,
    row?.employeeID,
    row?.employeeNumber,
    row?.canonicalEmployeeId,
    row?.workerId,
  ].map((value) => String(value || '').trim()).filter(Boolean);
}

function isActive(row) {
  return row?.active !== false && !['inactive', 'terminated', 'merged', 'deleted', 'removed', 'archived'].includes(String(row?.status || '').toLowerCase());
}

function setMessage(message, isError = false) {
  const lookup = document.getElementById('workerLookupStatus');
  const status = document.getElementById('workerStatusMessage');
  const stateValue = document.getElementById('workerStatusValue');
  if (lookup) {
    lookup.textContent = message;
    lookup.style.borderColor = isError ? 'rgba(255,90,90,.6)' : 'rgba(43,213,118,.5)';
  }
  if (status) status.textContent = message;
  if (stateValue) stateValue.textContent = isError ? 'Needs attention' : 'Saved';
}

function disableButtons(disabled) {
  document.querySelectorAll('.worker-action-btn').forEach((button) => {
    button.disabled = disabled;
    button.setAttribute('aria-busy', disabled ? 'true' : 'false');
  });
}

function installAgencyControl() {
  if (document.getElementById('workerAgencySelect')) return;
  const branchSelect = document.getElementById('workerBranchSelect');
  const branchLabel = branchSelect?.closest('label');
  if (!branchLabel) return;

  const label = document.createElement('label');
  label.id = 'workerAgencyField';
  label.innerHTML = `
    <span>Staffing agency <small>(required for temps)</small></span>
    <select id="workerAgencySelect" required aria-required="true">
      <option value="">Choose your staffing agency</option>
      ${[...VALID_AGENCIES.entries()].map(([value, text]) => `<option value="${value}">${text}</option>`).join('')}
    </select>
  `;
  branchLabel.insertAdjacentElement('afterend', label);

  const select = label.querySelector('select');
  const saved = String(localStorage.getItem('workerPunchAgency') || '').trim();
  if (VALID_AGENCIES.has(saved)) select.value = saved;
  select.addEventListener('change', () => {
    if (VALID_AGENCIES.has(select.value)) localStorage.setItem('workerPunchAgency', select.value);
  });
}

function stickyIdentityKey(name, siteId, agencyId) {
  return `canonicalWorkerIdentity:${siteId}:${agencyId}:${nameKey(name)}`;
}

async function loadStickyWorker(db, name, siteId, agencyId) {
  const key = stickyIdentityKey(name, siteId, agencyId);
  const employeeId = String(localStorage.getItem(key) || '').trim();
  if (!employeeId) return null;

  try {
    const snapshot = await getDoc(doc(db, 'employees', employeeId));
    if (!snapshot.exists()) {
      localStorage.removeItem(key);
      return null;
    }
    const row = { id: snapshot.id, ...snapshot.data() };
    const site = employeeSite(row);
    if (
      !isActive(row)
      || normalizeName(employeeName(row)) !== normalizeName(name)
      || (site && site !== siteId)
    ) {
      localStorage.removeItem(key);
      return null;
    }
    return row;
  } catch (_) {
    return null;
  }
}

function saveStickyWorker(name, siteId, agencyId, employee) {
  const employeeId = String(employee?.employeeId || employee?.id || '').trim();
  if (!employeeId) return;
  localStorage.setItem(stickyIdentityKey(name, siteId, agencyId), employeeId);
}

async function loadActiveExactNameMatches(db, name, siteId) {
  const normalized = normalizeName(name);
  const key = nameKey(name);
  const searches = [
    query(collection(db, 'employees'), where('active', '==', true), where('nameKey', '==', key), limit(30)),
    query(collection(db, 'employees'), where('status', '==', 'active'), where('nameKey', '==', key), limit(30)),
    query(collection(db, 'employees'), where('active', '==', true), limit(500)),
    query(collection(db, 'employees'), where('status', '==', 'active'), limit(500)),
  ];

  const rows = new Map();
  const results = await Promise.allSettled(searches.map((search) => getDocs(search)));
  results.forEach((result) => {
    if (result.status !== 'fulfilled') return;
    result.value.docs.forEach((snapshot) => rows.set(snapshot.id, { id: snapshot.id, ...snapshot.data() }));
  });

  return [...rows.values()]
    .filter((row) =>
      isActive(row)
      && String(row.companyId || COMPANY_ID).trim() === COMPANY_ID
      && normalizeName(employeeName(row)) === normalized
      && (!employeeSite(row) || employeeSite(row) === siteId)
    );
}

function chooseExistingWorker(matches, name, siteId, agencyId) {
  if (!matches.length) return { employee: null, agencyUpdate: false };

  const knownIdentity = KNOWN_CANONICAL_WORKERS.get(`${normalizeName(name)}|${siteId}|${agencyId}`);
  if (knownIdentity) {
    const knownMatches = matches.filter((row) => employeeIdentityValues(row).includes(knownIdentity));
    if (knownMatches.length === 1) {
      return {
        employee: knownMatches[0],
        agencyUpdate: employeeAgency(knownMatches[0]) !== agencyId,
      };
    }
  }

  const exactAgency = matches.filter((row) => employeeAgency(row) === agencyId);
  if (exactAgency.length === 1) return { employee: exactAgency[0], agencyUpdate: false };

  if (exactAgency.length > 1) {
    const canonical = exactAgency.filter((row) => {
      const canonicalId = String(row.canonicalEmployeeId || '').trim();
      return canonicalId && employeeIdentityValues(row).includes(canonicalId);
    });
    if (canonical.length === 1) return { employee: canonical[0], agencyUpdate: false };

    const numbered = exactAgency.filter((row) => /^EMP[-_ ]?\d+$/i.test(String(row.employeeNumber || row.employeeID || '')));
    if (numbered.length === 1) return { employee: numbered[0], agencyUpdate: false };

    throw new Error('Duplicate worker profiles already exist for this name and agency. Ask a manager to merge them before punching.');
  }

  const blankAgency = matches.filter((row) => !employeeAgency(row));
  if (blankAgency.length === 1 && matches.length === 1) {
    return { employee: blankAgency[0], agencyUpdate: true };
  }

  // Critical duplicate-prevention rule: if exactly one active person exists at this
  // branch, keep that employee ID and update only the current agency assignment.
  // Old punches retain their original agencyId, timestamps and hours.
  if (matches.length === 1) {
    return { employee: matches[0], agencyUpdate: employeeAgency(matches[0]) !== agencyId };
  }

  const canonical = matches.filter((row) => {
    const canonicalId = String(row.canonicalEmployeeId || '').trim();
    return canonicalId && employeeIdentityValues(row).includes(canonicalId);
  });
  if (canonical.length === 1) {
    return { employee: canonical[0], agencyUpdate: employeeAgency(canonical[0]) !== agencyId };
  }

  // Multiple unresolved active same-name profiles are still blocked so the public
  // clock never creates another duplicate or guesses between two real people.
  throw new Error('More than one active worker profile uses this name. Ask a manager to merge the duplicates before punching.');
}

async function createNewCanonicalWorker(db, name, siteId, agencyId) {
  const normalized = normalizeName(name);
  const key = nameKey(name);
  const employeeId = `public_canonical_${safeIdPart(siteId)}_${safeIdPart(key)}`;
  const employeeNumber = `PUBLIC-${safeIdPart(siteId).toUpperCase()}-${safeIdPart(key).toUpperCase()}`.slice(0, 60);
  const ref = doc(db, 'employees', employeeId);
  const existing = await getDoc(ref).catch(() => null);

  if (existing?.exists()) {
    const row = { id: existing.id, ...existing.data() };
    if (!isActive(row)) {
      throw new Error('An inactive worker profile already exists for this name. Ask a manager to reactivate or repair it.');
    }
    return row;
  }

  const payload = {
    name,
    nameKey: key,
    normalizedName: key,
    employeeNumber,
    employeeNumberKey: employeeNumber.toLowerCase(),
    companyId: COMPANY_ID,
    agencyId,
    assignedSiteId: siteId,
    siteId,
    siteIds: [siteId],
    status: 'active',
    active: true,
    employeeId,
    canonicalEmployeeId: employeeId,
    source: 'canonical_public_auto_created',
    identityVersion: 2,
    updatedAt: serverTimestamp(),
    createdAt: serverTimestamp(),
  };
  await setDoc(ref, payload, { merge: false });
  return { id: employeeId, ...payload };
}

async function resolveWorker(db, name, siteId, agencyId) {
  const sticky = await loadStickyWorker(db, name, siteId, agencyId);
  if (sticky) return sticky;

  const matches = await loadActiveExactNameMatches(db, name, siteId);
  const choice = chooseExistingWorker(matches, name, siteId, agencyId);

  if (!choice.employee) {
    return createNewCanonicalWorker(db, name, siteId, agencyId);
  }

  const employee = choice.employee;
  if (choice.agencyUpdate && employee.id) {
    const previousAgencyId = employeeAgency(employee);
    await setDoc(doc(db, 'employees', employee.id), {
      agencyId,
      previousAgencyId,
      agencyChangedAt: serverTimestamp(),
      agencyAssignmentSource: 'canonical_public_selection',
      canonicalEmployeeId: employee.canonicalEmployeeId || employee.id,
      updatedAt: serverTimestamp(),
    }, { merge: true });
    return { ...employee, agencyId, previousAgencyId, canonicalEmployeeId: employee.canonicalEmployeeId || employee.id };
  }

  return employee;
}

async function savePunch(action) {
  const workerCard = document.getElementById('workerCard');
  if (!workerCard || workerCard.classList.contains('hidden')) return;

  const name = prettyName(document.getElementById('workerNameInput')?.value);
  const siteId = selectedSite();
  const agencyId = selectedAgency();
  if (normalizeName(name).length < 2) throw new Error('Type your first and last name before punching.');
  if (!siteId) throw new Error('Choose OH01 or OHC before punching.');
  if (!agencyId) throw new Error('Choose Sterling, Excel, or Lifestyle Staffing before punching.');
  if (!VALID_ACTIONS.has(action)) throw new Error('That punch type is not valid.');

  const db = dbInstance();
  const employee = await resolveWorker(db, name, siteId, agencyId);
  const employeeId = String(employee.employeeId || employee.id || '').trim();
  const employeeNumber = String(employee.employeeNumber || employeeId).trim();
  if (!employeeId) throw new Error('The worker profile could not be verified. Ask a manager to check the employee record.');

  const now = new Date();
  const nowMs = Date.now();
  const duplicateKey = `canonicalPublicPunch:${employeeId}:${action}:${localDateKey(now)}`;
  const prior = Number(localStorage.getItem(duplicateKey) || 0);
  if (prior && nowMs - prior < 60000) {
    throw new Error(`${ACTION_LABELS[action]} was already saved. No second tap is needed.`);
  }

  await addDoc(collection(db, 'punches'), {
    companyId: COMPANY_ID,
    siteId,
    siteIds: [siteId],
    assignedSiteId: siteId,
    agencyId,
    employeeId,
    workerId: employee.canonicalEmployeeId || employee.workerId || employeeId,
    canonicalEmployeeId: employee.canonicalEmployeeId || employeeId,
    employeeNumber,
    name: employeeName(employee) || name,
    nameKey: nameKey(employeeName(employee) || name),
    action,
    timestamp: serverTimestamp(),
    timestampMs: nowMs,
    dateKey: localDateKey(now),
    weekKey: mondayKey(now),
    source: 'public_qr_canonical',
    createdAt: serverTimestamp(),
    locationStatus: 'not_requested',
    enforceLocation: false,
    active: true,
    status: 'active',
  });

  saveStickyWorker(name, siteId, agencyId, employee);
  localStorage.setItem(duplicateKey, String(nowMs));
  localStorage.setItem('workerPunchName', employeeName(employee) || name);
  localStorage.setItem('workerPunchAgency', agencyId);
  if (document.getElementById('workerNameValue')) document.getElementById('workerNameValue').textContent = employeeName(employee) || name;
  if (document.getElementById('workerLastActionValue')) document.getElementById('workerLastActionValue').textContent = ACTION_LABELS[action];
  if (document.getElementById('workerLastPunchValue')) document.getElementById('workerLastPunchValue').textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  setMessage(`${ACTION_LABELS[action]} saved for ${employeeName(employee) || name}.`);
}

function installClickGuard() {
  if (document.documentElement.dataset.canonicalPublicClockInstalled === 'true') return;
  document.documentElement.dataset.canonicalPublicClockInstalled = 'true';

  // Register at capture phase before the older hotfix handlers. This makes this
  // module the single writer for public punch-button clicks.
  document.addEventListener('click', async (event) => {
    const button = event.target.closest?.('.worker-action-btn');
    const workerCard = document.getElementById('workerCard');
    if (!button || !workerCard || workerCard.classList.contains('hidden') || saving) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    saving = true;
    disableButtons(true);
    const action = String(button.dataset.action || '');
    const name = prettyName(document.getElementById('workerNameInput')?.value);
    setMessage(`Saving ${ACTION_LABELS[action] || 'punch'} for ${name || 'worker'}...`);

    try {
      await savePunch(action);
      window.setTimeout(() => window.location.reload(), 900);
    } catch (error) {
      console.error('[canonical-public-clock]', error);
      setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
      saving = false;
      disableButtons(false);
    }
  }, true);
}

installClickGuard();
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', installAgencyControl, { once: true });
} else {
  installAgencyControl();
}

console.info('[QRTimeclock] Canonical public-clock returning-worker identity guard installed.');