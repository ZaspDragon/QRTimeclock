import { registerPunchWriter, PUNCH_WRITER_PRIORITY } from './punch-writer-lock.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  addDoc,
  collection,
  doc,
  getDocs,
  getFirestore,
  limit,
  query,
  serverTimestamp,
  setDoc,
  where,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

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
const VALID_AGENCIES = new Set(['sterling_staffing', 'excel_staffing', 'lifestyle_staffing']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ACTION_LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Start Lunch',
  end_lunch: 'End Lunch',
  clock_out: 'Clock Out',
};
let saving = false;

function prettyName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, (character) => character.toUpperCase());
}

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function safeIdPart(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9_-]/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '').slice(0, 48) || 'worker';
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

function selectedSite() {
  const querySite = String(new URLSearchParams(location.search).get('site') || '').toUpperCase();
  if (VALID_SITES.has(querySite)) return querySite;
  const selected = String(document.getElementById('workerBranchSelect')?.value || '').toUpperCase();
  return VALID_SITES.has(selected) ? selected : '';
}

function employeeName(row) {
  return String(row.name || row.employeeName || row.displayName || row.fullName || '').trim();
}

function employeeSite(row) {
  return String(row.siteId || row.assignedSiteId || row.branch || row.branchId || '').trim().toUpperCase();
}

function employeeAgency(row) {
  return String(row.agencyId || row.staffingAgencyId || '').trim();
}

function isActive(row) {
  return row.active !== false && !['inactive', 'terminated', 'merged', 'deleted'].includes(String(row.status || '').toLowerCase());
}

async function findExistingEmployee(db, name, siteId, agencyId) {
  const normalized = normalizeName(name);
  const nameKey = normalized.replaceAll(' ', '_');
  const searches = [
    query(collection(db, 'employees'), where('active', '==', true), where('nameKey', '==', nameKey), limit(20)),
    query(collection(db, 'employees'), where('status', '==', 'active'), where('nameKey', '==', nameKey), limit(20)),
    query(collection(db, 'employees'), where('active', '==', true), limit(500)),
    query(collection(db, 'employees'), where('status', '==', 'active'), limit(500)),
  ];

  const rows = new Map();
  const results = await Promise.allSettled(searches.map((search) => getDocs(search)));
  results.forEach((result) => {
    if (result.status !== 'fulfilled') return;
    result.value.docs.forEach((snapshot) => rows.set(snapshot.id, { id: snapshot.id, ...snapshot.data() }));
  });

  const exactName = [...rows.values()].filter((row) =>
    isActive(row)
    && normalizeName(employeeName(row)) === normalized
    && (!employeeSite(row) || employeeSite(row) === siteId)
  );
  const exactAgency = exactName.filter((row) => employeeAgency(row) === agencyId);
  if (exactAgency.length === 1) return { employee: exactAgency[0], assignAgency: false };

  const blankAgency = exactName.filter((row) => !employeeAgency(row));
  if (blankAgency.length === 1) return { employee: blankAgency[0], assignAgency: true };

  return null;
}

function publicEmployeePayload({ employeeId, employeeNumber, name, nameKey, normalized, agencyId, siteId }) {
  return {
    name,
    nameKey,
    normalizedName: normalized,
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
    source: 'auto_created',
    updatedAt: serverTimestamp(),
    createdAt: serverTimestamp(),
  };
}

async function ensurePublicEmployee(db, name, normalized, siteId, agencyId) {
  const nameKey = normalized.replaceAll(' ', '_');
  const base = `${safeIdPart(siteId)}_${safeIdPart(agencyId)}_${safeIdPart(nameKey)}`;
  const candidates = [`public_${base}`, `public_returning_${base}`, `public_returning2_${base}`, `public_returning3_${base}`];
  let lastError = null;

  for (const employeeId of candidates) {
    const employeeNumber = `PUBLIC-${safeIdPart(siteId).toUpperCase()}-${safeIdPart(agencyId).toUpperCase()}-${safeIdPart(nameKey).toUpperCase()}-${employeeId.split('_')[0].toUpperCase()}`.slice(0, 60);
    try {
      await setDoc(doc(db, 'employees', employeeId), publicEmployeePayload({
        employeeId,
        employeeNumber,
        name,
        nameKey,
        normalized,
        agencyId,
        siteId,
      }), { merge: true });
      return { employeeId, employeeNumber };
    } catch (error) {
      lastError = error;
    }
  }

  throw lastError || new Error('A usable worker profile could not be created.');
}

async function savePunch(action) {
  const name = prettyName(document.getElementById('workerNameInput')?.value);
  const normalized = normalizeName(name);
  const siteId = selectedSite();
  const agencyId = String(document.getElementById('workerAgencySelect')?.value || '').trim();

  if (normalized.length < 2) throw new Error('Type your first and last name before punching.');
  if (!siteId) throw new Error('Choose OH01 or OHC before punching.');
  if (!VALID_AGENCIES.has(agencyId)) throw new Error('Choose Sterling, Excel, or Lifestyle Staffing before punching.');
  if (!VALID_ACTIONS.has(action)) throw new Error('That punch type is not valid.');

  const app = getApps().length ? getApp() : initializeApp(FIREBASE_CONFIG);
  const db = getFirestore(app);
  const existingMatch = await findExistingEmployee(db, name, siteId, agencyId).catch(() => null);

  let employeeId;
  let employeeNumber;
  if (existingMatch) {
    const employee = existingMatch.employee;
    employeeId = String(employee.employeeId || employee.id || '').trim();
    employeeNumber = String(employee.employeeNumber || employee.employeeID || employeeId).trim();
    if (existingMatch.assignAgency && employee.id) {
      await setDoc(doc(db, 'employees', employee.id), {
        agencyId,
        agencyAssignedAt: serverTimestamp(),
        agencyAssignmentSource: 'worker_public_selection',
        updatedAt: serverTimestamp(),
      }, { merge: true });
    }
  } else {
    ({ employeeId, employeeNumber } = await ensurePublicEmployee(db, name, normalized, siteId, agencyId));
  }

  if (!employeeId) throw new Error('The worker profile could not be verified. Ask a manager to check the employee record.');

  const now = new Date();
  const nowMs = Date.now();
  const duplicateKey = `stablePublicPunch:${employeeId}:${action}:${localDateKey(now)}`;
  const prior = Number(localStorage.getItem(duplicateKey) || 0);
  if (prior && nowMs - prior < 30000) throw new Error(`${ACTION_LABELS[action]} was already saved. No second tap is needed.`);

  await addDoc(collection(db, 'punches'), {
    companyId: COMPANY_ID,
    siteId,
    siteIds: [siteId],
    assignedSiteId: siteId,
    agencyId,
    employeeId,
    workerId: employeeId,
    employeeNumber,
    name,
    nameKey: normalized.replaceAll(' ', '_'),
    action,
    timestamp: serverTimestamp(),
    timestampMs: nowMs,
    dateKey: localDateKey(now),
    weekKey: mondayKey(now),
    source: 'public_qr',
    createdAt: serverTimestamp(),
    locationStatus: 'not_requested',
    enforceLocation: false,
    active: true,
    status: 'active',
  });

  localStorage.setItem(duplicateKey, String(nowMs));
  localStorage.setItem('workerPunchName', name);
  localStorage.setItem('workerPunchAgency', agencyId);

  if (document.getElementById('workerNameValue')) document.getElementById('workerNameValue').textContent = name;
  if (document.getElementById('workerLastActionValue')) document.getElementById('workerLastActionValue').textContent = ACTION_LABELS[action];
  if (document.getElementById('workerLastPunchValue')) document.getElementById('workerLastPunchValue').textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  setMessage(`${ACTION_LABELS[action]} saved for ${name}.`);
}

function install() {
  registerPunchWriter('stable-public-clock-handler', PUNCH_WRITER_PRIORITY.STABLE, async (_requestedAction, button) => {
    if (saving) return;
    saving = true;
    disableButtons(true);
    const action = String(button.dataset.action || '');
    const name = prettyName(document.getElementById('workerNameInput')?.value);
    setMessage(`Saving ${ACTION_LABELS[action] || 'punch'} for ${name || 'worker'}...`);

    try {
      await savePunch(action);
      window.setTimeout(() => window.location.reload(), 900);
    } catch (error) {
      console.error('[stable-public-clock-handler]', error);
      setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
      saving = false;
      disableButtons(false);
    }
  });
  console.info('[QRTimeclock] Returning-worker-safe public clock handler installed.');
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
else install();
