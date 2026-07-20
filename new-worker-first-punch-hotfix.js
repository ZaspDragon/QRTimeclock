import { firebaseConfig } from './firebase-config.js';
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

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const db = getFirestore(app);
const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ACTION_LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Start Lunch',
  end_lunch: 'End Lunch',
  clock_out: 'Clock Out',
};
let saving = false;

const prettyName = (value) => String(value || '')
  .trim()
  .replace(/\s+/g, ' ')
  .toLowerCase()
  .replace(/\b\w/g, (char) => char.toUpperCase());

const nameKey = (value) => String(value || '')
  .trim()
  .toLowerCase()
  .replace(/\s+/g, ' ')
  .replace(/[^a-z0-9 ]/g, '')
  .replaceAll(' ', '_');

const safeIdPart = (value) => String(value || '')
  .toLowerCase()
  .replace(/[^a-z0-9_-]/g, '_')
  .replace(/_+/g, '_')
  .replace(/^_+|_+$/g, '')
  .slice(0, 48) || 'worker';

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

function shouldUseOneClickFallback() {
  const status = String(document.getElementById('workerLookupStatus')?.textContent || '').toLowerCase();
  return !status
    || status.includes('type your name')
    || status.includes('new worker')
    || status.includes('not found')
    || status.includes('could not')
    || status.includes('access denied')
    || status.includes('check employee setup');
}

function setMessage(message, isError = false) {
  const lookup = document.getElementById('workerLookupStatus');
  const status = document.getElementById('workerStatusMessage');
  const stateValue = document.getElementById('workerStatusValue');
  if (lookup) {
    lookup.textContent = message;
    lookup.style.borderColor = isError ? 'rgba(255,90,90,0.6)' : 'rgba(43,213,118,0.5)';
  }
  if (status) status.textContent = message;
  if (stateValue) stateValue.textContent = isError ? 'Needs attention' : 'Saved';
}

function setButtonsDisabled(disabled) {
  document.querySelectorAll('.worker-action-btn').forEach((button) => {
    button.disabled = disabled;
  });
}

async function findExistingEmployee(normalizedName, siteId) {
  const searches = [
    query(
      collection(db, 'employees'),
      where('companyId', '==', COMPANY_ID),
      where('siteId', '==', siteId),
      where('active', '==', true),
      where('nameKey', '==', normalizedName),
      limit(2),
    ),
    query(
      collection(db, 'employees'),
      where('companyId', '==', COMPANY_ID),
      where('siteId', '==', siteId),
      where('active', '==', true),
      where('normalizedName', '==', normalizedName),
      limit(2),
    ),
  ];

  for (const employeeQuery of searches) {
    try {
      const snapshot = await getDocs(employeeQuery);
      if (!snapshot.empty) {
        const employeeDoc = snapshot.docs[0];
        return { id: employeeDoc.id, ...employeeDoc.data() };
      }
    } catch (error) {
      console.warn('[one-click-punch] Employee lookup attempt failed:', error?.message || error);
    }
  }

  return null;
}

async function resolveEmployee(name, siteId) {
  const normalizedName = nameKey(name);
  const existing = await findExistingEmployee(normalizedName, siteId);
  if (existing) {
    return {
      ...existing,
      employeeId: existing.employeeId || existing.employeeID || existing.id,
      employeeNumber: existing.employeeNumber || existing.employeeID || existing.employeeId || existing.id,
      name: existing.name || existing.employeeName || name,
      nameKey: existing.nameKey || normalizedName,
      agencyId: existing.agencyId || '',
    };
  }

  const employeeId = `auto_${safeIdPart(siteId)}_${safeIdPart(normalizedName)}`;
  const employeeNumber = `AUTO-${safeIdPart(siteId).toUpperCase()}-${safeIdPart(normalizedName).toUpperCase()}`.slice(0, 60);
  const employee = {
    name,
    nameKey: normalizedName,
    normalizedName,
    employeeNumber,
    employeeNumberKey: employeeNumber.toLowerCase(),
    companyId: COMPANY_ID,
    agencyId: '',
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

  await setDoc(doc(db, 'employees', employeeId), employee, { merge: true });
  return employee;
}

async function saveOneClickPunch(action) {
  const name = prettyName(document.getElementById('workerNameInput')?.value);
  const selectedSite = document.getElementById('workerBranchSelect')?.value;
  const siteId = VALID_SITES.has(selectedSite) ? selectedSite : 'OH01';

  if (!name || name.length < 2) {
    throw new Error('Type your name before clocking in.');
  }
  if (!VALID_ACTIONS.has(action)) {
    throw new Error('That punch type is not valid.');
  }

  const employee = await resolveEmployee(name, siteId);
  const now = new Date();
  const nowMs = Date.now();
  const employeeId = employee.employeeId || employee.id;
  const duplicateKey = `oneClickPunch:${employeeId}:${action}:${localDateKey(now)}`;
  const previous = Number(localStorage.getItem(duplicateKey) || 0);

  if (previous && nowMs - previous < 30000) {
    throw new Error(`${ACTION_LABELS[action]} was already saved. No second tap is needed.`);
  }

  await addDoc(collection(db, 'punches'), {
    companyId: COMPANY_ID,
    siteId,
    siteIds: Array.isArray(employee.siteIds) && employee.siteIds.length ? employee.siteIds : [siteId],
    assignedSiteId: employee.assignedSiteId || siteId,
    agencyId: employee.agencyId || '',
    employeeId,
    workerId: employee.workerId || employeeId,
    employeeNumber: employee.employeeNumber || employeeId,
    name: employee.name || name,
    nameKey: employee.nameKey || nameKey(name),
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
  localStorage.setItem('workerPunchName', employee.name || name);

  const lastAction = document.getElementById('workerLastActionValue');
  const lastPunch = document.getElementById('workerLastPunchValue');
  const enteredName = document.getElementById('workerNameValue');
  if (lastAction) lastAction.textContent = ACTION_LABELS[action];
  if (lastPunch) lastPunch.textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  if (enteredName) enteredName.textContent = employee.name || name;

  setMessage(`${ACTION_LABELS[action]} saved for ${employee.name || name}.`);
  window.setTimeout(() => window.location.reload(), 900);
}

document.addEventListener('click', async (event) => {
  const button = event.target.closest('.worker-action-btn');
  if (!button || saving || !shouldUseOneClickFallback()) return;

  event.preventDefault();
  event.stopImmediatePropagation();
  saving = true;
  setButtonsDisabled(true);

  const name = prettyName(document.getElementById('workerNameInput')?.value);
  setMessage(`Saving ${ACTION_LABELS[button.dataset.action] || 'punch'} for ${name || 'worker'}...`);

  try {
    await saveOneClickPunch(button.dataset.action);
  } catch (error) {
    console.error('[one-click-punch]', error);
    setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
    saving = false;
    setButtonsDisabled(false);
  }
}, true);
