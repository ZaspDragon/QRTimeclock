import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  addDoc,
  collection,
  doc,
  getFirestore,
  serverTimestamp,
  setDoc,
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
const VALID_AGENCIES = new Set([
  'sterling_staffing',
  'excel_staffing',
  'lifestyle_staffing',
]);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ACTION_LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Lunch Out',
  end_lunch: 'Lunch In',
  clock_out: 'Clock Out',
};

let saving = false;

function prettyName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .replace(/\b\w/g, (character) => character.toUpperCase());
}

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
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
  const nameKey = normalized.replaceAll(' ', '_');
  const employeeId = `public_${safeIdPart(siteId)}_${safeIdPart(agencyId)}_${safeIdPart(nameKey)}`;
  const employeeNumber = `PUBLIC-${safeIdPart(siteId).toUpperCase()}-${safeIdPart(agencyId).toUpperCase()}-${safeIdPart(nameKey).toUpperCase()}`.slice(0, 60);
  const now = new Date();
  const nowMs = Date.now();
  const duplicateKey = `stablePublicPunch:${employeeId}:${action}:${localDateKey(now)}`;
  const prior = Number(localStorage.getItem(duplicateKey) || 0);

  if (prior && nowMs - prior < 30000) {
    throw new Error(`${ACTION_LABELS[action]} was already saved. No second tap is needed.`);
  }

  await setDoc(doc(db, 'employees', employeeId), {
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
  }, { merge: true });

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
    nameKey,
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

  const nameValue = document.getElementById('workerNameValue');
  const actionValue = document.getElementById('workerLastActionValue');
  const punchValue = document.getElementById('workerLastPunchValue');
  if (nameValue) nameValue.textContent = name;
  if (actionValue) actionValue.textContent = ACTION_LABELS[action];
  if (punchValue) punchValue.textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  setMessage(`${ACTION_LABELS[action]} saved for ${name}.`);
}

function install() {
  document.addEventListener('click', async (event) => {
    const button = event.target.closest('.worker-action-btn');
    if (!button || saving) return;

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
      console.error('[stable-public-clock-handler]', error);
      setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
      saving = false;
      disableButtons(false);
    }
  }, true);

  console.info('[QRTimeclock] Stable public clock handler installed.');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', install, { once: true });
} else {
  install();
}
