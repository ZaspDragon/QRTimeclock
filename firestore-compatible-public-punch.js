import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  addDoc,
  collection,
  doc,
  getFirestore,
  serverTimestamp,
  setDoc,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const db = getFirestore(app);
const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_AGENCIES = new Set(['sterling_staffing', 'excel_staffing', 'lifestyle_staffing']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ACTION_LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Lunch Out',
  end_lunch: 'Lunch In',
  clock_out: 'Clock Out',
};
let saving = false;

function normalizeName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/[^a-z0-9 ]/g, '').trim();
}

function prettyName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, (character) => character.toUpperCase());
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

function unresolvedWorkerStatus() {
  const message = String(document.getElementById('workerLookupStatus')?.textContent || '').toLowerCase();
  return message.includes('no existing worker record was found')
    || message.includes('no existing worker was found')
    || message.includes('employee directory could not be read')
    || message.includes('could not check that name')
    || message.includes('check employee setup');
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

function setButtonsDisabled(disabled) {
  document.querySelectorAll('.worker-action-btn').forEach((button) => { button.disabled = disabled; });
}

async function saveCompatiblePunch(action) {
  const name = prettyName(document.getElementById('workerNameInput')?.value);
  const normalizedName = normalizeName(name);
  const selectedSite = String(document.getElementById('workerBranchSelect')?.value || '').toUpperCase();
  const siteId = VALID_SITES.has(selectedSite) ? selectedSite : '';
  const agencyId = String(document.getElementById('workerAgencySelect')?.value || '').trim();

  if (!name || normalizedName.length < 2) throw new Error('Type your first and last name before punching.');
  if (!siteId) throw new Error('The QR-code branch could not be confirmed. Scan the correct OH01 or OHC code again.');
  if (!VALID_AGENCIES.has(agencyId)) throw new Error('Choose Sterling, Excel, or Lifestyle Staffing before punching.');
  if (!VALID_ACTIONS.has(action)) throw new Error('That punch type is not valid.');

  const employeeId = `public_compat_${safeIdPart(siteId)}_${safeIdPart(agencyId)}_${safeIdPart(normalizedName)}`;
  const employeeNumber = `PUBLIC-${safeIdPart(siteId).toUpperCase()}-${safeIdPart(agencyId).toUpperCase()}-${safeIdPart(normalizedName).toUpperCase()}`.slice(0, 60);
  const nameKey = normalizedName.replaceAll(' ', '_');

  await setDoc(doc(db, 'employees', employeeId), {
    name,
    nameKey,
    normalizedName,
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

  const now = new Date();
  const nowMs = Date.now();
  const duplicateKey = `compatiblePunch:${employeeId}:${action}:${localDateKey(now)}`;
  const previous = Number(localStorage.getItem(duplicateKey) || 0);
  if (previous && nowMs - previous < 30000) throw new Error(`${ACTION_LABELS[action]} was already saved. No second tap is needed.`);

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
  setMessage(`${ACTION_LABELS[action]} saved for ${name}.`);
  window.setTimeout(() => window.location.reload(), 900);
}

function install() {
  document.addEventListener('click', async (event) => {
    const button = event.target.closest('.worker-action-btn');
    if (!button || saving || !unresolvedWorkerStatus()) return;

    event.preventDefault();
    event.stopImmediatePropagation();
    saving = true;
    setButtonsDisabled(true);

    const action = String(button.dataset.action || '');
    const name = prettyName(document.getElementById('workerNameInput')?.value);
    setMessage(`Saving ${ACTION_LABELS[action] || 'punch'} for ${name || 'worker'}...`);

    try {
      await saveCompatiblePunch(action);
    } catch (error) {
      console.error('[firestore-compatible-public-punch]', error);
      setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
      saving = false;
      setButtonsDisabled(false);
    }
  }, true);

  console.info('[QRTimeclock] Firestore-compatible public punch fallback installed.');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', install, { once: true });
} else {
  install();
}
