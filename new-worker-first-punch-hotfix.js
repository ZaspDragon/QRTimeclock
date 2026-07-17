import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { addDoc, collection, doc, getFirestore, serverTimestamp, setDoc } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const db = getFirestore(app);
const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
let saving = false;

const prettyName = (value) => String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, (char) => char.toUpperCase());
const nameKey = (value) => String(value || '').trim().toLowerCase().replace(/\s+/g, ' ').replace(/[^a-z0-9 ]/g, '').replaceAll(' ', '_');
const safeIdPart = (value) => String(value || '').toLowerCase().replace(/[^a-z0-9_-]/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '').slice(0, 48) || 'worker';

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

function isNewWorkerScreen(name) {
  const status = document.getElementById('workerLookupStatus')?.textContent || '';
  return /new worker/i.test(status) && status.toLowerCase().includes(name.toLowerCase());
}

function setMessage(message, isError = false) {
  const lookup = document.getElementById('workerLookupStatus');
  const status = document.getElementById('workerStatusMessage');
  if (lookup) {
    lookup.textContent = message;
    lookup.style.borderColor = isError ? 'rgba(255,90,90,0.6)' : 'rgba(43,213,118,0.5)';
  }
  if (status) status.textContent = message;
}

function setButtonsDisabled(disabled) {
  document.querySelectorAll('.worker-action-btn').forEach((button) => { button.disabled = disabled; });
}

async function saveFirstPunch(action) {
  const name = prettyName(document.getElementById('workerNameInput')?.value);
  const selectedSite = document.getElementById('workerBranchSelect')?.value;
  const siteId = VALID_SITES.has(selectedSite) ? selectedSite : 'OH01';
  if (!name || name.length < 2) throw new Error('Type your first and last name before clocking in.');
  if (!VALID_ACTIONS.has(action)) throw new Error('That punch type is not valid.');

  const normalizedName = nameKey(name);
  const employeeId = `auto_${safeIdPart(siteId)}_${safeIdPart(normalizedName)}`;
  const now = new Date();
  const nowMs = Date.now();
  const duplicateKey = `firstPunchHotfix:${employeeId}:${action}`;
  const previous = Number(localStorage.getItem(duplicateKey) || 0);
  if (previous && nowMs - previous < 15000) throw new Error('That punch was already saved. Please wait a moment before trying again.');

  const employeeNumber = `AUTO-${safeIdPart(siteId).toUpperCase()}-${safeIdPart(normalizedName).toUpperCase()}`.slice(0, 60);
  await setDoc(doc(db, 'employees', employeeId), {
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
  }, { merge: true });

  await addDoc(collection(db, 'punches'), {
    companyId: COMPANY_ID,
    siteId,
    siteIds: [siteId],
    assignedSiteId: siteId,
    agencyId: '',
    employeeId,
    workerId: employeeId,
    employeeNumber,
    name,
    nameKey: normalizedName,
    action,
    timestamp: serverTimestamp(),
    timestampMs: nowMs,
    dateKey: localDateKey(now),
    weekKey: mondayKey(now),
    source: 'public_qr',
    createdAt: serverTimestamp(),
    locationStatus: 'not_requested',
    enforceLocation: false,
  });

  localStorage.setItem(duplicateKey, String(nowMs));
  localStorage.setItem('workerPunchName', name);
  const actionLabel = { clock_in: 'Clock In', start_lunch: 'Lunch Out', end_lunch: 'Lunch In', clock_out: 'Clock Out' }[action] || 'Punch';
  setMessage(`${actionLabel} saved for ${name}. Your employee profile is now active.`);
  window.setTimeout(() => window.location.reload(), 1200);
}

document.addEventListener('click', async (event) => {
  const button = event.target.closest('.worker-action-btn');
  if (!button || saving) return;
  const name = prettyName(document.getElementById('workerNameInput')?.value);
  if (!name || !isNewWorkerScreen(name)) return;

  event.preventDefault();
  event.stopImmediatePropagation();
  saving = true;
  setButtonsDisabled(true);
  setMessage(`Creating ${name} and saving the first punch...`);
  try {
    await saveFirstPunch(button.dataset.action);
  } catch (error) {
    console.error('[new-worker-first-punch-hotfix]', error);
    setMessage(error?.message || 'The first punch could not be saved. Please try again.', true);
    saving = false;
    setButtonsDisabled(false);
  }
}, true);
