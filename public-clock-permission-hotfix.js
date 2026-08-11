import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { addDoc, collection, doc, getDoc, getFirestore, serverTimestamp, setDoc } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';
import { findPublicWorkerMatches, chooseCanonicalPublicWorker, employeeName, employeeSite, employeeAgency, isActiveWorker, normalizeWorkerName, workerNameKey } from './public-worker-lookup-v3.js';

// Public clock writer aligned with Firestore rules.
// Every worker keeps access to Clock In, Start Lunch, End Lunch and Clock Out.
// No punch history is deleted or migrated here.

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
const LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Start Lunch',
  end_lunch: 'End Lunch',
  clock_out: 'Clock Out',
};
let saving = false;

function dbInstance() {
  const app = getApps().length ? getApp() : initializeApp(FIREBASE_CONFIG);
  return getFirestore(app);
}
function prettyName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, c => c.toUpperCase());
}
function safePart(value) {
  return String(value || '').toLowerCase().replace(/[^a-z0-9_-]/g, '_').replace(/_+/g, '_').replace(/^_+|_+$/g, '').slice(0, 48) || 'worker';
}
function dateKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}
function weekKey(d) {
  const monday = new Date(d);
  const day = monday.getDay();
  monday.setDate(monday.getDate() - (day === 0 ? 6 : day - 1));
  monday.setHours(0, 0, 0, 0);
  return dateKey(monday);
}
function selectedSite() {
  const qs = String(new URLSearchParams(location.search).get('site') || '').toUpperCase();
  if (VALID_SITES.has(qs)) return qs;
  const value = String(document.getElementById('workerBranchSelect')?.value || '').toUpperCase();
  return VALID_SITES.has(value) ? value : '';
}
function selectedAgency() {
  const value = String(document.getElementById('workerAgencySelect')?.value || localStorage.getItem('workerPunchAgency') || '').trim();
  return VALID_AGENCIES.has(value) ? value : '';
}
function setMessage(message, error = false) {
  const lookup = document.getElementById('workerLookupStatus');
  const status = document.getElementById('workerStatusMessage');
  const state = document.getElementById('workerStatusValue');
  if (lookup) {
    lookup.textContent = message;
    lookup.style.borderColor = error ? 'rgba(255,90,90,.6)' : 'rgba(43,213,118,.5)';
  }
  if (status) status.textContent = message;
  if (state) state.textContent = error ? 'Needs attention' : 'Saved';
}
function disableButtons(disabled) {
  document.querySelectorAll('.worker-action-btn').forEach(btn => {
    btn.disabled = disabled;
    btn.setAttribute('aria-busy', disabled ? 'true' : 'false');
  });
}
function installAgencyControl() {
  if (document.getElementById('workerAgencySelect')) return;
  const branch = document.getElementById('workerBranchSelect');
  const branchLabel = branch?.closest('label');
  if (!branchLabel) return;
  const label = document.createElement('label');
  label.id = 'workerAgencyField';
  label.innerHTML = `<span>Staffing agency <small>(required for every temp)</small></span><select id="workerAgencySelect" required aria-required="true"><option value="">Choose your staffing agency</option>${[...VALID_AGENCIES.entries()].map(([v,t]) => `<option value="${v}">${t}</option>`).join('')}</select>`;
  branchLabel.insertAdjacentElement('afterend', label);
  const select = label.querySelector('select');
  const saved = String(localStorage.getItem('workerPunchAgency') || '');
  if (VALID_AGENCIES.has(saved)) select.value = saved;
  select.addEventListener('change', () => {
    if (VALID_AGENCIES.has(select.value)) localStorage.setItem('workerPunchAgency', select.value);
  });
}

async function loadResolvedWorker(db, name, siteId, agencyId) {
  const rememberedId = String(localStorage.getItem('publicResolvedWorkerId') || '').trim();
  if (!rememberedId) return null;
  try {
    const snap = await getDoc(doc(db, 'employees', rememberedId));
    if (!snap.exists()) return null;
    const row = { id: snap.id, ...snap.data() };
    const site = employeeSite(row);
    if (!isActiveWorker(row)) return null;
    if (normalizeWorkerName(employeeName(row)) !== normalizeWorkerName(name)) return null;
    if (site && site !== siteId) return null;
    if (employeeAgency(row) && employeeAgency(row) !== agencyId) return null;
    return row;
  } catch (_) {
    return null;
  }
}

async function ensureWorker(db, name, siteId, agencyId) {
  let worker = await loadResolvedWorker(db, name, siteId, agencyId);
  if (!worker) {
    const matches = await findPublicWorkerMatches(name, siteId, agencyId);
    worker = chooseCanonicalPublicWorker(matches);
    if (!worker && matches.length > 1) {
      throw new Error('More than one separate worker uses that exact name. Ask a manager to select the correct profile before punching.');
    }
  }

  if (!worker) {
    const id = `public_canonical_${safePart(siteId)}_${safePart(workerNameKey(name))}`;
    const ref = doc(db, 'employees', id);
    const existing = await getDoc(ref).catch(() => null);
    if (existing?.exists()) {
      worker = { id: existing.id, ...existing.data() };
    } else {
      const employeeNumber = `PUBLIC-${safePart(siteId).toUpperCase()}-${safePart(workerNameKey(name)).toUpperCase()}`.slice(0, 60);
      const payload = {
        name,
        nameKey: workerNameKey(name),
        normalizedName: workerNameKey(name),
        employeeNumber,
        employeeNumberKey: employeeNumber.toLowerCase(),
        companyId: COMPANY_ID,
        agencyId,
        assignedSiteId: siteId,
        siteId,
        siteIds: [siteId],
        status: 'active',
        active: true,
        employeeId: id,
        canonicalEmployeeId: id,
        source: 'auto_created',
        identityVersion: 3,
        updatedAt: serverTimestamp(),
        createdAt: serverTimestamp(),
      };
      await setDoc(ref, payload);
      worker = { id, ...payload };
    }
  }

  const id = String(worker.employeeId || worker.id || '').trim();
  if (!id) throw new Error('Worker ID could not be verified.');

  // Keep current Firestore-compatible profile normalization so public punches are
  // accepted. Historical punches are not touched; only the active worker profile
  // may receive missing/current agency metadata.
  if (employeeAgency(worker) !== agencyId || !worker.nameKey || !worker.employeeNumber || worker.source !== 'auto_created') {
    const employeeNumber = String(worker.employeeNumber || worker.employeeID || worker.employeeId || id);
    const patch = {
      name: employeeName(worker) || name,
      nameKey: workerNameKey(employeeName(worker) || name),
      employeeNumber,
      companyId: String(worker.companyId || COMPANY_ID),
      siteId: employeeSite(worker) || siteId,
      agencyId,
      status: 'active',
      active: true,
      source: 'auto_created',
      canonicalEmployeeId: String(worker.canonicalEmployeeId || id),
      updatedAt: serverTimestamp(),
    };
    await setDoc(doc(db, 'employees', worker.id || id), patch, { merge: true });
    worker = { ...worker, ...patch };
  }

  localStorage.setItem('publicResolvedWorkerId', String(worker.id || worker.employeeId || id));
  return worker;
}

async function savePunch(action) {
  const name = prettyName(document.getElementById('workerNameInput')?.value);
  const siteId = selectedSite();
  const agencyId = selectedAgency();
  if (normalizeWorkerName(name).length < 2) throw new Error('Type your first and last name before punching.');
  if (!siteId) throw new Error('Choose OH01 or OHC before punching.');
  if (!agencyId) throw new Error('Choose your staffing agency before punching.');
  if (!LABELS[action]) throw new Error('That punch type is not valid.');

  const db = dbInstance();
  const worker = await ensureWorker(db, name, siteId, agencyId);
  const employeeId = String(worker.employeeId || worker.id || '').trim();
  const now = new Date();
  const nowMs = Date.now();
  const guard = `publicClockHotfix:${employeeId}:${action}:${dateKey(now)}`;
  const previous = Number(localStorage.getItem(guard) || 0);
  if (previous && nowMs - previous < 60000) throw new Error(`${LABELS[action]} was already saved. No second tap is needed.`);

  await addDoc(collection(db, 'punches'), {
    companyId: COMPANY_ID,
    siteId,
    siteIds: [siteId],
    assignedSiteId: siteId,
    agencyId,
    employeeId,
    workerId: String(worker.canonicalEmployeeId || worker.workerId || employeeId),
    canonicalEmployeeId: String(worker.canonicalEmployeeId || employeeId),
    employeeNumber: String(worker.employeeNumber || employeeId),
    name: employeeName(worker) || name,
    nameKey: workerNameKey(employeeName(worker) || name),
    action,
    timestamp: serverTimestamp(),
    timestampMs: nowMs,
    dateKey: dateKey(now),
    weekKey: weekKey(now),
    source: 'public_qr',
    createdAt: serverTimestamp(),
    locationStatus: 'not_requested',
    enforceLocation: false,
    active: true,
    status: 'active',
  });

  localStorage.setItem(guard, String(nowMs));
  localStorage.setItem('workerPunchName', employeeName(worker) || name);
  localStorage.setItem('workerPunchAgency', agencyId);
  if (document.getElementById('workerNameValue')) document.getElementById('workerNameValue').textContent = employeeName(worker) || name;
  if (document.getElementById('workerLastActionValue')) document.getElementById('workerLastActionValue').textContent = LABELS[action];
  if (document.getElementById('workerLastPunchValue')) document.getElementById('workerLastPunchValue').textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  setMessage(`${LABELS[action]} saved for ${employeeName(worker) || name}.`);
}

function actionFromButton(button) {
  const fromData = String(button.dataset.action || '').trim();
  if (LABELS[fromData]) return fromData;
  const text = String(button.textContent || '').trim().toLowerCase();
  if (text.includes('start lunch') || text.includes('lunch out')) return 'start_lunch';
  if (text.includes('end lunch') || text.includes('lunch in')) return 'end_lunch';
  if (text.includes('clock out')) return 'clock_out';
  if (text.includes('clock in')) return 'clock_in';
  return '';
}

function install() {
  if (document.documentElement.dataset.publicClockPermissionHotfixV2 === 'true') return;
  document.documentElement.dataset.publicClockPermissionHotfixV2 = 'true';
  document.addEventListener('click', async event => {
    const button = event.target.closest?.('.worker-action-btn');
    const card = document.getElementById('workerCard');
    if (!button || !card || card.classList.contains('hidden') || saving) return;
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    saving = true;
    disableButtons(true);
    const action = actionFromButton(button);
    const name = prettyName(document.getElementById('workerNameInput')?.value);
    setMessage(`Saving ${LABELS[action] || 'punch'} for ${name || 'worker'}...`);
    try {
      await savePunch(action);
      window.setTimeout(() => window.location.reload(), 900);
    } catch (error) {
      console.error('[public-clock-permission-hotfix-v2]', error);
      setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
      saving = false;
      disableButtons(false);
    }
  }, true);
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', installAgencyControl, { once: true });
  else installAgencyControl();
}

install();
console.info('[QRTimeclock] Rules-safe public clock writer installed.');
