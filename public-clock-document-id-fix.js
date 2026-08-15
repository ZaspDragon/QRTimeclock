import { registerPunchWriter, PUNCH_WRITER_PRIORITY } from './punch-writer-lock.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { addDoc, collection, doc, getDoc, getFirestore, serverTimestamp, setDoc } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';
import {
  findPublicWorkerMatches,
  chooseCanonicalPublicWorker,
  employeeName,
  employeeSite,
  employeeAgency,
  isActiveWorker,
  normalizeWorkerName,
  workerNameKey,
} from './public-worker-lookup-v3.js';

// Single public punch writer.
// Important: Firestore publicPunchCreate() validates employees/{employeeId} by
// document path, so every new punch uses the employee document ID in employeeId.
// Canonical / legacy identity remains in workerId + canonicalEmployeeId.
// No historical punch or employee record is deleted or migrated here.

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
  return String(value || '').trim().replace(/\s+/g, ' ').toLowerCase().replace(/\b\w/g, (c) => c.toUpperCase());
}

function safePart(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '')
    .slice(0, 48) || 'worker';
}

function dateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function weekKey(date) {
  const monday = new Date(date);
  const day = monday.getDay();
  monday.setDate(monday.getDate() - (day === 0 ? 6 : day - 1));
  monday.setHours(0, 0, 0, 0);
  return dateKey(monday);
}

function selectedSite() {
  const querySite = String(new URLSearchParams(location.search).get('site') || '').trim().toUpperCase();
  if (VALID_SITES.has(querySite)) return querySite;
  const value = String(document.getElementById('workerBranchSelect')?.value || '').trim().toUpperCase();
  return VALID_SITES.has(value) ? value : '';
}

function selectedAgency() {
  const value = String(document.getElementById('workerAgencySelect')?.value || localStorage.getItem('workerPunchAgency') || '').trim();
  return VALID_AGENCIES.has(value) ? value : '';
}

function setMessage(message, isError = false) {
  const lookup = document.getElementById('workerLookupStatus');
  const status = document.getElementById('workerStatusMessage');
  const state = document.getElementById('workerStatusValue');
  if (lookup) {
    lookup.textContent = message;
    lookup.style.borderColor = isError ? 'rgba(255,90,90,.6)' : 'rgba(43,213,118,.5)';
  }
  if (status) status.textContent = message;
  if (state) state.textContent = isError ? 'Needs attention' : 'Saved';
}

function disableButtons(disabled) {
  document.querySelectorAll('.worker-action-btn').forEach((button) => {
    button.disabled = disabled;
    button.setAttribute('aria-busy', disabled ? 'true' : 'false');
  });
}

function installAgencyControl() {
  if (document.getElementById('workerAgencySelect')) return;
  const branch = document.getElementById('workerBranchSelect');
  const branchLabel = branch?.closest('label');
  if (!branchLabel) return;

  const label = document.createElement('label');
  label.id = 'workerAgencyField';
  label.innerHTML = `
    <span>Staffing agency <small>(required for every temp)</small></span>
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

async function rememberedWorker(db, name, siteId) {
  const rememberedId = String(localStorage.getItem('publicResolvedWorkerDocId') || localStorage.getItem('publicResolvedWorkerId') || '').trim();
  if (!rememberedId) return null;
  try {
    const snapshot = await getDoc(doc(db, 'employees', rememberedId));
    if (!snapshot.exists()) return null;
    const row = { id: snapshot.id, ...snapshot.data() };
    if (!isActiveWorker(row)) return null;
    if (normalizeWorkerName(employeeName(row)) !== normalizeWorkerName(name)) return null;
    const site = employeeSite(row);
    if (site && site !== siteId) return null;
    return row;
  } catch (_) {
    return null;
  }
}

async function assignBlankAgency(db, worker, selectedAgencyId) {
  const currentAgency = employeeAgency(worker);
  if (currentAgency) return { ...worker, agencyId: currentAgency };
  const docId = String(worker.id || '').trim();
  if (!docId) return worker;

  // This exact four-field update is explicitly allowed by the current
  // publicEmployeeAgencyAssignment Firestore rule for blank-agency workers.
  await setDoc(doc(db, 'employees', docId), {
    agencyId: selectedAgencyId,
    agencyAssignedAt: serverTimestamp(),
    agencyAssignmentSource: 'worker_public_selection',
    updatedAt: serverTimestamp(),
  }, { merge: true });
  return { ...worker, agencyId: selectedAgencyId };
}

async function createPublicWorker(db, name, siteId, agencyId) {
  const id = `public_canonical_${safePart(siteId)}_${safePart(workerNameKey(name))}`;
  const ref = doc(db, 'employees', id);
  const existing = await getDoc(ref).catch(() => null);
  if (existing?.exists()) {
    const row = { id: existing.id, ...existing.data() };
    if (!isActiveWorker(row)) throw new Error('An inactive worker profile already exists for this name. Ask a manager to reactivate it.');
    return assignBlankAgency(db, row, agencyId);
  }

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
    identityVersion: 4,
    updatedAt: serverTimestamp(),
    createdAt: serverTimestamp(),
  };
  await setDoc(ref, payload);
  return { id, ...payload };
}

async function resolveWorker(db, name, siteId, selectedAgencyId) {
  let worker = await rememberedWorker(db, name, siteId);

  if (!worker) {
    const matches = await findPublicWorkerMatches(name, siteId, selectedAgencyId);
    worker = chooseCanonicalPublicWorker(matches);
    if (!worker && matches.length > 1) {
      throw new Error('More than one separate worker uses that exact name. Ask a manager to select the correct profile.');
    }
  }

  if (!worker) worker = await createPublicWorker(db, name, siteId, selectedAgencyId);
  else worker = await assignBlankAgency(db, worker, selectedAgencyId);

  const employeeDocId = String(worker.id || '').trim();
  if (!employeeDocId) throw new Error('The worker profile document could not be verified.');

  // Existing nonblank agency is authoritative. This prevents a shared kiosk's
  // previously selected agency from silently rewriting another worker's profile.
  const punchAgencyId = employeeAgency(worker) || selectedAgencyId;
  if (!VALID_AGENCIES.has(punchAgencyId)) throw new Error('The worker staffing agency could not be verified.');

  localStorage.setItem('publicResolvedWorkerDocId', employeeDocId);
  localStorage.setItem('publicResolvedWorkerId', employeeDocId);
  localStorage.setItem('workerPunchAgency', punchAgencyId);

  return { worker, employeeDocId, punchAgencyId };
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
  const { worker, employeeDocId, punchAgencyId } = await resolveWorker(db, name, siteId, agencyId);

  const canonicalEmployeeId = String(
    worker.canonicalEmployeeId || worker.employeeId || worker.employeeID || worker.workerId || employeeDocId
  ).trim() || employeeDocId;
  const employeeNumber = String(worker.employeeNumber || worker.employeeNo || worker.employeeID || canonicalEmployeeId).trim();

  const now = new Date();
  const nowMs = Date.now();
  const guard = `publicClockDocId:${employeeDocId}:${action}:${dateKey(now)}`;
  const previous = Number(localStorage.getItem(guard) || 0);
  if (previous && nowMs - previous < 60_000) {
    throw new Error(`${LABELS[action]} was already saved. No second tap is needed.`);
  }

  await addDoc(collection(db, 'punches'), {
    companyId: COMPANY_ID,
    siteId,
    siteIds: [siteId],
    assignedSiteId: siteId,
    agencyId: punchAgencyId,
    // MUST be the Firestore employees/{documentId} path used by the security rule.
    employeeId: employeeDocId,
    // Preserve canonical/legacy identity separately for history/export matching.
    workerId: canonicalEmployeeId,
    canonicalEmployeeId,
    employeeNumber,
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
  if (document.getElementById('workerNameValue')) document.getElementById('workerNameValue').textContent = employeeName(worker) || name;
  if (document.getElementById('workerLastActionValue')) document.getElementById('workerLastActionValue').textContent = LABELS[action];
  if (document.getElementById('workerLastPunchValue')) document.getElementById('workerLastPunchValue').textContent = now.toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
  setMessage(`${LABELS[action]} saved for ${employeeName(worker) || name}.`);
}

function actionFromButton(button) {
  const dataAction = String(button.dataset.action || '').trim();
  if (LABELS[dataAction]) return dataAction;
  const text = String(button.textContent || '').trim().toLowerCase();
  if (text.includes('start lunch') || text.includes('lunch out')) return 'start_lunch';
  if (text.includes('end lunch') || text.includes('lunch in')) return 'end_lunch';
  if (text.includes('clock out')) return 'clock_out';
  if (text.includes('clock in')) return 'clock_in';
  return '';
}

function install() {
  if (document.documentElement.dataset.publicClockDocumentIdFix === 'true') return;
  document.documentElement.dataset.publicClockDocumentIdFix = 'true';

  registerPunchWriter(
    'public-clock-document-id-fix',
    PUNCH_WRITER_PRIORITY.DOCUMENT_ID_FIX,
    async (_action, button) => {
      if (saving) return;
      saving = true;
      disableButtons(true);
      const action = actionFromButton(button);
      const name = prettyName(document.getElementById('workerNameInput')?.value);
      setMessage(`Saving ${LABELS[action] || 'punch'} for ${name || 'worker'}...`);

      try {
        await savePunch(action);
        window.setTimeout(() => window.location.reload(), 900);
      } catch (error) {
        console.error('[public-clock-document-id-fix]', error);
        setMessage(error?.message || 'The punch could not be saved. Please try again.', true);
        saving = false;
        disableButtons(false);
      }
    },
    () => {
      const card = document.getElementById('workerCard');
      return Boolean(card) && !card.classList.contains('hidden');
    }
  );

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', installAgencyControl, { once: true });
  } else {
    installAgencyControl();
  }
}

install();
console.info('[QRTimeclock] Document-ID-safe public punch writer installed.');
