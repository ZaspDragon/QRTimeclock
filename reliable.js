import { firebaseConfig } from './firebase-config.js';
import { initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  getAuth, signInWithEmailAndPassword, signOut, onAuthStateChanged
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  getFirestore, collection, doc, getDoc, getDocs, query, where,
  runTransaction, serverTimestamp, addDoc
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const COMPANY_ID = 'chadwell';
const OWNER_EMAIL = 'brandon.evanshine@chadwellsupply.com';
const ACTIONS = ['clock_in', 'start_lunch', 'end_lunch', 'clock_out'];
const LABELS = {
  clock_in: 'Clock In',
  start_lunch: 'Lunch Out',
  end_lunch: 'Lunch In',
  clock_out: 'Clock Out'
};
const DEFAULT_SCHEDULE = {
  clock_in: '07:00',
  start_lunch: '11:00',
  end_lunch: '11:30',
  clock_out: '15:30'
};

const el = (id) => document.getElementById(id);
const state = { user: null, profile: null, employees: [], dayPunches: [] };

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function normalizeRole(value) {
  return String(value || '').trim().toLowerCase().replace(/[ -]+/g, '_');
}
function dateKey(date = new Date()) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, '0');
  const d = String(date.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
}
function mondayKey(date) {
  const copy = new Date(date);
  const day = copy.getDay();
  copy.setDate(copy.getDate() - (day === 0 ? 6 : day - 1));
  return dateKey(copy);
}
function formatTime(ms) {
  return new Date(Number(ms)).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}
function setStatus(target, text, type = '') {
  target.textContent = text;
  target.className = `notice ${type}`.trim();
}
function activeRecord(row) {
  return row?.active !== false && !['deleted', 'inactive', 'removed', 'terminated', 'merged'].includes(String(row?.status || '').toLowerCase());
}
async function sha256(text) {
  const bytes = new TextEncoder().encode(text);
  const digest = await crypto.subtle.digest('SHA-256', bytes);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, '0')).join('');
}
async function stableId(prefix, parts) {
  return `${prefix}_${(await sha256(parts.join('|'))).slice(0, 40)}`;
}
async function findEmployeeByName(name, siteId) {
  const snap = await getDocs(query(
    collection(db, 'employees'),
    where('companyId', '==', COMPANY_ID),
    where('siteId', '==', siteId)
  ));
  const matches = snap.docs
    .map((d) => ({ id: d.id, ...d.data() }))
    .filter(activeRecord)
    .filter((row) => normalizeName(row.name || row.nameKey) === normalizeName(name));
  if (matches.length !== 1) {
    throw new Error(matches.length ? 'More than one employee has that name. Ask a manager to clean up the duplicate.' : 'Employee not found or inactive. Ask a manager to add or activate the employee.');
  }
  return matches[0];
}

async function saveFirstPunch(employee, action, siteId) {
  const now = new Date();
  const nowMs = Date.now();
  const day = dateKey(now);
  const employeeId = employee.employeeId || employee.id;
  const punchId = await stableId('first', [COMPANY_ID, siteId, employeeId, day, action]);
  const stateId = await stableId('public', [COMPANY_ID, siteId, employee.agencyId || '', employeeId]);
  const punchRef = doc(db, 'punches', punchId);
  const guardRef = doc(db, 'punchGuards', punchId);
  const stateRef = doc(db, 'punchStates', stateId);

  return runTransaction(db, async (tx) => {
    const [punchSnap, guardSnap] = await Promise.all([tx.get(punchRef), tx.get(guardRef)]);
    if (punchSnap.exists() && activeRecord(punchSnap.data())) {
      return { created: false, timestampMs: Number(punchSnap.data().timestampMs || nowMs) };
    }
    if (guardSnap.exists()) {
      return { created: false, timestampMs: Number(guardSnap.data().acceptedAtMs || nowMs) };
    }

    const payload = {
      companyId: COMPANY_ID,
      siteId,
      siteIds: [siteId],
      assignedSiteId: siteId,
      employeeId,
      workerId: employee.workerId || employee.id || employeeId,
      employeeNumber: employee.employeeNumber || '',
      agencyId: employee.agencyId || '',
      name: employee.name,
      nameKey: normalizeName(employee.name),
      action,
      timestamp: serverTimestamp(),
      timestampMs: nowMs,
      dateKey: day,
      weekKey: mondayKey(now),
      source: 'public_qr',
      createdAt: serverTimestamp(),
      duplicateGuardKey: punchId,
      idempotencyKey: punchId,
      workerStateKey: stateId,
      active: true,
      status: 'active'
    };
    tx.set(punchRef, payload);
    tx.set(guardRef, {
      duplicateGuardKey: punchId,
      punchId,
      companyId: COMPANY_ID,
      siteId,
      employeeId,
      agencyId: employee.agencyId || '',
      action,
      acceptedAtMs: nowMs,
      expiresAtMs: nowMs + 86400000,
      createdAt: serverTimestamp()
    });
    tx.set(stateRef, {
      workerStateKey: stateId,
      companyId: COMPANY_ID,
      siteId,
      employeeId,
      agencyId: employee.agencyId || '',
      lastAction: action,
      lastPunchAtMs: nowMs,
      lastPunchId: punchId,
      updatedAt: serverTimestamp()
    }, { merge: true });
    return { created: true, timestampMs: nowMs };
  });
}

async function handleWorkerPunch(action) {
  const status = el('workerStatus');
  const buttons = [...document.querySelectorAll('.punch-btn')];
  try {
    const name = el('workerName').value.trim();
    const siteId = el('workerSite').value;
    if (name.length < 2) throw new Error('Enter your first and last name.');
    buttons.forEach((b) => { b.disabled = true; });
    setStatus(status, `Saving ${LABELS[action]}…`);
    const employee = await findEmployeeByName(name, siteId);
    const result = await saveFirstPunch(employee, action, siteId);
    setStatus(status, result.created
      ? `${LABELS[action]} saved at ${formatTime(result.timestampMs)}.`
      : `Your first ${LABELS[action]} was already saved at ${formatTime(result.timestampMs)}.`, 'success');
  } catch (error) {
    console.error('[Reliable punch]', { action, result: 'error', message: error.message });
    setStatus(status, error.message || 'Punch could not be saved. Tap again or contact a manager.', 'error');
  } finally {
    buttons.forEach((b) => { b.disabled = false; });
  }
}

document.querySelectorAll('.punch-btn').forEach((button) => {
  button.addEventListener('click', () => handleWorkerPunch(button.dataset.action));
});

function protectedOwner() {
  return String(state.user?.email || state.profile?.email || '').toLowerCase() === OWNER_EMAIL;
}
function canEdit() {
  const role = normalizeRole(state.profile?.role);
  return protectedOwner()
    || ['owner', 'superadmin', 'super_admin', 'admin', 'manager', 'supervisor'].includes(role)
    || state.profile?.permissions?.canEditPunches === true;
}
function allowedSites() {
  if (protectedOwner()) return ['OH01', 'OHC'];
  const raw = state.profile?.branches || state.profile?.siteIds || [state.profile?.branch || state.profile?.siteId];
  return [...new Set((Array.isArray(raw) ? raw : [raw]).filter((x) => ['OH01', 'OHC'].includes(x)))];
}

async function loadEmployees() {
  const siteId = el('managerSite').value;
  const snap = await getDocs(query(
    collection(db, 'employees'),
    where('companyId', '==', COMPANY_ID),
    where('siteId', '==', siteId)
  ));
  state.employees = snap.docs.map((d) => ({ id: d.id, ...d.data() })).filter(activeRecord).sort((a, b) => String(a.name).localeCompare(String(b.name)));
  el('employeeSelect').innerHTML = state.employees.map((row) => `<option value="${row.id}">${row.name}</option>`).join('');
}

async function loadDay() {
  if (!canEdit()) throw new Error('Your account does not have permission to edit punches.');
  const employee = state.employees.find((row) => row.id === el('employeeSelect').value);
  const siteId = el('managerSite').value;
  const day = el('workDate').value;
  if (!employee || !day) throw new Error('Choose an employee and date.');
  const employeeId = employee.employeeId || employee.id;
  const snap = await getDocs(query(
    collection(db, 'punches'),
    where('companyId', '==', COMPANY_ID),
    where('siteId', '==', siteId),
    where('employeeId', '==', employeeId),
    where('dateKey', '==', day)
  ));
  state.dayPunches = snap.docs.map((d) => ({ id: d.id, ...d.data() })).filter(activeRecord);
  renderSlots(employee);
}

function earliestFor(action) {
  return state.dayPunches.filter((row) => row.action === action).sort((a, b) => Number(a.timestampMs) - Number(b.timestampMs))[0] || null;
}
function renderSlots(employee) {
  el('slots').innerHTML = ACTIONS.map((action) => {
    const punch = earliestFor(action);
    return `<div class="slot"><strong>${LABELS[action]}</strong><span>${punch ? formatTime(punch.timestampMs) : 'Missing'}</span>${punch ? '' : `<button class="secondary-btn schedule-btn" data-fill="${action}">Use Scheduled ${formatTime(new Date(`${el('workDate').value}T${DEFAULT_SCHEDULE[action]}:00`).getTime())}</button>`}</div>`;
  }).join('');
  el('slots').querySelectorAll('[data-fill]').forEach((button) => button.addEventListener('click', () => fillScheduled(employee, button.dataset.fill)));
  const missing = ACTIONS.filter((action) => !earliestFor(action));
  setStatus(el('dayStatus'), missing.length ? `${employee.name}: ${missing.length} scheduled punch${missing.length === 1 ? '' : 'es'} missing.` : `${employee.name}: all four punches are present.`, missing.length ? '' : 'success');
}

async function fillScheduled(employee, action) {
  try {
    if (!canEdit()) throw new Error('Your account does not have edit permission.');
    if (earliestFor(action)) throw new Error(`${LABELS[action]} already exists and was not changed.`);
    const siteId = el('managerSite').value;
    const day = el('workDate').value;
    const employeeId = employee.employeeId || employee.id;
    const timestampMs = new Date(`${day}T${DEFAULT_SCHEDULE[action]}:00`).getTime();
    const punchId = await stableId('scheduled', [COMPANY_ID, siteId, employeeId, day, action]);
    const punchRef = doc(db, 'punches', punchId);
    const existingSnap = await getDoc(punchRef);
    if (existingSnap.exists() && activeRecord(existingSnap.data())) throw new Error(`${LABELS[action]} already exists and was not changed.`);

    await runTransaction(db, async (tx) => {
      const check = await tx.get(punchRef);
      if (check.exists() && activeRecord(check.data())) throw new Error(`${LABELS[action]} already exists.`);
      tx.set(punchRef, {
        companyId: COMPANY_ID, siteId, siteIds: [siteId], assignedSiteId: siteId,
        employeeId, workerId: employee.workerId || employee.id || employeeId,
        employeeNumber: employee.employeeNumber || '', agencyId: employee.agencyId || '',
        name: employee.name, nameKey: normalizeName(employee.name), action,
        timestampMs, timestamp: new Date(timestampMs), dateKey: day,
        weekKey: mondayKey(new Date(`${day}T12:00:00`)), source: 'manager_scheduled_fill',
        active: true, status: 'active', createdAt: serverTimestamp(), updatedAt: serverTimestamp(),
        createdBy: state.user.uid, createdByEmail: state.user.email || '', editedBy: state.user.uid,
        editedByEmail: state.user.email || '', editReason: 'Filled from scheduled daily time'
      });
    });
    await addDoc(collection(db, 'punch_edits'), {
      companyId: COMPANY_ID, siteId, employeeId, punchId, action: 'scheduled_fill',
      punchAction: action, timestampMs, userId: state.user.uid, userEmail: state.user.email || '',
      role: normalizeRole(state.profile?.role || (protectedOwner() ? 'owner' : 'manager')),
      reason: 'Filled from scheduled daily time', createdAt: serverTimestamp()
    });
    await loadDay();
  } catch (error) {
    setStatus(el('dayStatus'), error.message || 'Scheduled punch could not be created.', 'error');
  }
}

el('fillAll').addEventListener('click', async () => {
  const employee = state.employees.find((row) => row.id === el('employeeSelect').value);
  const missing = ACTIONS.filter((action) => !earliestFor(action));
  if (!employee || !missing.length) return;
  const list = missing.map((a) => `${LABELS[a]} ${DEFAULT_SCHEDULE[a]}`).join('\n');
  if (!confirm(`Create these missing scheduled punches for ${employee.name}?\n\n${list}`)) return;
  for (const action of missing) await fillScheduled(employee, action);
});

el('loginForm').addEventListener('submit', async (event) => {
  event.preventDefault();
  try {
    setStatus(el('managerStatus'), 'Signing in…');
    await signInWithEmailAndPassword(auth, el('email').value.trim(), el('password').value);
  } catch (error) {
    setStatus(el('managerStatus'), error.message || 'Sign in failed.', 'error');
  }
});
el('signOut').addEventListener('click', () => signOut(auth));
el('managerSite').addEventListener('change', async () => { await loadEmployees(); el('slots').innerHTML = ''; });
el('loadDay').addEventListener('click', async () => { try { await loadDay(); } catch (error) { setStatus(el('dayStatus'), error.message, 'error'); } });
el('workDate').value = dateKey(new Date());

onAuthStateChanged(auth, async (user) => {
  state.user = user;
  state.profile = null;
  if (!user) {
    el('managerPanel').classList.add('hidden');
    el('signOut').classList.add('hidden');
    setStatus(el('managerStatus'), 'Not signed in.');
    return;
  }
  try {
    const profileSnap = await getDoc(doc(db, 'users', user.uid));
    state.profile = profileSnap.exists() ? { uid: user.uid, ...profileSnap.data() } : { uid: user.uid, email: user.email || '' };
    if (!canEdit()) throw new Error('Signed in, but this account is not an owner, admin, manager, or supervisor.');
    const sites = allowedSites();
    el('managerSite').innerHTML = sites.map((site) => `<option value="${site}">${site}</option>`).join('');
    el('managerPanel').classList.remove('hidden');
    el('signOut').classList.remove('hidden');
    setStatus(el('managerStatus'), `Signed in as ${state.profile.name || user.email}. Your user profile is read-only on this page and will not be changed.`, 'success');
    await loadEmployees();
  } catch (error) {
    el('managerPanel').classList.add('hidden');
    setStatus(el('managerStatus'), error.message || 'Could not load manager access.', 'error');
  }
});
