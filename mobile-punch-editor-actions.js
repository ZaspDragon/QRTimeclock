import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { getAuth } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  collection,
  doc,
  getDoc,
  getFirestore,
  serverTimestamp,
  updateDoc,
  writeBatch,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const VALID_SITES = new Set(['OH01', 'OHC']);
let busy = false;
let lastHandled = { key: '', at: 0 };

function showMessage(message, isError = false) {
  const status = document.getElementById('timesheetStatus')
    || document.getElementById('workerStatusMessage');
  if (status) {
    status.textContent = message;
    if (isError) status.style.borderColor = 'rgba(255,90,90,.6)';
  }
  if (typeof window.toast === 'function') window.toast(message, isError);
  else if (isError) window.alert(message);
}

function findActionButton(target) {
  if (!(target instanceof Element)) return null;
  return target.closest('.manager-edit-punch-btn, .manager-delete-punch-btn');
}

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}

function prettifyName(value) {
  return String(value || '').trim().replace(/\s+/g, ' ').replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function formatLocalEditValue(timestampMs) {
  const date = new Date(Number(timestampMs || 0));
  if (Number.isNaN(date.getTime())) return '';
  const pad = (value) => String(value).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())} ${pad(date.getHours())}:${pad(date.getMinutes())}`;
}

function parseLocalEditValue(value) {
  const match = String(value || '').trim().match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})$/);
  if (!match) return 0;
  const [, year, month, day, hour, minute] = match.map(Number);
  const date = new Date(year, month - 1, day, hour, minute, 0, 0);
  if (
    date.getFullYear() !== year
    || date.getMonth() !== month - 1
    || date.getDate() !== day
    || date.getHours() !== hour
    || date.getMinutes() !== minute
  ) return 0;
  return date.getTime();
}

function dateKeyFromMs(timestampMs) {
  const date = new Date(timestampMs);
  const pad = (value) => String(value).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}

function weekKeyFromMs(timestampMs) {
  const date = new Date(timestampMs);
  const day = date.getDay();
  const offset = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + offset);
  date.setHours(0, 0, 0, 0);
  return dateKeyFromMs(date.getTime());
}

function collectSites(row) {
  if (!row) return [];
  return [...new Set([
    row.siteId,
    row.branch,
    row.assignedSiteId,
    ...(Array.isArray(row.siteIds) ? row.siteIds : []),
  ]
    .map((value) => String(value || '').trim().toUpperCase())
    .filter((value) => VALID_SITES.has(value)))];
}

function resolvePunchSite(punch, employee) {
  const punchSites = collectSites(punch);
  if (punchSites.length === 1) return punchSites[0];
  if (punchSites.length > 1) {
    throw new Error('This punch is assigned to more than one branch. No changes were saved.');
  }

  const employeeSites = collectSites(employee);
  if (employeeSites.length === 1) return employeeSites[0];
  throw new Error('This punch does not have one clear OH01 or OHC branch. No changes were saved. Ask an administrator to correct the employee branch first.');
}

function punchIdentityId(punch) {
  return String(punch.employeeId || punch.employeeID || punch.workerId || '').trim();
}

async function loadEmployeeForPunch(punch) {
  const identityId = punchIdentityId(punch);
  if (!identityId) return null;
  try {
    const snapshot = await getDoc(doc(db, 'employees', identityId));
    return snapshot.exists() ? { id: snapshot.id, ...snapshot.data() } : null;
  } catch (_) {
    return null;
  }
}

async function loadContext(punchId) {
  const user = auth.currentUser;
  if (!user) throw new Error('Your manager session expired. Sign in again.');

  const [punchSnap, profileSnap] = await Promise.all([
    getDoc(doc(db, 'punches', punchId)),
    getDoc(doc(db, 'users', user.uid)),
  ]);
  if (!punchSnap.exists()) throw new Error('This punch no longer exists. Refresh and try again.');

  const punch = { id: punchSnap.id, ...punchSnap.data() };
  const employee = await loadEmployeeForPunch(punch);
  const profile = profileSnap.exists() ? profileSnap.data() : {};
  const siteId = resolvePunchSite(punch, employee);
  const companyId = String(punch.companyId || employee?.companyId || 'chadwell').trim();
  const agencyId = String(
    punch.agencyId
    || employee?.agencyId
    || employee?.staffingAgencyId
    || ''
  ).trim();
  return { user, punch, employee, profile, siteId, companyId, agencyId };
}

async function editPunch(button) {
  const punchId = String(button?.dataset?.id || '').trim();
  if (!punchId) throw new Error('This punch could not be identified. Refresh and try again.');
  const { user, punch, profile, siteId, companyId, agencyId } = await loadContext(punchId);

  const verifiedName = prettifyName(punch.name || punch.employeeName || punch.workerName || punch.nameKey || 'Worker');
  const enteredAction = window.prompt(
    `Edit action for ${verifiedName} (clock_in, start_lunch, end_lunch, clock_out):`,
    punch.action || 'clock_in'
  );
  if (enteredAction === null) return;
  const enteredDateTime = window.prompt(
    `Edit date/time for ${verifiedName} (YYYY-MM-DD HH:MM):`,
    formatLocalEditValue(punch.timestampMs)
  );
  if (enteredDateTime === null) return;
  const enteredReason = window.prompt('Reason required for this punch edit:');
  if (enteredReason === null) return;

  const action = String(enteredAction || '').trim().toLowerCase();
  const timestampMs = parseLocalEditValue(enteredDateTime);
  const editReason = String(enteredReason || '').trim();
  if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(action)) {
    throw new Error('Use clock_in, start_lunch, end_lunch, or clock_out.');
  }
  if (!timestampMs) throw new Error('Use date/time format YYYY-MM-DD HH:MM.');
  if (editReason.length < 2) throw new Error('Enter a clear reason for this punch edit.');

  button.disabled = true;
  button.setAttribute('aria-busy', 'true');
  const editor = String(profile.name || user.email || user.uid);
  const identity = {
    employeeId: punch.employeeId || '',
    employeeID: punch.employeeID || '',
    workerId: punch.workerId || '',
    employeeNumber: punch.employeeNumber || '',
    name: punch.name || '',
    nameKey: punch.nameKey || '',
  };
  const updated = {
    action,
    timestampMs,
    dateKey: dateKeyFromMs(timestampMs),
    weekKey: weekKeyFromMs(timestampMs),
    companyId,
    siteId,
    branch: siteId,
    editedAt: serverTimestamp(),
    editedBy: editor,
    editedByUid: user.uid,
    editReason,
    updatedAt: serverTimestamp(),
  };
  if (agencyId) updated.agencyId = agencyId;

  const editRef = doc(collection(db, 'punch_edits'));
  const batch = writeBatch(db);
  batch.update(doc(db, 'punches', punchId), updated);
  batch.set(editRef, {
    punchId,
    type: 'edit',
    original: {
      ...identity,
      action: punch.action || '',
      timestampMs: Number(punch.timestampMs || 0),
      dateKey: punch.dateKey || '',
      weekKey: punch.weekKey || '',
      companyId,
      siteId,
      branch: siteId,
      agencyId,
      source: punch.source || '',
    },
    updated: {
      ...identity,
      ...updated,
      agencyId,
      source: punch.source || '',
    },
    companyId,
    siteId,
    branch: siteId,
    agencyId,
    editedBy: editor,
    editedByUid: user.uid,
    reason: editReason,
    editReason,
    editedAt: serverTimestamp(),
  });
  await batch.commit();
  showMessage('Punch updated. Refresh Agency Export and preview the worker before printing.');
  window.setTimeout(() => window.location.reload(), 250);
}

async function mobileSoftDelete(button) {
  const punchId = String(button?.dataset?.id || '').trim();
  if (!punchId) throw new Error('This punch could not be identified. Refresh and try again.');
  const user = auth.currentUser;
  if (!user) throw new Error('Your manager session expired. Sign in again before deleting a punch.');
  const reason = window.prompt('Reason required for deleting this punch:');
  if (!String(reason || '').trim()) return;
  if (!window.confirm('Mark this punch deleted while preserving its history?')) return;
  button.disabled = true;
  button.setAttribute('aria-busy', 'true');
  await updateDoc(doc(db, 'punches', punchId), {
    status: 'deleted',
    active: false,
    deletedAt: serverTimestamp(),
    deletedBy: user.email || user.uid,
    deleteReason: String(reason).trim().slice(0, 300),
    updatedAt: serverTimestamp(),
  });
  button.closest('tr')?.remove();
  showMessage('Punch marked deleted. Historical data was preserved.');
}

function install() {
  const handle = async (event) => {
    const button = findActionButton(event.target);
    if (!button || busy) return;
    const actionType = button.classList.contains('manager-edit-punch-btn') ? 'edit' : 'delete';
    const key = `${actionType}:${button.dataset.id || ''}`;
    const now = Date.now();
    if (lastHandled.key === key && now - lastHandled.at < 800) return;
    lastHandled = { key, at: now };
    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    busy = true;
    try {
      if (actionType === 'edit') await editPunch(button);
      else await mobileSoftDelete(button);
    } catch (error) {
      console.error('[mobile-punch-editor-actions]', error);
      showMessage(error?.message || (actionType === 'edit' ? 'Could not update punch.' : 'Could not delete punch.'), true);
      button.disabled = false;
      button.removeAttribute('aria-busy');
    } finally {
      busy = false;
    }
  };
  document.addEventListener('click', handle, true);
  document.addEventListener('touchend', handle, { capture: true, passive: false });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', install, { once: true });
} else {
  install();
}
