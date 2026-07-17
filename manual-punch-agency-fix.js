import { firebaseConfig } from './firebase-config.js';
import { initializeApp, getApps } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { getAuth } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  addDoc,
  collection,
  doc,
  getDoc,
  getDocs,
  limit,
  query,
  serverTimestamp,
  updateDoc,
  where
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';
import { getFirestore } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const CURRENT_COMPANY_ID = 'chadwell';
const CURRENT_SITE_ID = 'OH01';
const BRANCH_OPTIONS = ['OH01', 'OHC'];
const PUNCH_EDIT_ROLES = new Set(['manager', 'admin', 'supervisor', 'agency_admin', 'owner', 'superadmin', 'super_admin']);
const PUNCH_DELETE_ROLES = new Set(['admin', 'owner', 'superadmin', 'super_admin']);

const els = {
  manualPunchForm: document.getElementById('manualPunchForm'),
  manualPunchNameInput: document.getElementById('manualPunchNameInput'),
  manualPunchActionInput: document.getElementById('manualPunchActionInput'),
  manualPunchDateInput: document.getElementById('manualPunchDateInput'),
  manualPunchTimeInput: document.getElementById('manualPunchTimeInput'),
  weekPicker: document.getElementById('weekPicker'),
};

let cachedProfile = null;
let profileUserId = '';

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .replaceAll(' ', '_');
}

function prettifyHumanName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .replace(/\b\w/g, (c) => c.toUpperCase());
}

function normalizeRole(role) {
  return String(role || '').trim().toLowerCase().replaceAll('-', '_');
}

function normalizeSiteId(value, fallback = CURRENT_SITE_ID) {
  const siteId = String(value || '').trim();
  if (BRANCH_OPTIONS.includes(siteId)) return siteId;
  return BRANCH_OPTIONS.includes(fallback) ? fallback : CURRENT_SITE_ID;
}

function parseSiteIds(value, fallbackToCurrent = true) {
  const raw = Array.isArray(value) ? value : String(value || '').split(',');
  const sites = raw.map((site) => normalizeSiteId(site, '')).filter(Boolean);
  if (!sites.length && !fallbackToCurrent) return [];
  return [...new Set(sites.length ? sites : [CURRENT_SITE_ID])];
}

function getAllowedSiteIds(profile = {}) {
  if (Array.isArray(profile.branches) && profile.branches.length) return parseSiteIds(profile.branches, false);
  if (profile.branch) return parseSiteIds([profile.branch], false);
  if (Array.isArray(profile.siteIds) && profile.siteIds.length) return parseSiteIds(profile.siteIds, false);
  return parseSiteIds(profile.siteId || profile.assignedSiteId || '', false);
}

function activeSiteId(profile = {}) {
  const allowed = getAllowedSiteIds(profile);
  const stored = sessionStorage.getItem(`managerActiveBranch:${profile.uid || ''}`);
  const requested = normalizeSiteId(stored, '');
  if (allowed.includes(requested)) return requested;
  return allowed[0] || CURRENT_SITE_ID;
}

function firstPresent(...values) {
  return values.find((value) => String(value || '').trim()) || '';
}

function agencyScopeId(profile = cachedProfile) {
  return String(profile?.agencyId || '').trim();
}

function isAgencyUser(profile = cachedProfile) {
  return normalizeRole(profile?.role) === 'agency_admin' || !!agencyScopeId(profile);
}

function canEditPunches(profile = cachedProfile) {
  const role = normalizeRole(profile?.role);
  if (['owner', 'superadmin', 'super_admin'].includes(role)) return true;
  return PUNCH_EDIT_ROLES.has(role) && (profile?.permissions?.canEditPunches === true || ['admin', 'agency_admin', 'manager', 'supervisor'].includes(role));
}

function canDeletePunches(profile = cachedProfile) {
  const role = normalizeRole(profile?.role);
  if (['owner', 'superadmin', 'super_admin'].includes(role)) return true;
  return PUNCH_DELETE_ROLES.has(role) && (profile?.permissions?.canDeletePunches === true || role === 'admin');
}

function normalizeUserProfile(profile = {}, authUser = null) {
  const role = profile.role || 'worker';
  const displayName = prettifyHumanName(
    profile.displayName ||
    profile.name ||
    [profile.firstName, profile.lastName].filter(Boolean).join(' ') ||
    authUser?.displayName ||
    profile.email ||
    authUser?.email ||
    ''
  );
  return {
    ...profile,
    uid: profile.uid || authUser?.uid || '',
    name: displayName || profile.email || authUser?.email || 'Signed in',
    email: String(profile.email || authUser?.email || '').trim().toLowerCase(),
    role,
    companyId: profile.companyId || CURRENT_COMPANY_ID,
    agencyId: profile.agencyId || '',
    branch: profile.branch || profile.siteId || profile.assignedSiteId || '',
    branches: Array.isArray(profile.branches) ? parseSiteIds(profile.branches, false) : [],
    siteId: profile.siteId || profile.branch || profile.assignedSiteId || '',
    siteIds: parseSiteIds(profile.branches || profile.branch || profile.siteIds || profile.siteId || profile.assignedSiteId || '', false),
    permissions: profile.permissions || {},
  };
}

async function getProfile() {
  const user = auth.currentUser;
  if (!user) throw new Error('Sign in before editing punches.');
  if (cachedProfile && profileUserId === user.uid) return cachedProfile;

  const profileSnap = await getDoc(doc(db, 'users', user.uid));
  cachedProfile = normalizeUserProfile(profileSnap.exists() ? profileSnap.data() : {}, user);
  profileUserId = user.uid;
  return cachedProfile;
}

function employeeName(employee) {
  if (!employee) return '';
  const full = prettifyHumanName([employee.firstName, employee.lastName].filter(Boolean).join(' '));
  return prettifyHumanName(employee.name || employee.displayName || employee.employeeName || employee.workerName || full || '');
}

function employeeBranchId(employee) {
  return String(employee?.branchId || employee?.branchCode || employee?.siteId || employee?.assignedSiteId || '').trim();
}

function employeeBranchName(employee) {
  return String(employee?.branchName || employee?.siteName || employeeBranchId(employee) || '').trim();
}

function isActiveEmployeeRecord(record) {
  if (!record || typeof record !== 'object') return false;
  if (record.active === false) return false;
  const status = String(record.status || 'active').trim().toLowerCase();
  return !['inactive', 'removed', 'terminated', 'disabled', 'archived', 'merged'].includes(status);
}

function rowSiteMatches(row, siteId) {
  const rowSites = [
    row.siteId,
    row.assignedSiteId,
    row.branchId,
    ...(Array.isArray(row.siteIds) ? row.siteIds : []),
  ].map((value) => String(value || '').trim()).filter(Boolean);
  return !rowSites.length || rowSites.includes(siteId);
}

function chooseBestEmployee(matches, profile) {
  const siteId = activeSiteId(profile);
  const agencyId = agencyScopeId(profile);
  return [...matches].sort((left, right) => {
    const leftScore = Number(rowSiteMatches(left, siteId)) * 4
      + Number(!agencyId || String(left.agencyId || '') === agencyId) * 3
      + Number(String(left.companyId || CURRENT_COMPANY_ID) === CURRENT_COMPANY_ID);
    const rightScore = Number(rowSiteMatches(right, siteId)) * 4
      + Number(!agencyId || String(right.agencyId || '') === agencyId) * 3
      + Number(String(right.companyId || CURRENT_COMPANY_ID) === CURRENT_COMPANY_ID);
    return rightScore - leftScore || String(left.name || '').localeCompare(String(right.name || ''));
  })[0] || null;
}

async function findManualPunchEmployeeMatch(nameKey, profile) {
  if (!nameKey) return null;
  const employeesRef = collection(db, 'employees');
  const baseFilters = [
    where('companyId', '==', CURRENT_COMPANY_ID),
    where('status', '==', 'active'),
  ];
  if (isAgencyUser(profile) && agencyScopeId(profile)) {
    baseFilters.push(where('agencyId', '==', agencyScopeId(profile)));
  }

  const matches = new Map();
  const addMatches = (snap) => {
    snap.docs.forEach((record) => {
      const row = { id: record.id, ...record.data() };
      const rowNameKey = row.nameKey || normalizeName(employeeName(row));
      if (isActiveEmployeeRecord(row) && rowNameKey === nameKey) {
        matches.set(record.id, row);
      }
    });
  };

  addMatches(await getDocs(query(employeesRef, ...baseFilters, where('nameKey', '==', nameKey), limit(20))));
  if (!matches.size) {
    addMatches(await getDocs(query(employeesRef, ...baseFilters, where('normalizedName', '==', nameKey), limit(20))));
  }

  return chooseBestEmployee([...matches.values()], profile);
}

function buildBranchPayload(siteId) {
  return {
    companyId: CURRENT_COMPANY_ID,
    siteId: normalizeSiteId(siteId),
  };
}

function buildIdentityPayload(employee, fallback, profile) {
  const fallbackName = prettifyHumanName(fallback?.name || '');
  const name = employee ? employeeName(employee) || fallbackName : fallbackName;
  const nameKey = normalizeName(name || fallback?.nameKey || '');
  const employeeDocId = employee?.id || '';
  const employeeId = firstPresent(employee?.employeeId, employee?.workerId, employeeDocId);
  const workerId = firstPresent(employee?.workerId, employee?.employeeId, employeeDocId);
  const siteId = normalizeSiteId(firstPresent(employee?.siteId, employee?.assignedSiteId, employeeBranchId(employee), activeSiteId(profile)), activeSiteId(profile));
  const agencyId = firstPresent(employee?.agencyId, agencyScopeId(profile));

  return {
    name,
    workerName: name,
    employeeName: name,
    displayName: name,
    nameKey,
    employeeId,
    workerId,
    workerIdentityKey: employeeId || workerId ? `worker:${employeeId || workerId}` : `name:${nameKey}`,
    employeeNumber: employee?.employeeNumber || '',
    companyId: CURRENT_COMPANY_ID,
    agencyId,
    agencyName: employee?.agencyName || '',
    siteId,
    branchId: firstPresent(employee?.branchId, employee?.branchCode, siteId),
    branchName: employeeBranchName(employee),
    assignedSiteId: normalizeSiteId(firstPresent(employee?.assignedSiteId, employee?.siteId, siteId), siteId),
    siteIds: [siteId],
  };
}

function toSnakePunchAction(action) {
  const value = String(action || '').trim();
  const aliases = {
    clockIn: 'clock_in',
    startLunch: 'start_lunch',
    endLunch: 'end_lunch',
    clockOut: 'clock_out',
    'clock in': 'clock_in',
    'start lunch': 'start_lunch',
    'end lunch': 'end_lunch',
    'clock out': 'clock_out',
  };
  return aliases[value] || aliases[value.toLowerCase()] || value;
}

function parseLocalDateAndTime(dateValue, timeValue) {
  if (!dateValue || !timeValue) return 0;
  const [year, month, day] = dateValue.split('-').map(Number);
  const [hour, minute] = timeValue.split(':').map(Number);
  const date = new Date(year, month - 1, day, hour, minute, 0, 0);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function parseLocalEditString(value) {
  const cleaned = String(value || '').trim().replace('T', ' ');
  const match = cleaned.match(/^(\d{4})-(\d{2})-(\d{2})\s+(\d{2}):(\d{2})$/);
  if (!match) return 0;
  const [, year, month, day, hour, minute] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day), Number(hour), Number(minute), 0, 0);
  return Number.isNaN(date.getTime()) ? 0 : date.getTime();
}

function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getMondayDate(inputDate) {
  const date = new Date(inputDate);
  date.setHours(0, 0, 0, 0);
  const day = date.getDay();
  const diff = day === 0 ? -6 : 1 - day;
  date.setDate(date.getDate() + diff);
  return date;
}

function formatDateInput(date) {
  return formatDateKey(date);
}

function formatTimeForInput(ms) {
  const date = new Date(ms);
  return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

function toLocalEditString(ms) {
  if (!ms) return '';
  const date = new Date(ms);
  return `${formatDateKey(date)} ${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

function selectedPreviewWeekKey() {
  const weekValue = els.weekPicker?.value;
  if (!weekValue) return formatDateKey(getMondayDate(new Date()));
  return formatDateKey(getMondayDate(new Date(`${weekValue}T00:00:00`)));
}

function ensureWeekWarning() {
  let warning = document.getElementById('manualPunchWeekWarning');
  if (!warning && els.manualPunchForm) {
    warning = document.createElement('div');
    warning.id = 'manualPunchWeekWarning';
    warning.className = 'inline-warning full-width hidden';
    const actions = els.manualPunchForm.querySelector('.form-actions');
    els.manualPunchForm.insertBefore(warning, actions || null);
  }
  return warning;
}

function updateManualPunchWeekWarning(explicitWeekKey = '') {
  const warning = ensureWeekWarning();
  if (!warning) return;
  const dateValue = els.manualPunchDateInput?.value || '';
  const punchWeekKey = explicitWeekKey || (dateValue ? formatDateKey(getMondayDate(new Date(`${dateValue}T00:00:00`))) : '');
  const previewWeekKey = selectedPreviewWeekKey();
  if (!punchWeekKey || punchWeekKey === previewWeekKey) {
    warning.classList.add('hidden');
    warning.textContent = '';
    return;
  }
  warning.textContent = `This manual punch is for week ${punchWeekKey}, but the current timesheet preview is week ${previewWeekKey}. Change the preview week to see it there.`;
  warning.classList.remove('hidden');
}

function installWarningStyles() {
  if (document.getElementById('manualPunchAgencyFixStyles')) return;
  const style = document.createElement('style');
  style.id = 'manualPunchAgencyFixStyles';
  style.textContent = `
    .inline-warning {
      border: 1px solid rgba(245, 158, 11, 0.45);
      background: rgba(245, 158, 11, 0.12);
      color: #92400e;
      border-radius: 10px;
      padding: 10px 12px;
      font-size: 0.9rem;
      font-weight: 600;
    }
  `;
  document.head.appendChild(style);
}

function actorName(profile) {
  return profile?.name || auth.currentUser?.email || 'Manager';
}

async function writeAudit(action, entityType, entityId, oldValue, newValue, reason, profile) {
  try {
    await addDoc(collection(db, 'auditLogs'), {
      ...buildBranchPayload(activeSiteId(profile)),
      agencyId: agencyScopeId(profile),
      userId: auth.currentUser?.uid || '',
      actorId: auth.currentUser?.uid || '',
      actorRole: profile?.role || '',
      role: profile?.role || '',
      action,
      eventType: action,
      entityType,
      entityId,
      affectedRecord: entityId,
      oldValue,
      newValue,
      reason,
      timestamp: serverTimestamp(),
      createdAt: serverTimestamp(),
    });
  } catch (error) {
    console.warn('Manual punch audit log write failed:', error.message);
  }
}

function requestAppTimesheetRefresh() {
  els.weekPicker?.dispatchEvent(new Event('change', { bubbles: true }));
  window.setTimeout(() => {
    document.getElementById('agencyPreviewBtn')?.click();
  }, 500);
}

function notify(message, isError = false) {
  const warning = ensureWeekWarning();
  if (!warning) {
    if (isError) alert(message);
    return;
  }
  warning.textContent = message;
  warning.classList.toggle('hidden', false);
  warning.style.borderColor = isError ? 'rgba(239, 68, 68, 0.55)' : 'rgba(43, 213, 118, 0.45)';
  warning.style.background = isError ? 'rgba(239, 68, 68, 0.12)' : 'rgba(43, 213, 118, 0.12)';
  warning.style.color = isError ? '#991b1b' : '#166534';
  window.setTimeout(() => {
    warning.removeAttribute('style');
    updateManualPunchWeekWarning();
  }, 3500);
}

async function handleManualPunchSubmit(event) {
  event.preventDefault();
  event.stopImmediatePropagation();

  try {
    const profile = await getProfile();
    if (!canEditPunches(profile)) throw new Error('You need edit-punch permission to add manual punches.');

    const typedName = prettifyHumanName(els.manualPunchNameInput?.value.trim());
    const typedNameKey = normalizeName(typedName);
    const action = toSnakePunchAction(els.manualPunchActionInput?.value);
    const dateValue = els.manualPunchDateInput?.value;
    const timeValue = els.manualPunchTimeInput?.value;

    if (!typedName || !typedNameKey) throw new Error('Enter a valid name.');
    if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(action) || !dateValue || !timeValue) {
      throw new Error('Fill out all manual punch fields.');
    }

    const parsedMs = parseLocalDateAndTime(dateValue, timeValue);
    if (!parsedMs) throw new Error('Invalid date or time.');

    const employeeMatch = await findManualPunchEmployeeMatch(typedNameKey, profile);
    const identity = buildIdentityPayload(employeeMatch, { name: typedName, nameKey: typedNameKey }, profile);
    const punchDate = new Date(parsedMs);
    const dateKey = formatDateKey(punchDate);
    const weekKey = formatDateKey(getMondayDate(punchDate));
    const siteId = identity.siteId || activeSiteId(profile);
    const payload = {
      ...identity,
      ...buildBranchPayload(siteId),
      action,
      timestamp: serverTimestamp(),
      timestampMs: parsedMs,
      dateKey,
      weekKey,
      source: 'manual_manager',
      createdAt: serverTimestamp(),
      createdBy: actorName(profile),
    };

    const punchRef = await addDoc(collection(db, 'punches'), payload);
    await addDoc(collection(db, 'punch_edits'), {
      punchId: punchRef.id,
      type: 'manual_add',
      ...payload,
      timestamp: null,
      createdAt: null,
      editedBy: actorName(profile),
      editedAt: serverTimestamp(),
    });

    els.manualPunchForm?.reset();
    if (els.manualPunchDateInput) els.manualPunchDateInput.value = formatDateInput(new Date());
    if (els.manualPunchTimeInput) els.manualPunchTimeInput.value = formatTimeForInput(Date.now());
    updateManualPunchWeekWarning();
    requestAppTimesheetRefresh();
    await writeAudit('punch_manual_added', 'punch', punchRef.id, {}, payload, 'Manual manager punch', profile);
    notify('Manual punch added.');
  } catch (error) {
    console.error(error);
    notify(error.message || 'Could not add manual punch.', true);
  }
}

async function editPunch(punchId) {
  try {
    const profile = await getProfile();
    if (!canEditPunches(profile)) throw new Error('You need edit-punch permission to edit punches.');

    const snap = await getDoc(doc(db, 'punches', punchId));
    if (!snap.exists()) throw new Error('Punch not found.');
    const row = { id: snap.id, ...snap.data() };

    const newName = prompt('Edit worker name:', row.name || '');
    if (newName === null) return;
    const newAction = prompt('Edit action (clock_in, start_lunch, end_lunch, clock_out):', row.action || 'clock_in');
    if (newAction === null) return;
    const newDateTime = prompt('Edit date/time (example: 2026-04-14 07:26):', toLocalEditString(row.timestampMs));
    if (newDateTime === null) return;

    const prettyName = prettifyHumanName(newName);
    const nameKey = normalizeName(prettyName);
    const action = toSnakePunchAction(newAction);
    if (!prettyName || nameKey.length < 2) throw new Error('Invalid name.');
    if (!['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(action)) throw new Error('Invalid action.');

    const parsedMs = parseLocalEditString(newDateTime);
    if (!parsedMs) throw new Error('Invalid date/time format. Use YYYY-MM-DD HH:MM');

    const date = new Date(parsedMs);
    const dateKey = formatDateKey(date);
    const weekKey = formatDateKey(getMondayDate(date));
    const employeeMatch = await findManualPunchEmployeeMatch(nameKey, profile);
    const identitySource = employeeMatch || row;
    const identity = buildIdentityPayload(identitySource, { name: prettyName, nameKey }, profile);
    const siteId = identity.siteId || row.siteId || activeSiteId(profile);
    const updatedPayload = {
      ...identity,
      ...buildBranchPayload(siteId),
      action,
      timestampMs: parsedMs,
      dateKey,
      weekKey,
      editedAt: serverTimestamp(),
      editedBy: actorName(profile),
    };

    await addDoc(collection(db, 'punch_edits'), {
      punchId,
      type: 'edit',
      original: {
        name: row.name || '',
        nameKey: row.nameKey || '',
        action: row.action || '',
        timestampMs: row.timestampMs || 0,
        dateKey: row.dateKey || '',
        weekKey: row.weekKey || '',
        employeeId: row.employeeId || '',
        workerId: row.workerId || '',
        agencyId: row.agencyId || '',
        source: row.source || '',
        editedBy: row.editedBy || '',
      },
      updated: {
        ...identity,
        action,
        timestampMs: parsedMs,
        dateKey,
        weekKey,
        source: row.source || '',
      },
      editedBy: actorName(profile),
      editedAt: serverTimestamp(),
      ...buildBranchPayload(siteId),
      agencyId: identity.agencyId || '',
      branchId: identity.branchId || siteId,
    });

    await updateDoc(doc(db, 'punches', punchId), updatedPayload);
    requestAppTimesheetRefresh();
    await writeAudit('punch_edited', 'punch', punchId, row, updatedPayload, 'Punch edited from manager dashboard', profile);
    notify('Punch updated.');
  } catch (error) {
    console.error(error);
    notify(error.message || 'Could not update punch.', true);
  }
}

async function deletePunchRecord(punchId) {
  try {
    const profile = await getProfile();
    if (!canDeletePunches(profile)) throw new Error('You need delete-punch permission to delete punches.');

    const snap = await getDoc(doc(db, 'punches', punchId));
    const row = snap.exists() ? { id: snap.id, ...snap.data() } : null;
    if (!confirm('Delete this punch?')) return;

    await addDoc(collection(db, 'punch_edits'), {
      punchId,
      type: 'delete',
      original: row || null,
      editedBy: actorName(profile),
      editedAt: serverTimestamp(),
      ...buildBranchPayload(row?.siteId || activeSiteId(profile)),
      agencyId: row?.agencyId || agencyScopeId(profile),
      branchId: row?.branchId || row?.siteId || activeSiteId(profile),
    });

    const deletePayload = {
      status: 'deleted',
      active: false,
      deletedAt: serverTimestamp(),
      deletedBy: actorName(profile),
      deleteReason: 'Manager deleted from punch editor',
      updatedAt: serverTimestamp(),
    };
    await updateDoc(doc(db, 'punches', punchId), deletePayload);
    requestAppTimesheetRefresh();
    await writeAudit('punch_deleted', 'punch', punchId, row || {}, deletePayload, 'Soft delete from punch editor', profile);
    notify('Punch marked deleted.');
  } catch (error) {
    console.error(error);
    notify(error.message || 'Could not delete punch.', true);
  }
}

function interceptEditDeleteClicks(event) {
  const editButton = event.target.closest('.manager-edit-punch-btn');
  const deleteButton = event.target.closest('.manager-delete-punch-btn');
  const button = editButton || deleteButton;
  if (!button?.dataset?.id) return;

  event.preventDefault();
  event.stopImmediatePropagation();

  if (editButton) {
    editPunch(button.dataset.id);
  } else {
    deletePunchRecord(button.dataset.id);
  }
}

function initManualPunchAgencyFix() {
  installWarningStyles();
  ensureWeekWarning();
  updateManualPunchWeekWarning();
  els.manualPunchForm?.addEventListener('submit', handleManualPunchSubmit, true);
  els.manualPunchDateInput?.addEventListener('input', () => updateManualPunchWeekWarning());
  els.manualPunchDateInput?.addEventListener('change', () => updateManualPunchWeekWarning());
  els.weekPicker?.addEventListener('change', () => updateManualPunchWeekWarning());
  document.addEventListener('click', interceptEditDeleteClicks, true);
  console.info('[QRTimeclock] Manual punch agency timesheet fix loaded. Clock Out action is clock_out.');
}

initManualPunchAgencyFix();
