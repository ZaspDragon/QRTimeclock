import { getApp, getApps } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  addDoc,
  collection,
  doc,
  getDocs,
  getFirestore,
  limit,
  query,
  serverTimestamp,
  updateDoc,
  where,
  writeBatch,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

// Manager-only repair helper for legacy duplicate employee profiles.
// This tool never hard-deletes employees or punches. It preserves every punch
// time/action/agency and only redirects identity references to the chosen primary.

const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
let groups = [];

function dbInstance() {
  if (!getApps().length) throw new Error('Timeclock database is not ready yet.');
  return getFirestore(getApp());
}

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function siteOf(row) {
  return String(row?.siteId || row?.assignedSiteId || row?.branchId || row?.branch || '').trim().toUpperCase();
}

function agencyOf(row) {
  return String(row?.agencyId || row?.staffingAgencyId || '').trim();
}

function nameOf(row) {
  return String(row?.name || row?.employeeName || row?.displayName || row?.fullName || '').trim();
}

function isActive(row) {
  return row?.active !== false && !['inactive', 'terminated', 'merged', 'deleted', 'removed', 'archived'].includes(String(row?.status || '').toLowerCase());
}

function currentSite() {
  const fromUrl = String(new URLSearchParams(location.search).get('site') || '').toUpperCase();
  if (VALID_SITES.has(fromUrl)) return fromUrl;
  const fromEmployeeForm = String(document.getElementById('empSiteInput')?.value || '').toUpperCase();
  return VALID_SITES.has(fromEmployeeForm) ? fromEmployeeForm : 'OH01';
}

function agencyLabel(value) {
  if (value === 'sterling_staffing') return 'Sterling Staffing';
  if (value === 'excel_staffing') return 'Excel Staffing';
  if (value === 'lifestyle_staffing') return 'Lifestyle Staffing';
  return value || 'Direct / blank';
}

function recordScore(row) {
  let score = 0;
  if (agencyOf(row)) score += 30;
  const number = String(row.employeeNumber || '').trim();
  if (number && !/^PUBLIC-|^AUTO-/i.test(number)) score += 20;
  if (!/^public_|^auto_/i.test(String(row.id || ''))) score += 10;
  if (String(row.source || '') !== 'auto_created') score += 5;
  return score;
}

function chooseSuggestedPrimary(rows) {
  return [...rows].sort((a, b) => recordScore(b) - recordScore(a) || String(a.id).localeCompare(String(b.id)))[0];
}

async function loadActiveEmployees(siteId) {
  const db = dbInstance();
  const jobs = [
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('siteId', '==', siteId), limit(500)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('assignedSiteId', '==', siteId), limit(500)),
  ];
  const unique = new Map();
  const results = await Promise.allSettled(jobs.map((job) => getDocs(job)));
  let successful = 0;
  results.forEach((result) => {
    if (result.status !== 'fulfilled') return;
    successful += 1;
    result.value.docs.forEach((snapshot) => unique.set(snapshot.id, { id: snapshot.id, ...snapshot.data() }));
  });
  if (!successful) throw new Error('Employee roster could not be read with the current permissions.');
  return [...unique.values()].filter((row) => isActive(row) && (!siteOf(row) || siteOf(row) === siteId));
}

function buildGroups(rows) {
  const byName = new Map();
  rows.forEach((row) => {
    const key = `${normalizeName(nameOf(row))}|${siteOf(row) || currentSite()}`;
    if (!normalizeName(nameOf(row))) return;
    if (!byName.has(key)) byName.set(key, []);
    byName.get(key).push(row);
  });
  return [...byName.values()]
    .filter((records) => records.length > 1)
    .map((records) => ({ records, suggested: chooseSuggestedPrimary(records) }))
    .sort((a, b) => nameOf(a.records[0]).localeCompare(nameOf(b.records[0])));
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>'"]/g, (char) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', "'": '&#39;', '"': '&quot;' }[char]));
}

function setStatus(message, error = false) {
  const el = document.getElementById('crossAgencyDuplicateStatus');
  if (!el) return;
  el.textContent = message;
  el.style.borderColor = error ? 'rgba(255,90,90,.6)' : '';
}

function renderGroups() {
  const host = document.getElementById('crossAgencyDuplicateGroups');
  if (!host) return;
  if (!groups.length) {
    host.innerHTML = '<div class="status-box">No active same-name duplicate profiles were found for this branch.</div>';
    return;
  }

  host.innerHTML = groups.map((group, index) => {
    const name = nameOf(group.records[0]);
    const options = group.records.map((row) => {
      const selected = row.id === group.suggested.id ? ' selected' : '';
      return `<option value="${escapeHtml(row.id)}"${selected}>${escapeHtml(nameOf(row))} — ${escapeHtml(agencyLabel(agencyOf(row)))} — ${escapeHtml(row.employeeNumber || row.id)}</option>`;
    }).join('');
    const details = group.records.map((row) => `<li><strong>${escapeHtml(agencyLabel(agencyOf(row)))}</strong> · ${escapeHtml(row.employeeNumber || 'No employee #')} · <code>${escapeHtml(row.id)}</code></li>`).join('');
    return `
      <div class="status-box" style="margin-top:10px;">
        <strong>${escapeHtml(name)}</strong> — ${group.records.length} active profiles
        <ul style="margin:8px 0 10px 20px;">${details}</ul>
        <label>
          <span>Keep as primary worker</span>
          <select class="cross-agency-primary" data-group="${index}">${options}</select>
        </label>
        <button class="secondary-btn cross-agency-merge" data-group="${index}" type="button" style="margin-top:10px;">Merge duplicates into selected primary</button>
      </div>`;
  }).join('');

  host.querySelectorAll('.cross-agency-merge').forEach((button) => {
    button.addEventListener('click', () => mergeGroup(Number(button.dataset.group)));
  });
}

async function queryIdentityReferences(collectionName, field, employeeId) {
  try {
    const snapshot = await getDocs(query(collection(dbInstance(), collectionName), where(field, '==', employeeId), limit(500)));
    return snapshot.docs.map((snap) => ({ ref: snap.ref, id: snap.id, data: snap.data() }));
  } catch (error) {
    console.warn(`[duplicate-repair] ${collectionName}.${field} lookup skipped:`, error?.message || error);
    return [];
  }
}

async function collectReferences(collectionName, duplicateId) {
  const results = await Promise.all([
    queryIdentityReferences(collectionName, 'employeeId', duplicateId),
    queryIdentityReferences(collectionName, 'workerId', duplicateId),
  ]);
  const unique = new Map();
  results.flat().forEach((row) => unique.set(row.id, row));
  return [...unique.values()];
}

async function commitInChunks(operations, chunkSize = 350) {
  const db = dbInstance();
  for (let i = 0; i < operations.length; i += chunkSize) {
    const batch = writeBatch(db);
    operations.slice(i, i + chunkSize).forEach((operation) => batch.update(operation.ref, operation.payload));
    await batch.commit();
  }
}

async function mergeDuplicateIntoPrimary(primary, duplicate) {
  const collections = ['punches', 'timesheets', 'missedPunchRequests'];
  const allRefs = [];
  for (const collectionName of collections) {
    const refs = await collectReferences(collectionName, duplicate.id);
    refs.forEach((row) => allRefs.push({ collectionName, ...row }));
  }

  const operations = allRefs.map((row) => {
    const originalEmployeeId = String(row.data.employeeId || '');
    const originalWorkerId = String(row.data.workerId || '');
    const payload = {
      employeeId: primary.id,
      workerId: primary.id,
      canonicalEmployeeId: primary.id,
      employeeNumber: primary.employeeNumber || row.data.employeeNumber || '',
      identityMergedAt: serverTimestamp(),
      identityMergedFrom: duplicate.id,
      mergePreservedOriginal: {
        employeeId: originalEmployeeId,
        workerId: originalWorkerId,
        employeeNumber: row.data.employeeNumber || '',
        name: row.data.name || row.data.workerName || '',
      },
    };
    // Deliberately do NOT change action, timestampMs, date/week, hours, agencyId,
    // signatures, approval state, or source. Payroll history remains intact.
    return { ref: row.ref, payload };
  });
  await commitInChunks(operations);

  await updateDoc(doc(dbInstance(), 'employees', duplicate.id), {
    status: 'merged',
    active: false,
    mergedInto: primary.id,
    canonicalEmployeeId: primary.id,
    mergedAt: serverTimestamp(),
    mergeReason: 'Cross-agency duplicate identity repair',
    updatedAt: serverTimestamp(),
  });

  await updateDoc(doc(dbInstance(), 'employees', primary.id), {
    canonicalEmployeeId: primary.id,
    identityVersion: 2,
    updatedAt: serverTimestamp(),
  });

  await addDoc(collection(dbInstance(), 'employee_merges'), {
    companyId: COMPANY_ID,
    siteId: currentSite(),
    primaryEmployeeId: primary.id,
    duplicateEmployeeId: duplicate.id,
    primaryAgencyId: agencyOf(primary),
    duplicateAgencyId: agencyOf(duplicate),
    recordsReassigned: operations.length,
    hardDeletedRecords: 0,
    preservedHistoricalAgency: true,
    createdAt: serverTimestamp(),
    reason: 'Cross-agency duplicate identity repair',
  }).catch((error) => console.warn('[duplicate-repair] merge log skipped:', error?.message || error));

  return operations.length;
}

async function mergeGroup(groupIndex) {
  const group = groups[groupIndex];
  const select = document.querySelector(`.cross-agency-primary[data-group="${groupIndex}"]`);
  const primaryId = String(select?.value || '');
  const primary = group?.records.find((row) => row.id === primaryId);
  if (!group || !primary) return;
  const duplicates = group.records.filter((row) => row.id !== primaryId);

  const summary = duplicates.map((row) => `${agencyLabel(agencyOf(row))} (${row.employeeNumber || row.id})`).join(', ');
  const ok = window.confirm(
    `Keep ${nameOf(primary)} — ${agencyLabel(agencyOf(primary))} as the primary worker?\n\n` +
    `The following duplicate profile(s) will be marked MERGED: ${summary}.\n\n` +
    'All existing punch times, actions, dates, hours, and historical agency values will be preserved. Nothing is hard-deleted.'
  );
  if (!ok) return;

  setStatus(`Merging ${duplicates.length} duplicate profile(s) for ${nameOf(primary)}...`);
  document.querySelectorAll('.cross-agency-merge').forEach((button) => { button.disabled = true; });
  try {
    let reassigned = 0;
    for (const duplicate of duplicates) reassigned += await mergeDuplicateIntoPrimary(primary, duplicate);
    setStatus(`Merged ${duplicates.length} duplicate profile(s). ${reassigned} related record(s) were relinked; no time records were deleted. Reloading...`);
    window.setTimeout(() => window.location.reload(), 1400);
  } catch (error) {
    console.error('[duplicate-repair]', error);
    setStatus(error?.message || 'Duplicate merge failed. No additional changes will be attempted.', true);
    document.querySelectorAll('.cross-agency-merge').forEach((button) => { button.disabled = false; });
  }
}

async function scan() {
  const siteId = currentSite();
  setStatus(`Scanning ${siteId} for active same-name profiles...`);
  try {
    groups = buildGroups(await loadActiveEmployees(siteId));
    renderGroups();
    setStatus(groups.length
      ? `Found ${groups.length} same-name duplicate group(s). Review the agency and employee number before merging.`
      : `No active same-name duplicate profiles found for ${siteId}.`);
  } catch (error) {
    console.error('[duplicate-repair scan]', error);
    setStatus(error?.message || 'Could not scan employee records.', true);
  }
}

function install() {
  const employeesTab = document.getElementById('employeesTab');
  if (!employeesTab || document.getElementById('crossAgencyDuplicateRepairCard')) return;

  const card = document.createElement('div');
  card.id = 'crossAgencyDuplicateRepairCard';
  card.className = 'card';
  card.innerHTML = `
    <div class="card-head split-head">
      <div>
        <h2>Cross-Agency Duplicate Repair</h2>
        <p>Find the same worker stored under multiple agencies. Merging relinks identity only; punch times and historical agency values are preserved.</p>
      </div>
      <button id="scanCrossAgencyDuplicatesBtn" class="secondary-btn" type="button">Scan Duplicates</button>
    </div>
    <div id="crossAgencyDuplicateStatus" class="status-box">Nothing changes until you scan, review a group, and confirm a merge.</div>
    <div id="crossAgencyDuplicateGroups"></div>
  `;
  employeesTab.appendChild(card);
  document.getElementById('scanCrossAgencyDuplicatesBtn')?.addEventListener('click', scan);
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
else install();

console.info('[QRTimeclock] Cross-agency duplicate repair tool installed.');
