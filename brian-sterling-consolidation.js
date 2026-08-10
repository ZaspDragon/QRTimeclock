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

const COMPANY_ID = 'chadwell';
const SITE_ID = 'OH01';
const TARGET_NAME = 'Brian Lewis Jr';
const TARGET_NAME_KEY = 'brian_lewis_jr';
const TARGET_AGENCY = 'sterling_staffing';

function dbInstance() {
  if (!getApps().length) throw new Error('Timeclock database is not ready yet.');
  return getFirestore(getApp());
}

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function rowName(row) {
  return String(row?.name || row?.employeeName || row?.displayName || row?.fullName || '').trim();
}

function rowSite(row) {
  return String(row?.siteId || row?.assignedSiteId || row?.branchId || row?.branch || '').trim().toUpperCase();
}

function rowAgency(row) {
  return String(row?.agencyId || row?.staffingAgencyId || '').trim();
}

function isActive(row) {
  return row?.active !== false && !['inactive', 'terminated', 'merged', 'deleted', 'removed', 'archived'].includes(String(row?.status || '').toLowerCase());
}

function scoreSterlingPrimary(row) {
  let score = 0;
  const employeeNumber = String(row.employeeNumber || '').trim();
  if (rowAgency(row) === TARGET_AGENCY) score += 100;
  if (employeeNumber && !/^PUBLIC-|^AUTO-/i.test(employeeNumber)) score += 30;
  if (!/^public_|^auto_/i.test(String(row.id || ''))) score += 20;
  if (String(row.source || '') !== 'auto_created') score += 10;
  return score;
}

async function loadBrianProfiles() {
  const db = dbInstance();
  const jobs = [
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('siteId', '==', SITE_ID), limit(500)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('assignedSiteId', '==', SITE_ID), limit(500)),
    query(collection(db, 'employees'), where('nameKey', '==', TARGET_NAME_KEY), limit(100)),
  ];
  const unique = new Map();
  const results = await Promise.allSettled(jobs.map((job) => getDocs(job)));
  results.forEach((result) => {
    if (result.status !== 'fulfilled') return;
    result.value.docs.forEach((snap) => unique.set(snap.id, { id: snap.id, ...snap.data() }));
  });
  return [...unique.values()].filter((row) =>
    normalizeName(rowName(row)) === normalizeName(TARGET_NAME)
    && (!rowSite(row) || rowSite(row) === SITE_ID)
  );
}

async function queryRefs(collectionName, field, id) {
  try {
    const snap = await getDocs(query(collection(dbInstance(), collectionName), where(field, '==', id), limit(500)));
    return snap.docs.map((record) => ({ ref: record.ref, id: record.id, data: record.data() }));
  } catch (error) {
    console.warn(`[brian-sterling] ${collectionName}.${field} lookup skipped:`, error?.message || error);
    return [];
  }
}

async function collectRefs(collectionName, ids) {
  const unique = new Map();
  for (const id of ids) {
    const results = await Promise.all([
      queryRefs(collectionName, 'employeeId', id),
      queryRefs(collectionName, 'workerId', id),
      queryRefs(collectionName, 'canonicalEmployeeId', id),
    ]);
    results.flat().forEach((row) => unique.set(row.id, row));
  }
  return [...unique.values()];
}

async function commitInChunks(operations, chunkSize = 350) {
  const db = dbInstance();
  for (let i = 0; i < operations.length; i += chunkSize) {
    const batch = writeBatch(db);
    operations.slice(i, i + chunkSize).forEach(({ ref, payload }) => batch.update(ref, payload));
    await batch.commit();
  }
}

function setStatus(message, isError = false) {
  const el = document.getElementById('brianSterlingStatus');
  if (!el) return;
  el.textContent = message;
  el.style.borderColor = isError ? 'rgba(255,90,90,.6)' : '';
}

async function consolidateBrianToSterling() {
  const profiles = await loadBrianProfiles();
  if (!profiles.length) throw new Error('No Brian Lewis Jr profile was found in OH01.');

  const activeSterling = profiles.filter((row) => isActive(row) && rowAgency(row) === TARGET_AGENCY);
  if (!activeSterling.length) throw new Error('No active Sterling Staffing profile was found for Brian Lewis Jr.');

  const primary = [...activeSterling].sort((a, b) => scoreSterlingPrimary(b) - scoreSterlingPrimary(a) || String(a.id).localeCompare(String(b.id)))[0];
  const duplicateProfiles = profiles.filter((row) => row.id !== primary.id && isActive(row));
  const allIds = profiles.map((row) => row.id);

  const ok = window.confirm(
    `Combine every OH01 Brian Lewis Jr profile into Sterling Staffing?\n\n` +
    `Primary record: ${primary.employeeNumber || primary.id}\n` +
    `Profiles being merged: ${duplicateProfiles.length}\n\n` +
    'All punch times, dates, actions, and hours will be preserved. Brian records will be assigned to Sterling Staffing and duplicate profiles will be marked MERGED.'
  );
  if (!ok) return;

  setStatus(`Consolidating ${profiles.length} Brian Lewis Jr profile(s) into Sterling Staffing...`);

  const collections = ['punches', 'timesheets', 'missedPunchRequests'];
  let changedRecords = 0;
  for (const collectionName of collections) {
    const refs = await collectRefs(collectionName, allIds);
    const operations = refs.map((row) => ({
      ref: row.ref,
      payload: {
        employeeId: primary.id,
        workerId: primary.id,
        canonicalEmployeeId: primary.id,
        employeeNumber: primary.employeeNumber || row.data.employeeNumber || primary.id,
        name: TARGET_NAME,
        workerName: row.data.workerName !== undefined ? TARGET_NAME : row.data.workerName,
        employeeName: row.data.employeeName !== undefined ? TARGET_NAME : row.data.employeeName,
        nameKey: TARGET_NAME_KEY,
        agencyId: TARGET_AGENCY,
        assignedSiteId: SITE_ID,
        siteId: row.data.siteId || SITE_ID,
        brianSterlingConsolidatedAt: serverTimestamp(),
      },
    }));
    await commitInChunks(operations);
    changedRecords += operations.length;
  }

  await updateDoc(doc(dbInstance(), 'employees', primary.id), {
    name: TARGET_NAME,
    nameKey: TARGET_NAME_KEY,
    normalizedName: TARGET_NAME_KEY,
    agencyId: TARGET_AGENCY,
    assignedSiteId: SITE_ID,
    siteId: SITE_ID,
    canonicalEmployeeId: primary.id,
    identityVersion: 2,
    status: 'active',
    active: true,
    updatedAt: serverTimestamp(),
  });

  for (const duplicate of duplicateProfiles) {
    await updateDoc(doc(dbInstance(), 'employees', duplicate.id), {
      status: 'merged',
      active: false,
      mergedInto: primary.id,
      canonicalEmployeeId: primary.id,
      mergedAt: serverTimestamp(),
      mergeReason: 'Brian Lewis Jr consolidated to Sterling Staffing',
      updatedAt: serverTimestamp(),
    });
  }

  await addDoc(collection(dbInstance(), 'employee_merges'), {
    companyId: COMPANY_ID,
    siteId: SITE_ID,
    workerName: TARGET_NAME,
    primaryEmployeeId: primary.id,
    primaryAgencyId: TARGET_AGENCY,
    duplicateEmployeeIds: duplicateProfiles.map((row) => row.id),
    recordsReassigned: changedRecords,
    punchTimesPreserved: true,
    targetAgencyApplied: true,
    createdAt: serverTimestamp(),
    reason: 'Consolidate all Brian Lewis Jr records into Sterling Staffing',
  }).catch((error) => console.warn('[brian-sterling] merge log skipped:', error?.message || error));

  setStatus(`Done. Brian Lewis Jr is consolidated into one Sterling Staffing worker. ${duplicateProfiles.length} duplicate profile(s) were merged and ${changedRecords} related record(s) were relinked. No punch times were deleted.`);
  window.setTimeout(() => window.location.reload(), 1600);
}

function install() {
  const employeesTab = document.getElementById('employeesTab');
  if (!employeesTab || document.getElementById('brianSterlingConsolidationCard')) return;

  const card = document.createElement('div');
  card.id = 'brianSterlingConsolidationCard';
  card.className = 'card';
  card.innerHTML = `
    <div class="card-head split-head">
      <div>
        <h2>Brian Lewis Jr — Sterling Consolidation</h2>
        <p>Combines all OH01 Brian Lewis Jr employee identities into one Sterling Staffing worker while preserving every punch time and hour.</p>
      </div>
      <button id="consolidateBrianSterlingBtn" class="danger-btn" type="button">Combine Brian into Sterling</button>
    </div>
    <div id="brianSterlingStatus" class="status-box">Ready. This only runs after you click the button and confirm.</div>
  `;
  employeesTab.appendChild(card);

  document.getElementById('consolidateBrianSterlingBtn')?.addEventListener('click', async (event) => {
    const button = event.currentTarget;
    button.disabled = true;
    try {
      await consolidateBrianToSterling();
    } catch (error) {
      console.error('[brian-sterling]', error);
      setStatus(error?.message || 'Brian consolidation failed.', true);
      button.disabled = false;
    }
  });
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
else install();

console.info('[QRTimeclock] Brian Lewis Jr Sterling consolidation tool installed.');
