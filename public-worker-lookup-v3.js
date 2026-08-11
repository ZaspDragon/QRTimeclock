// Shared read-only resolver for public clock/name lookup.
// Never deletes, merges, migrates, or rewrites employee/punch history.
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { collection, getDocs, getFirestore, limit, query, where } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

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

function dbInstance() {
  const app = getApps().length ? getApp() : initializeApp(FIREBASE_CONFIG);
  return getFirestore(app);
}

export function normalizeWorkerName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}

export function workerNameKey(value) {
  return normalizeWorkerName(value).replaceAll(' ', '_');
}

export function employeeName(row) {
  return String(row?.name || row?.employeeName || row?.displayName || row?.fullName || row?.nameKey || '').trim();
}

export function employeeSite(row) {
  return String(row?.siteId || row?.branch || '').trim().toUpperCase();
}

export function employeeAgency(row) {
  return String(row?.agencyId || row?.staffingAgencyId || '').trim();
}

export function isActiveWorker(row) {
  return row?.active !== false && !['inactive', 'terminated', 'merged', 'deleted', 'removed', 'archived'].includes(String(row?.status || '').toLowerCase());
}

export function identityValues(row) {
  return [row?.id,row?.employeeId,row?.employeeID,row?.employeeNumber,row?.workerId,row?.canonicalEmployeeId,row?.mergedInto,
    ...(Array.isArray(row?.legacyWorkerIds) ? row.legacyWorkerIds : []),
    ...(Array.isArray(row?.linkedWorkerIds) ? row.linkedWorkerIds : []),
    ...(Array.isArray(row?.aliases) ? row.aliases : []),
    ...(Array.isArray(row?.identityAliases) ? row.identityAliases : [])]
    .map((value) => String(value || '').trim()).filter(Boolean);
}

function identityOverlap(left, right) {
  const values = new Set(identityValues(left));
  return identityValues(right).some((value) => values.has(value));
}

export async function findPublicWorkerMatches(name, siteId, agencyId = '') {
  const db = dbInstance();
  const normalized = normalizeWorkerName(name);
  const key = workerNameKey(name);
  const site = String(siteId || '').trim().toUpperCase();
  if (!normalized || !VALID_SITES.has(site)) return [];

  // Every list query proves company, valid branch and active state to Firestore.
  // Exact-name filtering keeps reads small and avoids the old 500-row public scans.
  const searches = [
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('siteId', '==', site), where('active', '==', true), where('nameKey', '==', key), limit(20)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('siteId', '==', site), where('status', '==', 'active'), where('nameKey', '==', key), limit(20)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('branch', '==', site), where('active', '==', true), where('nameKey', '==', key), limit(20)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('branch', '==', site), where('status', '==', 'active'), where('nameKey', '==', key), limit(20)),
  ];

  const rows = new Map();
  const results = await Promise.allSettled(searches.map((ref) => getDocs(ref)));
  for (const result of results) {
    if (result.status !== 'fulfilled') continue;
    result.value.docs.forEach((snap) => rows.set(snap.id, { id: snap.id, ...snap.data() }));
  }

  return [...rows.values()].filter((row) => {
    if (!isActiveWorker(row)) return false;
    if (String(row.companyId || COMPANY_ID) !== COMPANY_ID) return false;
    const rowSite = employeeSite(row);
    if (rowSite && rowSite !== site) return false;
    if (normalizeWorkerName(employeeName(row)) !== normalized) return false;
    if (agencyId && employeeAgency(row) && employeeAgency(row) !== agencyId) return false;
    return true;
  });
}

export function chooseCanonicalPublicWorker(matches) {
  if (!matches.length) return null;
  if (matches.length === 1) return matches[0];

  const canonical = matches.filter((row) => {
    const canonicalId = String(row.canonicalEmployeeId || '').trim();
    return canonicalId && identityValues(row).includes(canonicalId)
      && (canonicalId === row.id || canonicalId === row.employeeId || canonicalId === row.employeeNumber);
  });
  if (canonical.length === 1 && matches.every((row) => row === canonical[0] || identityOverlap(row, canonical[0]))) return canonical[0];

  const numbered = matches.filter((row) => /^EMP[-_ ]?\d+$/i.test(String(row.employeeNumber || row.employeeID || '')));
  if (numbered.length === 1 && matches.every((row) => row === numbered[0] || identityOverlap(row, numbered[0]))) return numbered[0];

  const seed = matches[0];
  if (matches.every((row) => identityOverlap(row, seed))) return seed;
  return null;
}
