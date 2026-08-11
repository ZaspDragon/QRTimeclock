// Shared public worker lookup for the QR clock.
// Read-only: this module never creates, updates, deletes, merges, or migrates employee data.
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

function employeeName(row) {
  return String(row?.name || row?.employeeName || row?.displayName || row?.fullName || row?.nameKey || '').trim();
}

function employeeSite(row) {
  return String(row?.siteId || row?.assignedSiteId || row?.branch || row?.branchId || '').trim().toUpperCase();
}

function employeeAgency(row) {
  return String(row?.agencyId || row?.staffingAgencyId || '').trim();
}

function isActive(row) {
  return row?.active !== false && !['inactive', 'terminated', 'merged', 'deleted', 'removed', 'archived'].includes(String(row?.status || '').toLowerCase());
}

function identityValues(row) {
  return [
    row?.id,
    row?.employeeId,
    row?.employeeID,
    row?.employeeNumber,
    row?.workerId,
    row?.canonicalEmployeeId,
    row?.mergedInto,
    ...(Array.isArray(row?.legacyWorkerIds) ? row.legacyWorkerIds : []),
    ...(Array.isArray(row?.linkedWorkerIds) ? row.linkedWorkerIds : []),
    ...(Array.isArray(row?.aliases) ? row.aliases : []),
    ...(Array.isArray(row?.identityAliases) ? row.identityAliases : []),
  ].map((value) => String(value || '').trim()).filter(Boolean);
}

function sameIdentity(left, right) {
  const leftValues = new Set(identityValues(left));
  return identityValues(right).some((value) => leftValues.has(value));
}

export async function findPublicWorkerMatches(name, siteId, agencyId = '') {
  const db = dbInstance();
  const normalized = normalizeWorkerName(name);
  const key = workerNameKey(name);
  const site = String(siteId || '').toUpperCase();
  if (!normalized || !VALID_SITES.has(site)) return [];

  // Every query includes company + branch + active proof so it is compatible with
  // the public Firestore read rules. Name filtering remains exact in memory so
  // legacy rows whose nameKey format differs can still resolve.
  const searches = [
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('siteId', '==', site), where('active', '==', true), where('nameKey', '==', key), limit(30)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('siteId', '==', site), where('status', '==', 'active'), where('nameKey', '==', key), limit(30)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('assignedSiteId', '==', site), where('active', '==', true), where('nameKey', '==', key), limit(30)),
    query(collection(db, 'employees'), where('companyId', '==', COMPANY_ID), where('assignedSiteId', '==', site), where('status', '==', 'active'), where('nameKey', '==', key), limit(30)),
  ];

  const rows = new Map();
  const results = await Promise.allSettled(searches.map((ref) => getDocs(ref)));
  for (const result of results) {
    if (result.status !== 'fulfilled') continue;
    result.value.docs.forEach((snap) => rows.set(snap.id, { id: snap.id, ...snap.data() }));
  }

  return [...rows.values()].filter((row) => {
    if (!isActive(row)) return false;
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

  const canonicalSelf = matches.filter((row) => {
    const canonical = String(row.canonicalEmployeeId || '').trim();
    return canonical && identityValues(row).includes(canonical) && (canonical === row.id || canonical === row.employeeId || canonical === row.employeeNumber);
  });
  if (canonicalSelf.length === 1 && matches.every((row) => sameIdentity(row, canonicalSelf[0]) || row === canonicalSelf[0])) return canonicalSelf[0];

  const numbered = matches.filter((row) => /^EMP[-_ ]?\d+$/i.test(String(row.employeeNumber || row.employeeID || '')));
  if (numbered.length === 1 && matches.every((row) => sameIdentity(row, numbered[0]) || row === numbered[0])) return numbered[0];

  // If all profiles are linked aliases of one person, prefer the first active
  // canonical-looking profile. Otherwise return null so we never guess between
  // two genuinely different people with the same name.
  const seed = matches[0];
  if (matches.every((row) => sameIdentity(row, seed))) return seed;
  return null;
}
