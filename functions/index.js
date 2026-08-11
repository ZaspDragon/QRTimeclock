'use strict';

const { onRequest } = require('firebase-functions/v2/https');
const { initializeApp } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');

initializeApp();
const db = getFirestore();

const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_AGENCIES = new Set(['sterling_staffing', 'excel_staffing', 'lifestyle_staffing']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ALLOWED_ORIGINS = new Set([
  'https://zaspdragon.github.io',
  'https://qrtimeclock-42764.web.app',
  'https://qrtimeclock-42764.firebaseapp.com',
]);
const MAX_RANGE_DAYS = 31;
const MAX_ROWS = 400;
const RATE_WINDOW_MS = 60_000;
const RATE_LIMIT = 20;
const rateBuckets = new Map();

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').replace(/\s+/g, ' ').trim();
}
function compactNameKey(value) {
  return normalizeName(value).replaceAll(' ', '_');
}
function timestampMs(data) {
  if (Number.isFinite(data?.timestampMs)) return Number(data.timestampMs);
  if (data?.timestamp?.toMillis) return data.timestamp.toMillis();
  if (Number.isFinite(data?.timestamp?.seconds)) return data.timestamp.seconds * 1000;
  return 0;
}
function setCors(req, res) {
  const origin = String(req.headers.origin || '');
  if (ALLOWED_ORIGINS.has(origin)) res.set('Access-Control-Allow-Origin', origin);
  res.set('Vary', 'Origin');
  res.set('Access-Control-Allow-Headers', 'Content-Type');
  res.set('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.set('Cache-Control', 'no-store, max-age=0');
  res.set('X-Content-Type-Options', 'nosniff');
}
function clientKey(req) {
  return String(req.headers['x-forwarded-for'] || req.ip || 'unknown').split(',')[0].trim();
}
function rateAllowed(req) {
  const key = clientKey(req);
  const now = Date.now();
  const current = rateBuckets.get(key);
  if (!current || now - current.startedAt >= RATE_WINDOW_MS) {
    rateBuckets.set(key, { startedAt: now, count: 1 });
    return true;
  }
  current.count += 1;
  return current.count <= RATE_LIMIT;
}
async function runQuery(collectionName, field, value) {
  const snapshot = await db.collection(collectionName).where(field, '==', value).limit(MAX_ROWS).get();
  return snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
}
function uniqueRows(rows) {
  const result = new Map();
  for (const row of rows) if (!result.has(row.id)) result.set(row.id, row);
  return [...result.values()];
}
function workerSiteMatches(worker, siteId) {
  const sites = [worker.siteId, worker.assignedSiteId, worker.branch, ...(Array.isArray(worker.siteIds) ? worker.siteIds : [])]
    .filter(Boolean);
  return sites.length === 0 || sites.includes(siteId);
}
function workerAgencyId(worker) {
  return String(worker?.agencyId || worker?.staffingAgencyId || '').trim();
}
function workerActive(worker) {
  return worker?.active === true || String(worker?.status || '').toLowerCase() === 'active';
}
function workerIdentityValues(worker) {
  return [
    worker?.id,
    worker?.canonicalEmployeeId,
    worker?.employeeId,
    worker?.employeeID,
    worker?.workerId,
    worker?.employeeNumber,
    worker?.mergedInto,
    ...(Array.isArray(worker?.linkedWorkerIds) ? worker.linkedWorkerIds : []),
    ...(Array.isArray(worker?.aliases) ? worker.aliases : []),
    ...(Array.isArray(worker?.legacyWorkerIds) ? worker.legacyWorkerIds : []),
    ...(Array.isArray(worker?.identityAliases) ? worker.identityAliases : []),
  ].map((value) => String(value || '').trim()).filter(Boolean);
}
function identitiesOverlap(left, right) {
  const values = new Set(workerIdentityValues(left));
  return workerIdentityValues(right).some((value) => values.has(value));
}
function chooseCanonicalWorker(matches) {
  if (matches.length === 1) return matches[0];
  if (!matches.length) return null;

  const canonicalSelf = matches.filter((worker) => {
    const canonicalId = String(worker.canonicalEmployeeId || '').trim();
    return canonicalId && (canonicalId === worker.id || canonicalId === worker.employeeId || canonicalId === worker.employeeNumber);
  });
  if (canonicalSelf.length === 1 && matches.every((worker) => worker === canonicalSelf[0] || identitiesOverlap(worker, canonicalSelf[0]))) {
    return canonicalSelf[0];
  }

  const numbered = matches.filter((worker) => /^EMP[-_ ]?\d+$/i.test(String(worker.employeeNumber || worker.employeeID || '')));
  if (numbered.length === 1 && matches.every((worker) => worker === numbered[0] || identitiesOverlap(worker, numbered[0]))) {
    return numbered[0];
  }

  const seed = matches[0];
  if (matches.every((worker) => identitiesOverlap(worker, seed))) return seed;
  return null;
}
async function validateWorkerById(employeeId, name, siteId, agencyId) {
  const id = String(employeeId || '').trim();
  if (!id) return null;
  const snap = await db.collection('employees').doc(id).get();
  if (!snap.exists) return null;
  const worker = { id: snap.id, ...snap.data() };
  const workerName = normalizeName(worker.name || worker.employeeName || worker.nameKey || worker.normalizedName);
  const companyMatches = !worker.companyId || worker.companyId === COMPANY_ID;
  const agencyMatches = !workerAgencyId(worker) || workerAgencyId(worker) === agencyId;
  if (!workerActive(worker) || !companyMatches || !workerSiteMatches(worker, siteId) || !agencyMatches) return null;
  if (workerName && workerName !== normalizeName(name)) return null;
  return worker;
}
async function resolveWorker(name, siteId, agencyId, employeeId) {
  const direct = await validateWorkerById(employeeId, name, siteId, agencyId);
  if (direct) return direct;

  const normalized = normalizeName(name);
  const keys = [...new Set([normalized, compactNameKey(name)])];
  const jobs = [];
  for (const key of keys) {
    jobs.push(runQuery('employees', 'nameKey', key));
    jobs.push(runQuery('employees', 'normalizedName', key));
  }

  const matches = uniqueRows((await Promise.all(jobs)).flat()).filter((worker) => {
    const workerName = normalizeName(worker.name || worker.employeeName || worker.nameKey || worker.normalizedName);
    const companyMatches = !worker.companyId || worker.companyId === COMPANY_ID;
    const agencyMatches = !workerAgencyId(worker) || workerAgencyId(worker) === agencyId;
    return workerName === normalized
      && companyMatches
      && workerActive(worker)
      && workerSiteMatches(worker, siteId)
      && agencyMatches;
  });

  if (matches.length === 0) {
    const error = new Error('No active worker matched that exact name, branch, and staffing agency.');
    error.status = 404;
    throw error;
  }

  const worker = chooseCanonicalWorker(matches);
  if (!worker) {
    const error = new Error('More than one separate active profile matched. Ask a manager to link the duplicate profiles.');
    error.status = 409;
    throw error;
  }
  return worker;
}

async function loadPunches(worker, siteId, agencyId, fromMs, toMs) {
  const ids = [...new Set(workerIdentityValues(worker))];
  const employeeNumbers = [...new Set([worker.employeeNumber, worker.employeeNo]
    .map((value) => String(value || '').trim()).filter(Boolean))];
  const jobs = [];

  for (const id of ids) {
    jobs.push(runQuery('punches', 'employeeId', id));
    jobs.push(runQuery('punches', 'workerId', id));
    jobs.push(runQuery('punches', 'canonicalEmployeeId', id));
  }
  for (const number of employeeNumbers) jobs.push(runQuery('punches', 'employeeNumber', number));
  jobs.push(runQuery('punches', 'nameKey', normalizeName(worker.name || worker.nameKey)));
  jobs.push(runQuery('punches', 'nameKey', compactNameKey(worker.name || worker.nameKey)));

  return uniqueRows((await Promise.all(jobs)).flat())
    .map((row) => ({ ...row, timestampMs: timestampMs(row) }))
    .filter((row) => row.timestampMs >= fromMs && row.timestampMs <= toMs)
    .filter((row) => row.status !== 'deleted' && row.active !== false)
    .filter((row) => !row.companyId || row.companyId === COMPANY_ID)
    .filter((row) => {
      const rowSite = row.siteId || row.assignedSiteId || row.branch || '';
      return !rowSite || rowSite === siteId;
    })
    .filter((row) => {
      const rowAgency = String(row.agencyId || '').trim();
      return !rowAgency || rowAgency === agencyId;
    })
    .filter((row) => VALID_ACTIONS.has(row.action))
    .filter((row) => {
      const rowIds = [row.employeeId, row.workerId, row.canonicalEmployeeId]
        .map((value) => String(value || '').trim()).filter(Boolean);
      const idMatch = rowIds.some((id) => ids.includes(id));
      const numberMatch = employeeNumbers.includes(String(row.employeeNumber || '').trim());
      const legacyNameOnly = rowIds.length === 0
        && !row.employeeNumber
        && normalizeName(row.name || row.nameKey) === normalizeName(worker.name || worker.nameKey);
      return idMatch || numberMatch || legacyNameOnly;
    })
    .sort((left, right) => left.timestampMs - right.timestampMs)
    .slice(0, MAX_ROWS)
    .map((row) => ({
      action: row.action,
      timestampMs: row.timestampMs,
      dateKey: row.dateKey || new Date(row.timestampMs).toISOString().slice(0, 10),
    }));
}

exports.publicWorkerTimeLookup = onRequest({ region: 'us-central1', cors: false }, async (req, res) => {
  setCors(req, res);
  if (req.method === 'OPTIONS') return res.status(204).send('');
  if (req.method !== 'POST') return res.status(405).json({ error: 'POST required.' });

  const origin = String(req.headers.origin || '');
  if (origin && !ALLOWED_ORIGINS.has(origin)) return res.status(403).json({ error: 'Origin not allowed.' });
  if (!rateAllowed(req)) return res.status(429).json({ error: 'Too many lookup attempts. Try again shortly.' });

  try {
    const name = String(req.body?.name || '').trim();
    const siteId = String(req.body?.siteId || '').trim().toUpperCase();
    const agencyId = String(req.body?.agencyId || '').trim();
    const employeeId = String(req.body?.employeeId || '').trim();
    const fromMs = Number(req.body?.fromMs);
    const toMs = Number(req.body?.toMs);

    if (name.length < 2 || name.length > 80) return res.status(400).json({ error: 'Enter your first and last name.' });
    if (!VALID_SITES.has(siteId)) return res.status(400).json({ error: 'Choose a valid branch.' });
    if (!VALID_AGENCIES.has(agencyId)) return res.status(400).json({ error: 'Choose your staffing agency.' });
    if (!Number.isFinite(fromMs) || !Number.isFinite(toMs) || fromMs > toMs) {
      return res.status(400).json({ error: 'Choose a valid date range.' });
    }
    if (toMs - fromMs > MAX_RANGE_DAYS * 86_400_000) {
      return res.status(400).json({ error: `Time lookup is limited to ${MAX_RANGE_DAYS} days at a time.` });
    }

    const worker = await resolveWorker(name, siteId, agencyId, employeeId);
    const punches = await loadPunches(worker, siteId, agencyId, fromMs, toMs);
    return res.status(200).json({
      worker: {
        name: worker.name || name,
        siteId,
        employeeId: worker.employeeId || worker.id || '',
      },
      punches,
    });
  } catch (error) {
    console.error('[publicWorkerTimeLookup]', error);
    return res.status(error.status || 500).json({ error: error.message || 'Time lookup failed.' });
  }
});
