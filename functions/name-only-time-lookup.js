'use strict';

const { onRequest } = require('firebase-functions/v2/https');
const { getApps, initializeApp } = require('firebase-admin/app');
const { getFirestore } = require('firebase-admin/firestore');

if (getApps().length === 0) initializeApp();
const db = getFirestore();

const COMPANY_ID = 'chadwell';
const VALID_SITES = new Set(['OH01', 'OHC']);
const VALID_ACTIONS = new Set(['clock_in', 'start_lunch', 'end_lunch', 'clock_out']);
const ALLOWED_ORIGINS = new Set([
  'https://zaspdragon.github.io',
  'https://qrtimeclock-42764.web.app',
  'https://qrtimeclock-42764.firebaseapp.com',
]);
const MAX_RANGE_DAYS = 31;
const MAX_ROWS = 500;
const RATE_WINDOW_MS = 60_000;
const RATE_LIMIT = 30;
const rateBuckets = new Map();

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
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
  return String(req.headers['x-forwarded-for'] || req.ip || 'unknown')
    .split(',')[0]
    .trim();
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
  if (value === undefined || value === null || value === '') return [];
  try {
    const snapshot = await db
      .collection(collectionName)
      .where(field, '==', value)
      .limit(MAX_ROWS)
      .get();
    return snapshot.docs.map((document) => ({
      id: document.id,
      collectionName,
      ...document.data(),
    }));
  } catch (error) {
    console.warn(`[publicWorkerTimeLookupByName] ${collectionName}.${field} query failed`, error.message);
    return [];
  }
}

function uniqueRows(rows) {
  const result = new Map();
  rows.forEach((row) => {
    const key = `${row.collectionName || 'unknown'}:${row.id}`;
    if (!result.has(key)) result.set(key, row);
  });
  return [...result.values()];
}

function siteValues(data) {
  return [
    data?.siteId,
    data?.assignedSiteId,
    data?.branch,
    ...(Array.isArray(data?.siteIds) ? data.siteIds : []),
    ...(Array.isArray(data?.branches) ? data.branches : []),
  ]
    .map((value) => String(value || '').trim().toUpperCase())
    .filter(Boolean);
}

function siteMatches(data, siteId) {
  const values = siteValues(data);
  return values.length === 0 || values.includes(siteId);
}

function profileName(profile) {
  return profile?.name
    || profile?.employeeName
    || profile?.displayName
    || profile?.fullName
    || profile?.nameKey
    || profile?.normalizedName
    || '';
}

function profileIsActive(profile) {
  if (profile?.active === false) return false;
  const status = String(profile?.status || '').trim().toLowerCase();
  return !['inactive', 'terminated', 'merged', 'deleted'].includes(status);
}

async function resolveProfiles(name, siteId) {
  const normalized = normalizeName(name);
  const compact = compactNameKey(name);
  const exactNames = [...new Set([
    String(name || '').trim(),
    String(name || '').trim().toLowerCase(),
  ].filter(Boolean))];
  const jobs = [];

  for (const collectionName of ['employees', 'workers']) {
    for (const key of [normalized, compact]) {
      jobs.push(runQuery(collectionName, 'nameKey', key));
      jobs.push(runQuery(collectionName, 'normalizedName', key));
    }
    for (const exactName of exactNames) {
      jobs.push(runQuery(collectionName, 'name', exactName));
      jobs.push(runQuery(collectionName, 'employeeName', exactName));
    }
  }

  return uniqueRows((await Promise.all(jobs)).flat()).filter((profile) => {
    const sameName = normalizeName(profileName(profile)) === normalized;
    const sameCompany = !profile.companyId
      || String(profile.companyId).trim().toLowerCase() === COMPANY_ID;
    return sameName && sameCompany && profileIsActive(profile) && siteMatches(profile, siteId);
  });
}

function identitySet(profiles, fieldNames) {
  const values = [];
  profiles.forEach((profile) => {
    fieldNames.forEach((fieldName) => {
      const value = profile?.[fieldName];
      if (Array.isArray(value)) values.push(...value);
      else values.push(value);
    });
  });
  return [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))];
}

async function loadPunches(name, siteId, profiles, fromMs, toMs) {
  const normalized = normalizeName(name);
  const compact = compactNameKey(name);
  const ids = identitySet(profiles, [
    'id',
    'canonicalEmployeeId',
    'employeeId',
    'employeeID',
    'workerId',
    'linkedWorkerIds',
    'aliases',
    'legacyWorkerIds',
    'identityAliases',
  ]);
  const employeeNumbers = identitySet(profiles, [
    'employeeNumber',
    'employeeNo',
    'employeeID',
  ]);
  const knownNames = [...new Set([
    String(name || '').trim(),
    ...profiles.map(profileName),
  ].map((value) => String(value || '').trim()).filter(Boolean))];
  const jobs = [];

  ids.forEach((id) => {
    jobs.push(runQuery('punches', 'employeeId', id));
    jobs.push(runQuery('punches', 'workerId', id));
  });
  employeeNumbers.forEach((number) => {
    jobs.push(runQuery('punches', 'employeeNumber', number));
  });
  jobs.push(runQuery('punches', 'nameKey', normalized));
  jobs.push(runQuery('punches', 'nameKey', compact));
  jobs.push(runQuery('punches', 'normalizedName', normalized));
  knownNames.forEach((knownName) => jobs.push(runQuery('punches', 'name', knownName)));

  const rawRows = uniqueRows((await Promise.all(jobs)).flat());
  return rawRows
    .map((row) => ({ ...row, timestampMs: timestampMs(row) }))
    .filter((row) => row.timestampMs >= fromMs && row.timestampMs <= toMs)
    .filter((row) => row.status !== 'deleted' && row.active !== false)
    .filter((row) => !row.companyId || String(row.companyId).toLowerCase() === COMPANY_ID)
    .filter((row) => siteMatches(row, siteId))
    .filter((row) => VALID_ACTIONS.has(row.action))
    .filter((row) => {
      const rowIds = [row.employeeId, row.employeeID, row.workerId]
        .map((value) => String(value || '').trim())
        .filter(Boolean);
      const idMatch = rowIds.some((id) => ids.includes(id));
      const numberMatch = employeeNumbers.includes(String(row.employeeNumber || '').trim());
      const exactNameMatch = normalizeName(row.name || row.employeeName || row.nameKey) === normalized;
      return idMatch || numberMatch || exactNameMatch;
    })
    .sort((left, right) => left.timestampMs - right.timestampMs)
    .slice(0, MAX_ROWS)
    .map((row) => ({
      action: row.action,
      timestampMs: row.timestampMs,
      dateKey: row.dateKey || new Date(row.timestampMs).toISOString().slice(0, 10),
    }));
}

exports.publicWorkerTimeLookupByName = onRequest(
  { region: 'us-central1', cors: false },
  async (req, res) => {
    setCors(req, res);
    if (req.method === 'OPTIONS') return res.status(204).send('');
    if (req.method !== 'POST') return res.status(405).json({ error: 'POST required.' });

    const origin = String(req.headers.origin || '');
    if (origin && !ALLOWED_ORIGINS.has(origin)) {
      return res.status(403).json({ error: 'Origin not allowed.' });
    }
    if (!rateAllowed(req)) {
      return res.status(429).json({ error: 'Too many lookup attempts. Try again shortly.' });
    }

    try {
      const name = String(req.body?.name || '').trim();
      const siteId = String(req.body?.siteId || '').trim().toUpperCase();
      const fromMs = Number(req.body?.fromMs);
      const toMs = Number(req.body?.toMs);

      if (name.length < 2 || name.length > 80) {
        return res.status(400).json({ error: 'Enter the worker name.' });
      }
      if (!VALID_SITES.has(siteId)) {
        return res.status(400).json({ error: 'Choose a valid branch.' });
      }
      if (!Number.isFinite(fromMs) || !Number.isFinite(toMs) || fromMs > toMs) {
        return res.status(400).json({ error: 'Choose a valid date range.' });
      }
      if (toMs - fromMs > MAX_RANGE_DAYS * 86_400_000) {
        return res.status(400).json({
          error: `Time lookup is limited to ${MAX_RANGE_DAYS} days at a time.`,
        });
      }

      const profiles = await resolveProfiles(name, siteId);
      const punches = await loadPunches(name, siteId, profiles, fromMs, toMs);
      if (profiles.length === 0 && punches.length === 0) {
        return res.status(404).json({
          error: 'No saved time was found for that exact name in this branch.',
        });
      }

      const displayName = profiles.map(profileName).find(Boolean) || name;
      return res.status(200).json({
        worker: {
          name: displayName,
          siteId,
          matchedProfiles: profiles.length,
        },
        punches,
      });
    } catch (error) {
      console.error('[publicWorkerTimeLookupByName]', error);
      return res.status(500).json({ error: error.message || 'Time lookup failed.' });
    }
  },
);
