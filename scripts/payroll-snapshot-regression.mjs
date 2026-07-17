import fs from 'node:fs';
import path from 'node:path';

const backupDir = process.env.QRTIMECLOCK_BACKUP_DIR
  || path.join(process.env.USERPROFILE || '', 'Downloads', 'qrtimeclock-firestore-backups');

const beforePath = path.join(backupDir, 'qrtimeclock-firestore-readonly-snapshot-2026-06-24T10-29-33-660Z.json');
const afterPath = path.join(backupDir, 'qrtimeclock-firestore-post-backfill-snapshot-2026-06-24T10-31-02-853Z.json');

const requiredCollections = [
  'employees',
  'punches',
  'timesheets',
  'missedPunchRequests',
  'punch_edits',
  'auditLogs',
  'mergeLogs',
  'users',
];

function readSnapshot(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`Missing snapshot: ${filePath}`);
  }
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function collectionRows(snapshot, collectionName) {
  return Array.isArray(snapshot.collections?.[collectionName])
    ? snapshot.collections[collectionName]
    : [];
}

function activePunch(row) {
  const data = row.data || {};
  return String(data.status || 'active').toLowerCase() !== 'deleted' && data.active !== false;
}

const before = readSnapshot(beforePath);
const after = readSnapshot(afterPath);
const failures = [];
const warnings = [];

for (const collectionName of requiredCollections) {
  const beforeRows = collectionRows(before, collectionName);
  const afterRows = collectionRows(after, collectionName);
  const afterIds = new Set(afterRows.map((row) => row.id));
  const missing = beforeRows.filter((row) => !afterIds.has(row.id));
  if (missing.length) {
    failures.push(`${collectionName}: ${missing.length} pre-existing document(s) missing after backfill`);
  }
}

const afterPunches = collectionRows(after, 'punches');
const invalidPunches = afterPunches.filter((row) => {
  if (!activePunch(row)) return false;
  const data = row.data || {};
  return !['clock_in', 'start_lunch', 'end_lunch', 'clock_out'].includes(data.action)
    || !Number.isFinite(Number(data.timestampMs))
    || Number(data.timestampMs) <= 0
    || !data.dateKey
    || !data.weekKey
    || !data.companyId
    || !data.siteId;
});

if (invalidPunches.length) {
  failures.push(`punches: ${invalidPunches.length} active punch(es) have invalid payroll fields`);
}

const signatureCounts = new Map();
for (const row of afterPunches.filter(activePunch)) {
  const data = row.data || {};
  const signature = [
    data.employeeId || data.workerId || data.nameKey || data.name || '',
    data.action || '',
    data.timestampMs || '',
    data.siteId || '',
    data.agencyId || '',
  ].join('|');
  signatureCounts.set(signature, (signatureCounts.get(signature) || 0) + 1);
}
const duplicateSignatures = [...signatureCounts.values()].filter((count) => count > 1).length;
if (duplicateSignatures) {
  warnings.push(`snapshot contains ${duplicateSignatures} duplicate-looking active punch signature(s); current app validation prevents new malformed saves but does not rewrite history`);
}

if (warnings.length) {
  console.warn(warnings.map((warning) => `WARN: ${warning}`).join('\n'));
}

if (failures.length) {
  console.error(failures.map((failure) => `FAIL: ${failure}`).join('\n'));
  process.exit(1);
}

console.log('payroll snapshot regression passed');
