/*
  Safe branch backfill utility for legacy QRTimeclock records.

  Default mode is dry-run and only prints what would change.
  To write non-destructive backfills, run with --apply after reviewing output.

  Required:
    GOOGLE_APPLICATION_CREDENTIALS must point to a Firebase service account JSON.

  Examples:
    node scripts/branch-backfill-preview.js
    node scripts/branch-backfill-preview.js --apply

  This never deletes documents. It only fills companyId/siteId when the legacy
  companyId can be mapped safely.
*/

const admin = require('firebase-admin');

const APPLY = process.argv.includes('--apply');
const COMPANY_ID = 'chadwell';
const LEGACY_COMPANY_TO_SITE = {
  chadwellOH01: 'OH01',
  chadwellOHC: 'OHC',
};
const COLLECTIONS = [
  'employees',
  'punches',
  'timesheets',
  'punch_edits',
  'auditLogs',
  'missedPunchRequests',
  'mergeLogs',
];

admin.initializeApp({
  credential: admin.credential.applicationDefault(),
});

const db = admin.firestore();

function inferSiteId(data) {
  if (data.siteId === 'OH01' || data.siteId === 'OHC') return data.siteId;
  return LEGACY_COMPANY_TO_SITE[data.companyId] || '';
}

async function flushBatch(batch, count) {
  if (!APPLY || count === 0) return;
  await batch.commit();
}

async function scanCollection(collectionName) {
  const snap = await db.collection(collectionName).get();
  let inspected = 0;
  let eligible = 0;
  let skipped = 0;
  let batch = db.batch();
  let batchCount = 0;

  for (const doc of snap.docs) {
    inspected += 1;
    const data = doc.data();
    const siteId = inferSiteId(data);
    const needsCompany = data.companyId !== COMPANY_ID && LEGACY_COMPANY_TO_SITE[data.companyId];
    const needsSite = !data.siteId && siteId;

    if (!needsCompany && !needsSite) continue;

    if (!siteId) {
      skipped += 1;
      continue;
    }

    eligible += 1;
    const update = {
      companyId: COMPANY_ID,
      siteId,
      branchBackfilledAt: admin.firestore.FieldValue.serverTimestamp(),
    };
    if (data.companyId && data.companyId !== COMPANY_ID) {
      update.legacyCompanyId = data.companyId;
    }

    console.log(`${APPLY ? 'UPDATE' : 'DRY-RUN'} ${collectionName}/${doc.id}`, update);

    if (APPLY) {
      batch.set(doc.ref, update, { merge: true });
      batchCount += 1;
      if (batchCount >= 450) {
        await flushBatch(batch, batchCount);
        batch = db.batch();
        batchCount = 0;
      }
    }
  }

  await flushBatch(batch, batchCount);
  console.log(`${collectionName}: inspected=${inspected} eligible=${eligible} skipped=${skipped}`);
}

async function main() {
  console.log(APPLY ? 'APPLY MODE: writing safe branch backfills' : 'DRY-RUN MODE: no writes');
  for (const collectionName of COLLECTIONS) {
    await scanCollection(collectionName);
  }
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
