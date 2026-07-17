import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createHash } from 'node:crypto';
import {
  assertFails,
  assertSucceeds,
  initializeTestEnvironment,
} from '@firebase/rules-unit-testing';
import {
  collection,
  doc,
  getDoc,
  getDocs,
  runTransaction,
  serverTimestamp,
  setDoc,
} from 'firebase/firestore';

const PROJECT_ID = 'qrtimeclock-rules-test';
const DUPLICATE_WINDOW_MS = 60 * 1000;
const baseMs = Date.parse('2026-07-17T08:00:00-04:00');

function hash(value) {
  return createHash('sha256').update(value).digest('hex');
}

function publicPunchGuardKey(payload) {
  const boundedWindow = Math.floor(Number(payload.timestampMs) / DUPLICATE_WINDOW_MS);
  return `public_${hash([
    payload.companyId,
    payload.siteId,
    payload.agencyId || '',
    payload.employeeId,
    payload.action,
    boundedWindow,
  ].join('|')).slice(0, 40)}`;
}

function publicPunchStateKey(payload) {
  return `public_${hash([
    payload.companyId,
    payload.siteId,
    payload.agencyId || '',
    payload.employeeId,
  ].join('|')).slice(0, 40)}`;
}

function payload(employeeId, action, siteId = 'OH01', timestampMs = baseMs) {
  return {
    companyId: 'chadwell',
    siteId,
    source: 'public_qr',
    name: 'Jordan Smith',
    nameKey: 'jordan_smith',
    action,
    timestamp: serverTimestamp(),
    timestampMs,
    dateKey: '2026-07-17',
    weekKey: '2026-07-13',
    employeeId,
    workerId: employeeId,
    employeeNumber: '100',
    agencyId: '',
    assignedSiteId: siteId,
    siteIds: [siteId],
    qrSlug: siteId.toLowerCase(),
    createdAt: serverTimestamp(),
  };
}

async function publicPunchTransaction(db, punchPayload) {
  const duplicateGuardKey = publicPunchGuardKey(punchPayload);
  const workerStateKey = publicPunchStateKey(punchPayload);
  const punchRef = doc(db, 'punches', duplicateGuardKey);
  const guardRef = doc(db, 'punchGuards', duplicateGuardKey);
  const stateRef = doc(db, 'punchStates', workerStateKey);
  await runTransaction(db, async (transaction) => {
    const guardSnap = await transaction.get(guardRef);
    if (guardSnap.exists()) return;
    await transaction.get(stateRef);
    transaction.set(punchRef, {
      ...punchPayload,
      duplicateGuardKey,
      idempotencyKey: duplicateGuardKey,
      workerStateKey,
    });
    transaction.set(guardRef, {
      duplicateGuardKey,
      punchId: duplicateGuardKey,
      companyId: punchPayload.companyId,
      siteId: punchPayload.siteId,
      agencyId: punchPayload.agencyId || '',
      employeeId: punchPayload.employeeId,
      action: punchPayload.action,
      acceptedAtMs: punchPayload.timestampMs,
      createdAt: serverTimestamp(),
      expiresAtMs: punchPayload.timestampMs + DUPLICATE_WINDOW_MS,
    });
    transaction.set(stateRef, {
      workerStateKey,
      companyId: punchPayload.companyId,
      siteId: punchPayload.siteId,
      agencyId: punchPayload.agencyId || '',
      employeeId: punchPayload.employeeId,
      lastAction: punchPayload.action,
      lastPunchAtMs: punchPayload.timestampMs,
      lastPunchId: duplicateGuardKey,
      updatedAt: serverTimestamp(),
    }, { merge: true });
  });
  return { punchId: duplicateGuardKey, stateId: workerStateKey };
}

const testEnv = await initializeTestEnvironment({
  projectId: PROJECT_ID,
  firestore: {
    rules: readFileSync(new URL('../firestore.rules', import.meta.url), 'utf8'),
  },
});

try {
  await testEnv.clearFirestore();
  await testEnv.withSecurityRulesDisabled(async (context) => {
    const adminDb = context.firestore();
    await setDoc(doc(adminDb, 'employees', 'emp_oh01'), {
      companyId: 'chadwell',
      siteId: 'OH01',
      status: 'active',
      active: true,
      name: 'Jordan Smith',
      nameKey: 'jordan_smith',
    });
    await setDoc(doc(adminDb, 'employees', 'emp_ohc'), {
      companyId: 'chadwell',
      assignedSiteId: 'OHC',
      active: true,
      name: 'Olivia Worker',
      nameKey: 'olivia_worker',
    });
    await setDoc(doc(adminDb, 'employees', 'emp_inactive'), {
      companyId: 'chadwell',
      siteId: 'OH01',
      status: 'inactive',
      active: false,
    });
  });

  const db = testEnv.unauthenticatedContext().firestore();

  const first = await assertSucceeds(publicPunchTransaction(db, payload('emp_oh01', 'clock_in')));
  assert.equal((await getDoc(doc(db, 'punches', first.punchId))).exists(), true, 'punch created');
  assert.equal((await getDoc(doc(db, 'punchGuards', first.punchId))).exists(), true, 'guard created');
  assert.equal((await getDoc(doc(db, 'punchStates', first.stateId))).exists(), true, 'state created');

  await assertSucceeds(publicPunchTransaction(db, payload('emp_oh01', 'clock_in')));

  const startLunch = await assertSucceeds(publicPunchTransaction(db, payload('emp_oh01', 'start_lunch', 'OH01', baseMs + 4 * 60 * 60 * 1000)));
  assert.equal(startLunch.stateId, first.stateId, 'same punchState updated on next action');
  await assertSucceeds(publicPunchTransaction(db, payload('emp_oh01', 'end_lunch', 'OH01', baseMs + 4.5 * 60 * 60 * 1000)));
  await assertSucceeds(publicPunchTransaction(db, payload('emp_oh01', 'clock_out', 'OH01', baseMs + 9 * 60 * 60 * 1000)));

  await assertSucceeds(publicPunchTransaction(db, payload('emp_ohc', 'clock_in', 'OHC', baseMs + 10 * 60 * 1000)));

  await assertFails(publicPunchTransaction(db, payload('emp_inactive', 'clock_in', 'OH01', baseMs + 20 * 60 * 1000)));
  await assertFails(publicPunchTransaction(db, payload('unknown', 'clock_in', 'OH01', baseMs + 30 * 60 * 1000)));
  await assertFails(publicPunchTransaction(db, { ...payload('emp_oh01', 'clock_in', 'OH01', baseMs + 40 * 60 * 1000), companyId: 'other' }));
  await assertFails(publicPunchTransaction(db, payload('emp_oh01', 'clock_in', 'BAD', baseMs + 50 * 60 * 1000)));
  await assertFails(publicPunchTransaction(db, payload('emp_oh01', 'bad_action', 'OH01', baseMs + 60 * 60 * 1000)));

  await assertFails(setDoc(doc(db, 'punches', first.punchId), { action: 'clock_out' }, { merge: true }));
  await assertFails(getDocs(collection(db, 'punchGuards')));
  await assertFails(getDocs(collection(db, 'punchStates')));
  await assertFails(getDoc(doc(db, 'users', 'manager')));
  await assertFails(getDoc(doc(db, 'payroll', 'summary')));
  await assertFails(getDoc(doc(db, 'approvals', 'approval_1')));
  await assertFails(getDoc(doc(db, 'auditLogs', 'audit_1')));
  await assertFails(getDoc(doc(db, 'timesheets', 'ts_1')));

  console.log('firestore public punch rules test passed');
} finally {
  await testEnv.cleanup();
}
