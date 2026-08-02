import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import {
  collection,
  getDocs,
  getFirestore,
  limit,
  query,
  where,
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const db = getFirestore(app);
let requestId = 0;

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function compactName(value) {
  return normalizeName(value).replaceAll(' ', '_');
}

function employeeName(row) {
  return String(row.name || row.employeeName || row.displayName || row.fullName || '').trim();
}

function isActive(row) {
  if (row.active === false) return false;
  return !['inactive', 'terminated', 'merged', 'deleted'].includes(String(row.status || '').toLowerCase());
}

function branchOf(row) {
  return String(row.siteId || row.assignedSiteId || row.branch || row.branchId || '').trim().toUpperCase();
}

function agencyOf(row) {
  return String(row.agencyId || row.staffingAgencyId || '').trim();
}

function setStatus(message, isError = false) {
  const status = document.getElementById('workerLookupStatus');
  if (!status) return;
  status.textContent = message;
  status.style.borderColor = isError ? 'rgba(255,90,90,.55)' : 'rgba(43,213,118,.45)';
}

function hideCreateNewSuggestion() {
  document.querySelectorAll('#workerAutocompleteList button, #workerAutocompleteList [role="option"]').forEach((node) => {
    if (/create new/i.test(node.textContent || '')) node.remove();
  });
}

async function findByExactName(name) {
  const normalized = normalizeName(name);
  const compact = compactName(name);
  if (normalized.length < 2) return [];

  const jobs = [
    query(collection(db, 'employees'), where('nameKey', '==', normalized), limit(20)),
    query(collection(db, 'employees'), where('nameKey', '==', compact), limit(20)),
    query(collection(db, 'employees'), where('normalizedName', '==', normalized), limit(20)),
    query(collection(db, 'employees'), where('name', '==', name.trim()), limit(20)),
  ];

  const rows = new Map();
  const results = await Promise.allSettled(jobs.map((job) => getDocs(job)));
  results.forEach((result) => {
    if (result.status !== 'fulfilled') return;
    result.value.docs.forEach((docSnap) => rows.set(docSnap.id, { id: docSnap.id, ...docSnap.data() }));
  });

  return [...rows.values()].filter((row) =>
    isActive(row) && normalizeName(employeeName(row) || row.nameKey || row.normalizedName) === normalized
  );
}

async function resolveTypedName() {
  const input = document.getElementById('workerNameInput');
  const typed = String(input?.value || '').trim();
  const currentRequest = ++requestId;
  if (typed.length < 2) return;

  window.setTimeout(async () => {
    if (currentRequest !== requestId) return;
    try {
      const matches = await findByExactName(typed);
      if (currentRequest !== requestId) return;
      hideCreateNewSuggestion();

      if (matches.length === 1) {
        const match = matches[0];
        const branch = branchOf(match);
        const agency = agencyOf(match);
        const branchSelect = document.getElementById('workerBranchSelect');
        const agencySelect = document.getElementById('workerAgencySelect');

        if (branch && branchSelect && [...branchSelect.options].some((option) => option.value === branch)) {
          branchSelect.disabled = false;
          branchSelect.value = branch;
          branchSelect.dispatchEvent(new Event('change', { bubbles: true }));
        }
        if (agency && agencySelect && [...agencySelect.options].some((option) => option.value === agency)) {
          agencySelect.value = agency;
          agencySelect.dispatchEvent(new Event('change', { bubbles: true }));
        }

        input.value = employeeName(match) || typed;
        localStorage.setItem('workerPunchName', input.value);
        setStatus(`✓ Found: ${input.value}. Ready to punch.`);
        return;
      }

      if (matches.length > 1) {
        setStatus('More than one worker has that exact name. Ask a manager to link the duplicate profiles.', true);
        return;
      }

      setStatus('No existing worker was found for that name.', true);
    } catch (error) {
      console.warn('[name-only-worker-resolver]', error);
      setStatus('Could not check that name. Try again.', true);
    }
  }, 350);
}

function install() {
  const input = document.getElementById('workerNameInput');
  if (!input || input.dataset.nameOnlyResolver === 'true') return;
  input.dataset.nameOnlyResolver = 'true';
  input.addEventListener('input', resolveTypedName);
  input.addEventListener('change', resolveTypedName);
  input.addEventListener('blur', resolveTypedName);
  if (input.value.trim()) resolveTypedName();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', install, { once: true });
} else {
  install();
}

console.info('[QRTimeclock] Name-only worker resolver installed.');
