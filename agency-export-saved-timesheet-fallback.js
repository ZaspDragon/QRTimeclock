import { firebaseConfig } from './firebase-config.js';
import { initializeApp, getApps } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { getAuth, onAuthStateChanged } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  collection,
  doc,
  getDoc,
  getDocs,
  getFirestore,
  query,
  where
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApps()[0] : initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

const CURRENT_COMPANY_ID = 'chadwell';
const CURRENT_SITE_ID = 'OH01';
const BRANCH_OPTIONS = ['OH01', 'OHC'];
const FALLBACK_PREFIX = 'saved-timesheet:';
const AGENCY_NAMES = {
  sterling_staffing: 'Sterling Staffing',
  excel_staffing: 'Excel Staffing',
};

const state = {
  profile: null,
  savedSheets: new Map(),
  refreshTimer: null,
};

function normalizeRole(value) {
  return String(value || '').trim().toLowerCase().replace(/[\s-]+/g, '_');
}

function normalizeName(value) {
  return String(value || '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .replace(/[^a-z0-9 ]/g, '')
    .replaceAll(' ', '_');
}

function prettifyHumanName(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase()
    .replace(/\b\w/g, (char) => char.toUpperCase());
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function agencyLabel(agencyId) {
  if (!agencyId) return 'Direct';
  return AGENCY_NAMES[agencyId] || agencyId;
}

function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function getMondayDate(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

function selectedWeekKey() {
  const value = document.getElementById('weekPicker')?.value;
  if (!value) return formatDateKey(getMondayDate(new Date()));
  return formatDateKey(getMondayDate(new Date(`${value}T00:00:00`)));
}

function parseSiteIds(value) {
  const raw = Array.isArray(value) ? value : String(value || '').split(',');
  return [...new Set(raw.map((site) => String(site || '').trim()).filter((site) => BRANCH_OPTIONS.includes(site)))];
}

function allowedSiteIds(profile = state.profile) {
  if (!profile) return [];
  if (Array.isArray(profile.branches) && profile.branches.length) return parseSiteIds(profile.branches);
  if (profile.branch) return parseSiteIds([profile.branch]);
  if (Array.isArray(profile.siteIds) && profile.siteIds.length) return parseSiteIds(profile.siteIds);
  return parseSiteIds(profile.siteId || profile.assignedSiteId || '');
}

function activeSiteId(profile = state.profile) {
  const allowed = allowedSiteIds(profile);
  const stored = sessionStorage.getItem(`managerActiveBranch:${profile?.uid || ''}`);
  return allowed.includes(stored) ? stored : allowed[0] || CURRENT_SITE_ID;
}

function isAgencyUser(profile = state.profile) {
  return normalizeRole(profile?.role) === 'agency_admin' || !!profile?.agencyId;
}

function sheetName(sheet) {
  return prettifyHumanName(sheet.name || sheet.workerName || sheet.employeeName || sheet.displayName || '');
}

function sheetSignature(sheet) {
  return [
    normalizeName(sheetName(sheet) || sheet.nameKey),
    String(sheet.agencyId || '').trim(),
    String(sheet.siteId || sheet.branchId || '').trim()
  ].join('|');
}

function existingOptionSignatures(select) {
  const signatures = new Set();
  Array.from(select.options || []).forEach((option) => {
    if (!option.value || option.value.startsWith(FALLBACK_PREFIX)) return;
    const text = option.textContent || '';
    const name = text.replace(/\s*\([^)]*\)\s*$/, '');
    const details = (text.match(/\(([^)]*)\)/)?.[1] || '').split('·').map((part) => part.trim());
    signatures.add([normalizeName(name), '', details[1] || details[0] || ''].join('|'));
  });
  return signatures;
}

function addFallbackOptions() {
  const select = document.getElementById('agencyWorkerSelect');
  if (!select || !state.savedSheets.size) return;

  const existingValues = new Set(Array.from(select.options || []).map((option) => option.value));
  const signatures = existingOptionSignatures(select);
  const rows = [...state.savedSheets.values()].sort((left, right) => sheetName(left).localeCompare(sheetName(right)));

  rows.forEach((sheet) => {
    const value = `${FALLBACK_PREFIX}${sheet.id}`;
    if (existingValues.has(value) || signatures.has(sheetSignature(sheet))) return;
    const option = document.createElement('option');
    option.value = value;
    const details = [agencyLabel(sheet.agencyId), sheet.siteId || sheet.branchId].filter(Boolean).join(' · ');
    option.textContent = `${sheetName(sheet) || sheet.nameKey || sheet.id}${details ? ` (${details})` : ''}`;
    option.dataset.savedTimesheetFallback = 'true';
    select.appendChild(option);
  });
}

async function loadSavedSheets() {
  if (!auth.currentUser || !state.profile) return;
  const siteId = activeSiteId();
  const filters = [
    where('companyId', '==', CURRENT_COMPANY_ID),
    where('siteId', '==', siteId),
    where('weekKey', '==', selectedWeekKey()),
  ];
  if (isAgencyUser() && state.profile.agencyId) {
    filters.push(where('agencyId', '==', state.profile.agencyId));
  }

  const snap = await getDocs(query(collection(db, 'timesheets'), ...filters));
  state.savedSheets = new Map(snap.docs.map((record) => [record.id, { id: record.id, ...record.data() }]));
  addFallbackOptions();
}

function scheduleRefresh() {
  window.clearTimeout(state.refreshTimer);
  state.refreshTimer = window.setTimeout(() => {
    loadSavedSheets().catch((error) => {
      console.warn('Agency export saved-timesheet fallback failed:', error.message);
    });
  }, 250);
}

function formatDateTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleString([], {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  });
}

function buildDailyRows(dailyTotals) {
  const keys = Object.keys(dailyTotals || {}).sort();
  if (!keys.length) {
    return '<tr><td colspan="6" style="border:1px solid #bbb;padding:10px;">No detailed daily punches are stored on this saved sheet.</td></tr>';
  }

  return keys.map((dateKey) => {
    const row = dailyTotals[dateKey] || {};
    return `
      <tr>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(dateKey)}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.clock_in || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.start_lunch || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.end_lunch || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(row.clock_out || '-')}</td>
        <td style="border:1px solid #bbb;padding:10px;">${Number(row.hours || 0).toFixed(2)}</td>
      </tr>
    `;
  }).join('');
}

function renderSavedSheet(sheet) {
  const preview = document.getElementById('agencyPreview');
  if (!preview) return;
  const signedAt = sheet.managerSignedAt?.seconds ? formatDateTime(sheet.managerSignedAt.seconds * 1000) : '-';
  const dailyTotals = sheet.dailyTotals && typeof sheet.dailyTotals === 'object' ? sheet.dailyTotals : {};

  preview.innerHTML = `
    <div id="agencyPrintableSheet" style="background:#fff;color:#111;border-radius:12px;padding:24px;min-height:200px;">
      <div style="display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap;margin-bottom:18px;">
        <div>
          <h2 style="margin:0 0 8px;font-size:28px;">Weekly Time Sheet</h2>
          <div style="font-size:15px;line-height:1.6;">
            <div><strong>Worker:</strong> ${escapeHtml(sheetName(sheet) || sheet.nameKey || '-')}</div>
            <div><strong>Agency:</strong> ${escapeHtml(agencyLabel(sheet.agencyId))}</div>
            <div><strong>Branch:</strong> ${escapeHtml(sheet.siteId || sheet.branchId || '-')}</div>
            <div><strong>Week Start:</strong> ${escapeHtml(sheet.weekKey || selectedWeekKey())}</div>
            <div><strong>Status:</strong> ${escapeHtml(sheet.status || 'open')}</div>
          </div>
        </div>
        <div style="font-size:14px;line-height:1.7;text-align:right;">
          <div><strong>Company:</strong> Chadwell</div>
          <div><strong>Generated:</strong> ${escapeHtml(formatDateTime(Date.now()))}</div>
        </div>
      </div>
      <table style="width:100%;border-collapse:collapse;margin-bottom:18px;">
        <thead>
          <tr>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Date</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Clock In</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Lunch Out</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Lunch In</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Clock Out</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Hours</th>
          </tr>
        </thead>
        <tbody>${buildDailyRows(dailyTotals)}</tbody>
      </table>
      <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-top:24px;">
        <div style="font-size:15px;line-height:1.8;">
          <div><strong>Total Hours:</strong> ${Number(sheet.weeklyHours || 0).toFixed(2)}</div>
          <div><strong>Days Worked:</strong> ${Number(sheet.daysWorked || Object.keys(dailyTotals).length || 0)}</div>
        </div>
        <div style="font-size:15px;line-height:1.8;text-align:right;">
          <div><strong>Manager:</strong> ${escapeHtml(sheet.managerSignedBy || '-')}</div>
          <div><strong>Signed:</strong> ${escapeHtml(signedAt)}</div>
        </div>
      </div>
    </div>
  `;
}

function handleFallbackPreview(event) {
  const select = document.getElementById('agencyWorkerSelect');
  const value = select?.value || '';
  if (!value.startsWith(FALLBACK_PREFIX)) return;
  const sheet = state.savedSheets.get(value.slice(FALLBACK_PREFIX.length));
  if (!sheet) return;
  event.preventDefault();
  event.stopImmediatePropagation();
  renderSavedSheet(sheet);
}

function wireFallback() {
  document.getElementById('agencyPreviewBtn')?.addEventListener('click', handleFallbackPreview, true);
  document.getElementById('agencyWorkerSelect')?.addEventListener('change', handleFallbackPreview, true);
  document.getElementById('agencyWorkerSelect')?.addEventListener('focus', scheduleRefresh);
  document.getElementById('agencyWorkerSelect')?.addEventListener('mousedown', scheduleRefresh);
  document.getElementById('weekPicker')?.addEventListener('change', scheduleRefresh);
  document.getElementById('agencyTabBtn')?.addEventListener('click', scheduleRefresh);
}

onAuthStateChanged(auth, async (user) => {
  state.savedSheets.clear();
  if (!user) return;
  try {
    const profileSnap = await getDoc(doc(db, 'users', user.uid));
    state.profile = profileSnap.exists() ? { uid: user.uid, ...profileSnap.data() } : { uid: user.uid };
    scheduleRefresh();
  } catch (error) {
    console.warn('Agency export fallback profile load failed:', error.message);
  }
});

wireFallback();
