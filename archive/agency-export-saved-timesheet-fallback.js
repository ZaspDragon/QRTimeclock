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
const AGENCY_NAMES = {
  sterling_staffing: 'Sterling Staffing',
  excel_staffing: 'Excel Staffing',
  lifestyle_staffing: 'Lifestyle Staffing',
};

const state = {
  profile: null,
  busy: false,
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

function profileSiteIds(profile = state.profile) {
  const values = profile?.branches || profile?.siteIds || profile?.branch || profile?.siteId || profile?.assignedSiteId || [];
  const raw = Array.isArray(values) ? values : [values];
  return [...new Set(raw.map((value) => String(value || '').trim()).filter(Boolean))];
}

function activeSiteId(profile = state.profile) {
  const allowed = profileSiteIds(profile);
  const stored = sessionStorage.getItem(`managerActiveBranch:${profile?.uid || ''}`);
  return allowed.includes(stored) ? stored : allowed[0] || CURRENT_SITE_ID;
}

function isAgencyUser(profile = state.profile) {
  return normalizeRole(profile?.role) === 'agency_admin' || !!profile?.agencyId;
}

function parseSelectedOption() {
  const select = document.getElementById('agencyLegacyWorkerSelect');
  const option = select?.selectedOptions?.[0];
  if (!select || !option || !select.value) return null;

  const rawText = String(option.textContent || '').trim();
  const detailsMatch = rawText.match(/\(([^)]*)\)\s*$/);
  const details = (detailsMatch?.[1] || '').split('·').map((part) => part.trim());
  const name = prettifyHumanName(rawText.replace(/\s*\([^)]*\)\s*$/, ''));
  const agencyText = details[0] || '';
  const agencyId = Object.entries(AGENCY_NAMES).find(([, label]) => label === agencyText)?.[0] || agencyText;
  const siteId = details[1] || activeSiteId();
  const rawIdentity = String(select.value || '');
  const identityId = rawIdentity.startsWith('worker:') ? rawIdentity.slice(7) : rawIdentity;

  return {
    select,
    option,
    rawIdentity,
    identityId,
    name,
    nameKey: normalizeName(name),
    agencyId,
    siteId,
  };
}

function isActivePunch(punch) {
  if (!punch || punch.active === false) return false;
  return String(punch.status || '').trim().toLowerCase() !== 'deleted';
}

function punchTimestampMs(punch) {
  const explicit = Number(punch?.timestampMs || 0);
  if (explicit) return explicit;
  if (punch?.timestamp?.toMillis instanceof Function) return punch.timestamp.toMillis();
  if (punch?.createdAt?.toMillis instanceof Function) return punch.createdAt.toMillis();
  const parsed = Date.parse(String(punch?.timestamp || punch?.createdAt || ''));
  return Number.isFinite(parsed) ? parsed : 0;
}

function punchSiteId(punch) {
  return String(punch?.siteId || punch?.branch || punch?.assignedSiteId || '').trim();
}

function matchesScope(row, selected) {
  const companyId = String(row?.companyId || row?.companyID || CURRENT_COMPANY_ID).trim();
  const siteId = punchSiteId(row);
  const agencyId = String(row?.agencyId || '').trim();
  return companyId === CURRENT_COMPANY_ID
    && (!selected.siteId || !siteId || siteId === selected.siteId)
    && (!selected.agencyId || !agencyId || agencyId === selected.agencyId);
}

async function resolveWorkerIdentity(selected) {
  const ids = new Set();
  const employeeRows = new Map();

  const addEmployee = (snapshot) => {
    if (!snapshot?.exists?.()) return;
    const row = { id: snapshot.id, ...snapshot.data() };
    employeeRows.set(row.id, row);
    ids.add(row.id);
    if (row.employeeId) ids.add(String(row.employeeId));
    if (row.workerId) ids.add(String(row.workerId));
    if (row.mergedInto) ids.add(String(row.mergedInto));
  };

  if (selected.identityId && !selected.identityId.includes('|')) {
    try {
      addEmployee(await getDoc(doc(db, 'employees', selected.identityId)));
    } catch (error) {
      console.warn('Legacy preview direct employee lookup skipped:', error.message);
    }
  }

  if (selected.nameKey) {
    const nameSnap = await getDocs(query(collection(db, 'employees'), where('nameKey', '==', selected.nameKey)));
    nameSnap.docs.forEach((record) => {
      const row = { id: record.id, ...record.data() };
      if (!matchesScope(row, selected)) return;
      employeeRows.set(row.id, row);
      ids.add(row.id);
      if (row.employeeId) ids.add(String(row.employeeId));
      if (row.workerId) ids.add(String(row.workerId));
      if (row.mergedInto) ids.add(String(row.mergedInto));
    });
  }

  let changed = true;
  while (changed) {
    changed = false;
    for (const id of [...ids]) {
      const mergedSnap = await getDocs(query(collection(db, 'employees'), where('mergedInto', '==', id)));
      mergedSnap.docs.forEach((record) => {
        const row = { id: record.id, ...record.data() };
        if (!matchesScope(row, selected)) return;
        employeeRows.set(row.id, row);
        if (!ids.has(row.id)) {
          ids.add(row.id);
          changed = true;
        }
        if (row.employeeId && !ids.has(String(row.employeeId))) {
          ids.add(String(row.employeeId));
          changed = true;
        }
        if (row.workerId && !ids.has(String(row.workerId))) {
          ids.add(String(row.workerId));
          changed = true;
        }
      });
    }
  }

  return { ids, employees: [...employeeRows.values()] };
}

async function loadSelectedWeekPunches(selected, identity) {
  const weekKey = selectedWeekKey();
  const snap = await getDocs(query(collection(db, 'punches'), where('weekKey', '==', weekKey)));
  const rows = snap.docs.map((record) => ({ id: record.id, ...record.data() }));

  return rows.filter((punch) => {
    if (!isActivePunch(punch) || !matchesScope(punch, selected)) return false;
    const employeeId = String(punch.employeeId || '').trim();
    const workerId = String(punch.workerId || '').trim();
    const idMatch = identity.ids.has(employeeId) || identity.ids.has(workerId);
    const nameMatch = normalizeName(punch.name || punch.workerName || punch.employeeName || punch.nameKey) === selected.nameKey;
    return idMatch || nameMatch;
  }).sort((left, right) => punchTimestampMs(left) - punchTimestampMs(right));
}

function formatTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' });
}

function buildDailyTotals(punches) {
  const grouped = new Map();
  punches.forEach((punch) => {
    const ms = punchTimestampMs(punch);
    const dateKey = String(punch.dateKey || (ms ? formatDateKey(new Date(ms)) : '')).trim();
    if (!dateKey || !ms) return;
    if (!grouped.has(dateKey)) grouped.set(dateKey, []);
    grouped.get(dateKey).push({ ...punch, timestampMs: ms });
  });

  const dailyTotals = {};
  let weeklyHours = 0;

  [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([dateKey, dayPunches]) => {
    dayPunches.sort((a, b) => a.timestampMs - b.timestampMs);
    const byAction = (action) => dayPunches.filter((punch) => punch.action === action);
    const clockIn = byAction('clock_in')[0];
    const clockOutRows = byAction('clock_out');
    const clockOut = clockOutRows[clockOutRows.length - 1];
    const lunchOut = byAction('start_lunch')[0];
    const lunchIn = byAction('end_lunch')[0];

    let hours = 0;
    if (clockIn && clockOut && clockOut.timestampMs > clockIn.timestampMs) {
      let workedMs = clockOut.timestampMs - clockIn.timestampMs;
      if (lunchOut && lunchIn && lunchIn.timestampMs > lunchOut.timestampMs) {
        workedMs -= lunchIn.timestampMs - lunchOut.timestampMs;
      }
      hours = Math.max(0, workedMs / 3600000);
      weeklyHours += hours;
    }

    dailyTotals[dateKey] = {
      clock_in: formatTime(clockIn?.timestampMs),
      start_lunch: formatTime(lunchOut?.timestampMs),
      end_lunch: formatTime(lunchIn?.timestampMs),
      clock_out: formatTime(clockOut?.timestampMs),
      hours,
    };
  });

  return { dailyTotals, weeklyHours, daysWorked: grouped.size };
}

async function loadSavedSheet(selected, identity) {
  const weekKey = selectedWeekKey();
  const snap = await getDocs(query(collection(db, 'timesheets'), where('weekKey', '==', weekKey)));
  const rows = snap.docs.map((record) => ({ id: record.id, ...record.data() }));
  return rows.find((sheet) => {
    if (!matchesScope(sheet, selected)) return false;
    const idMatch = identity.ids.has(String(sheet.employeeId || '')) || identity.ids.has(String(sheet.workerId || ''));
    const nameMatch = normalizeName(sheet.name || sheet.workerName || sheet.employeeName || sheet.nameKey) === selected.nameKey;
    return idMatch || nameMatch;
  }) || null;
}

function buildDailyRows(dailyTotals) {
  const keys = Object.keys(dailyTotals || {}).sort();
  if (!keys.length) {
    return '<tr><td colspan="6" style="border:1px solid #bbb;padding:10px;">No punches recorded for this week.</td></tr>';
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

function formatDateTime(ms) {
  if (!ms) return '-';
  return new Date(ms).toLocaleString([], {
    year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit'
  });
}

function showPreviewStatus(message, isError = false) {
  const preview = document.getElementById('agencyPreview');
  if (!preview) return;
  preview.innerHTML = `<div class="empty-state${isError ? ' error' : ''}">${escapeHtml(message)}</div>`;
}

function renderLiveSheet(selected, summary, savedSheet) {
  const preview = document.getElementById('agencyPreview');
  if (!preview) return;
  const signedAtMs = Number(savedSheet?.managerSignedAt?.seconds || 0) * 1000;
  const status = savedSheet?.status || 'open';
  const managerSignedBy = savedSheet?.managerSignedBy || '-';

  preview.innerHTML = `
    <div id="agencyPrintableSheet" style="background:#fff;color:#111;border-radius:12px;padding:24px;min-height:200px;">
      <div style="display:flex;justify-content:space-between;gap:16px;align-items:flex-start;flex-wrap:wrap;margin-bottom:18px;">
        <div>
          <h2 style="margin:0 0 8px;font-size:28px;">Weekly Time Sheet</h2>
          <div style="font-size:15px;line-height:1.6;">
            <div><strong>Worker:</strong> ${escapeHtml(selected.name)}</div>
            <div><strong>Agency:</strong> ${escapeHtml(agencyLabel(selected.agencyId))}</div>
            <div><strong>Branch:</strong> ${escapeHtml(selected.siteId || activeSiteId())}</div>
            <div><strong>Week Start:</strong> ${escapeHtml(selectedWeekKey())}</div>
            <div><strong>Status:</strong> ${escapeHtml(status)}</div>
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
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Start Lunch</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">End Lunch</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Clock Out</th>
            <th style="border:1px solid #bbb;padding:10px;text-align:left;background:#f3f6fa;">Hours</th>
          </tr>
        </thead>
        <tbody>${buildDailyRows(summary.dailyTotals)}</tbody>
      </table>
      <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-top:24px;">
        <div style="font-size:15px;line-height:1.8;">
          <div><strong>Total Hours:</strong> ${Number(summary.weeklyHours || 0).toFixed(2)}</div>
          <div><strong>Days Worked:</strong> ${Number(summary.daysWorked || 0)}</div>
        </div>
        <div style="font-size:15px;line-height:1.8;text-align:right;">
          <div><strong>Manager:</strong> ${escapeHtml(managerSignedBy)}</div>
          <div><strong>Signed:</strong> ${escapeHtml(signedAtMs ? formatDateTime(signedAtMs) : '-')}</div>
        </div>
      </div>
    </div>
  `;
}

async function handleLegacyPreview(event) {
  const selected = parseSelectedOption();
  if (!selected || state.busy) return;

  event.preventDefault();
  event.stopImmediatePropagation();
  state.busy = true;
  showPreviewStatus('Loading current punches...');

  try {
    const identity = await resolveWorkerIdentity(selected);
    const punches = await loadSelectedWeekPunches(selected, identity);
    const summary = buildDailyTotals(punches);
    const savedSheet = await loadSavedSheet(selected, identity);
    renderLiveSheet(selected, summary, savedSheet);
  } catch (error) {
    console.error('Legacy agency preview rebuild failed:', error);
    showPreviewStatus(error.message || 'Could not rebuild the weekly sheet from punches.', true);
  } finally {
    state.busy = false;
  }
}

function wireLegacyPreview() {
  document.getElementById('agencyPreviewBtn')?.addEventListener('click', handleLegacyPreview, true);
}

onAuthStateChanged(auth, async (user) => {
  state.profile = null;
  if (!user) return;
  try {
    const profileSnap = await getDoc(doc(db, 'users', user.uid));
    state.profile = profileSnap.exists() ? { uid: user.uid, ...profileSnap.data() } : { uid: user.uid };
  } catch (error) {
    console.warn('Legacy preview profile load failed:', error.message);
  }
});

wireLegacyPreview();
