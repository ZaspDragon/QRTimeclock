import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { getAuth, onAuthStateChanged } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import { collection, doc, getDoc, getDocs, getFirestore, query, where } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
const COMPANY_ID = 'chadwell';
const AGENCY_NAMES = {
  sterling_staffing: 'Sterling Staffing',
  excel_staffing: 'Excel Staffing',
  lifestyle_staffing: 'Lifestyle Staffing',
};
let profile = null;
let busy = false;

function normalizeName(value) {
  return String(value || '').trim().toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function escapeHtml(value) {
  return String(value ?? '').replaceAll('&', '&amp;').replaceAll('<', '&lt;').replaceAll('>', '&gt;').replaceAll('"', '&quot;').replaceAll("'", '&#39;');
}
function formatDateKey(date) {
  const pad = (value) => String(value).padStart(2, '0');
  return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`;
}
function mondayFor(value) {
  const date = new Date(value);
  const day = date.getDay();
  date.setDate(date.getDate() + (day === 0 ? -6 : 1 - day));
  date.setHours(0, 0, 0, 0);
  return date;
}
function selectedWeekRange() {
  const picker = document.getElementById('weekPicker');
  const start = mondayFor(picker?.value ? `${picker.value}T00:00:00` : new Date());
  const end = new Date(start);
  end.setDate(end.getDate() + 7);
  return { start, end, startMs: start.getTime(), endMs: end.getTime(), weekKey: formatDateKey(start) };
}
function agencyIdFromLabel(label) {
  const value = String(label || '').trim();
  return Object.entries(AGENCY_NAMES).find(([, name]) => name === value)?.[0] || value;
}
function agencyLabel(id) {
  return AGENCY_NAMES[id] || id || 'Direct';
}
function parseSelectedWorker() {
  const select = document.getElementById('agencyLegacyWorkerSelect');
  const option = select?.selectedOptions?.[0];
  if (!select?.value || !option) return null;
  const text = String(option.textContent || '').trim();
  const details = (text.match(/\(([^)]*)\)\s*$/)?.[1] || '').split('·').map((part) => part.trim());
  const name = text.replace(/\s*\([^)]*\)\s*$/, '').trim();
  const rawIdentity = String(select.value || '');
  return {
    name,
    nameKey: normalizeName(name),
    identityId: rawIdentity.startsWith('worker:') ? rawIdentity.slice(7) : rawIdentity,
    agencyId: agencyIdFromLabel(details[0] || ''),
    siteId: details[1] || profile?.siteId || profile?.branch || 'OH01',
  };
}
function actionKey(value) {
  const key = String(value || '').trim().toLowerCase().replace(/[\s-]+/g, '_');
  const map = {
    clockin: 'clock_in', clock_in: 'clock_in',
    startlunch: 'start_lunch', start_lunch: 'start_lunch', lunch_out: 'start_lunch', lunchout: 'start_lunch',
    endlunch: 'end_lunch', end_lunch: 'end_lunch', lunch_in: 'end_lunch', lunchin: 'end_lunch',
    clockout: 'clock_out', clock_out: 'clock_out',
  };
  return map[key] || key;
}
function timestampMs(row) {
  const direct = Number(row?.timestampMs || 0);
  if (direct) return direct;
  if (row?.timestamp?.toMillis instanceof Function) return row.timestamp.toMillis();
  if (row?.createdAt?.toMillis instanceof Function) return row.createdAt.toMillis();
  const parsed = Date.parse(String(row?.timestamp || row?.createdAt || ''));
  return Number.isFinite(parsed) ? parsed : 0;
}
function rowSite(row) {
  return String(row?.siteId || row?.branch || row?.assignedSiteId || '').trim();
}
function isActive(row) {
  return row && row.active !== false && String(row.status || '').toLowerCase() !== 'deleted';
}
function scopeMatch(row, selected) {
  const company = String(row.companyId || row.companyID || COMPANY_ID).trim();
  const site = rowSite(row);
  const agency = String(row.agencyId || '').trim();
  return company === COMPANY_ID
    && (!site || site === selected.siteId)
    && (!agency || agency === selected.agencyId);
}

async function resolveIds(selected) {
  const ids = new Set([selected.identityId].filter(Boolean));
  const add = (row, id = '') => {
    if (id) ids.add(id);
    [row?.employeeId, row?.employeeID, row?.workerId, row?.mergedInto].filter(Boolean).forEach((value) => ids.add(String(value)));
  };
  if (selected.identityId && !selected.identityId.includes('|')) {
    try {
      const snap = await getDoc(doc(db, 'employees', selected.identityId));
      if (snap.exists()) add(snap.data(), snap.id);
    } catch (_) {}
  }
  const nameSnap = await getDocs(query(collection(db, 'employees'), where('nameKey', '==', selected.nameKey.replaceAll(' ', '_'))));
  nameSnap.docs.forEach((record) => {
    const row = record.data();
    if (scopeMatch(row, selected)) add(row, record.id);
  });
  let changed = true;
  while (changed) {
    changed = false;
    for (const id of [...ids]) {
      const merged = await getDocs(query(collection(db, 'employees'), where('mergedInto', '==', id)));
      merged.docs.forEach((record) => {
        if (!scopeMatch(record.data(), selected)) return;
        const before = ids.size;
        add(record.data(), record.id);
        if (ids.size !== before) changed = true;
      });
    }
  }
  return ids;
}

async function loadPunches(selected, ids, range) {
  const snap = await getDocs(query(
    collection(db, 'punches'),
    where('timestampMs', '>=', range.startMs),
    where('timestampMs', '<', range.endMs)
  ));
  return snap.docs.map((record) => ({ id: record.id, ...record.data() })).filter((row) => {
    if (!isActive(row) || !scopeMatch(row, selected)) return false;
    const idMatch = ids.has(String(row.employeeId || '')) || ids.has(String(row.employeeID || '')) || ids.has(String(row.workerId || ''));
    const nameMatch = normalizeName(row.name || row.workerName || row.employeeName || row.nameKey) === selected.nameKey;
    return idMatch || nameMatch;
  }).map((row) => ({ ...row, action: actionKey(row.action), timestampMs: timestampMs(row) }))
    .sort((a, b) => a.timestampMs - b.timestampMs);
}

async function loadSavedSheet(selected, ids, range) {
  const snap = await getDocs(query(collection(db, 'timesheets'), where('weekKey', '==', range.weekKey)));
  return snap.docs.map((record) => ({ id: record.id, ...record.data() })).find((row) => {
    if (!scopeMatch(row, selected)) return false;
    return ids.has(String(row.employeeId || '')) || ids.has(String(row.workerId || '')) || normalizeName(row.name || row.workerName || row.nameKey) === selected.nameKey;
  }) || null;
}

function formatTime(ms) {
  return ms ? new Date(ms).toLocaleTimeString([], { hour: 'numeric', minute: '2-digit' }) : '-';
}
function summarize(punches) {
  const groups = new Map();
  punches.forEach((punch) => {
    const dateKey = punch.dateKey || formatDateKey(new Date(punch.timestampMs));
    if (!groups.has(dateKey)) groups.set(dateKey, []);
    groups.get(dateKey).push(punch);
  });
  const daily = {};
  let total = 0;
  [...groups.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([dateKey, rows]) => {
    rows.sort((a, b) => a.timestampMs - b.timestampMs);
    const first = (action) => rows.find((row) => row.action === action);
    const all = (action) => rows.filter((row) => row.action === action);
    const clockIn = first('clock_in');
    const clockOuts = all('clock_out');
    const clockOut = clockOuts[clockOuts.length - 1];
    const lunchOut = first('start_lunch');
    const lunchIn = first('end_lunch');
    let hours = 0;
    if (clockIn && clockOut && clockOut.timestampMs > clockIn.timestampMs) {
      let worked = clockOut.timestampMs - clockIn.timestampMs;
      if (lunchOut && lunchIn && lunchIn.timestampMs > lunchOut.timestampMs) worked -= lunchIn.timestampMs - lunchOut.timestampMs;
      hours = Math.max(0, worked / 3600000);
      total += hours;
    }
    daily[dateKey] = { clockIn, lunchOut, lunchIn, clockOut, hours };
  });
  return { daily, total, daysWorked: groups.size };
}
function render(selected, range, summary, saved) {
  const preview = document.getElementById('agencyPreview');
  if (!preview) return;
  const rows = Object.entries(summary.daily).map(([dateKey, day]) => `<tr>
    <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(dateKey)}</td>
    <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(formatTime(day.clockIn?.timestampMs))}</td>
    <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(formatTime(day.lunchOut?.timestampMs))}</td>
    <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(formatTime(day.lunchIn?.timestampMs))}</td>
    <td style="border:1px solid #bbb;padding:10px;">${escapeHtml(formatTime(day.clockOut?.timestampMs))}</td>
    <td style="border:1px solid #bbb;padding:10px;">${day.hours.toFixed(2)}</td>
  </tr>`).join('') || '<tr><td colspan="6" style="border:1px solid #bbb;padding:10px;">No punches found for this worker and week.</td></tr>';
  const signedAt = saved?.managerSignedAt?.seconds ? new Date(saved.managerSignedAt.seconds * 1000).toLocaleString() : '-';
  preview.innerHTML = `<div id="agencyPrintableSheet" style="background:#fff;color:#111;border-radius:12px;padding:24px;min-height:200px;">
    <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;margin-bottom:18px;">
      <div><h2 style="margin:0 0 8px;font-size:28px;">Weekly Time Sheet</h2>
        <div><strong>Worker:</strong> ${escapeHtml(selected.name)}</div><div><strong>Agency:</strong> ${escapeHtml(agencyLabel(selected.agencyId))}</div>
        <div><strong>Branch:</strong> ${escapeHtml(selected.siteId)}</div><div><strong>Week Start:</strong> ${escapeHtml(range.weekKey)}</div>
        <div><strong>Status:</strong> ${escapeHtml(saved?.status || 'open')}</div></div>
      <div style="text-align:right;"><div><strong>Company:</strong> Chadwell</div><div><strong>Generated:</strong> ${escapeHtml(new Date().toLocaleString())}</div></div>
    </div>
    <table style="width:100%;border-collapse:collapse;margin-bottom:18px;"><thead><tr>
      <th style="border:1px solid #bbb;padding:10px;background:#f3f6fa;">Date</th><th style="border:1px solid #bbb;padding:10px;background:#f3f6fa;">Clock In</th>
      <th style="border:1px solid #bbb;padding:10px;background:#f3f6fa;">Start Lunch</th><th style="border:1px solid #bbb;padding:10px;background:#f3f6fa;">End Lunch</th>
      <th style="border:1px solid #bbb;padding:10px;background:#f3f6fa;">Clock Out</th><th style="border:1px solid #bbb;padding:10px;background:#f3f6fa;">Hours</th>
    </tr></thead><tbody>${rows}</tbody></table>
    <div style="display:flex;justify-content:space-between;gap:16px;flex-wrap:wrap;"><div><div><strong>Total Hours:</strong> ${summary.total.toFixed(2)}</div><div><strong>Days Worked:</strong> ${summary.daysWorked}</div></div>
    <div style="text-align:right;"><div><strong>Manager:</strong> ${escapeHtml(saved?.managerSignedBy || '-')}</div><div><strong>Signed:</strong> ${escapeHtml(signedAt)}</div></div></div>
  </div>`;
}

async function preview(event) {
  const button = event.target instanceof Element ? event.target.closest('#agencyPreviewBtn') : null;
  if (!button || busy) return;
  const selected = parseSelectedWorker();
  if (!selected) return;
  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
  busy = true;
  const previewEl = document.getElementById('agencyPreview');
  if (previewEl) previewEl.innerHTML = '<div class="empty-state">Loading current edited punches...</div>';
  try {
    const range = selectedWeekRange();
    const ids = await resolveIds(selected);
    const [punches, saved] = await Promise.all([loadPunches(selected, ids, range), loadSavedSheet(selected, ids, range)]);
    render(selected, range, summarize(punches), saved);
  } catch (error) {
    console.error('[agency-export-live-preview-v2]', error);
    if (previewEl) previewEl.innerHTML = `<div class="empty-state">${escapeHtml(error.message || 'Could not load current punches.')}</div>`;
  } finally {
    busy = false;
  }
}

document.addEventListener('click', preview, true);
onAuthStateChanged(auth, async (user) => {
  profile = null;
  if (!user) return;
  try {
    const snap = await getDoc(doc(db, 'users', user.uid));
    profile = snap.exists() ? { uid: user.uid, ...snap.data() } : { uid: user.uid };
  } catch (_) {
    profile = { uid: user.uid };
  }
});
