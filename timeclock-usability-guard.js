import { getApps } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { getAuth, onAuthStateChanged } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import {
  getFirestore,
  collection,
  getDocs,
  limit,
  query,
  where
} from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const COMPANY_ID = 'chadwell';
const MAX_FALLBACK_ROWS = 1000;
const REFRESH_DEBOUNCE_MS = 800;
const MANAGER_CONTROL_RETRY_MS = [500, 1200, 2500];
let refreshTimer = null;
let controlsInstalled = false;
let agencySimpleModeInstalled = false;

function byId(id) {
  return document.getElementById(id);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}

function normalizeSiteId(value) {
  return String(value || '').trim().toUpperCase() || 'OH01';
}

function currentManagerSiteId() {
  const visibleSelect = document.querySelector('.manager-branch-select, [data-manager-branch-select], #managerBranchSelect');
  if (visibleSelect?.value) return normalizeSiteId(visibleSelect.value);
  const workerBranch = byId('workerBranchSelect');
  if (workerBranch?.value) return normalizeSiteId(workerBranch.value);
  const sessionKey = Object.keys(sessionStorage).find((key) => key.startsWith('managerActiveBranch:'));
  if (sessionKey) return normalizeSiteId(sessionStorage.getItem(sessionKey));
  return 'OH01';
}

function punchTimestampMs(row) {
  if (Number.isFinite(Number(row.timestampMs))) return Number(row.timestampMs);
  if (row.timestamp?.toMillis) return row.timestamp.toMillis();
  if (row.createdAt?.toMillis) return row.createdAt.toMillis();
  return 0;
}

function activePunch(row) {
  return row?.status !== 'deleted' && row?.active !== false;
}

function actionLabel(action) {
  const labels = {
    clock_in: 'Clock In',
    start_lunch: 'Start Lunch',
    end_lunch: 'End Lunch',
    clock_out: 'Clock Out'
  };
  return labels[action] || String(action || 'Punch').replaceAll('_', ' ');
}

function formatPunchTime(ms) {
  if (!ms) return 'Time unavailable';
  return new Date(ms).toLocaleString([], {
    month: 'short',
    day: 'numeric',
    hour: 'numeric',
    minute: '2-digit'
  });
}

function showStatus(message, isError = false) {
  let box = byId('punchVisibilityGuardStatus');
  if (!box) {
    box = document.createElement('div');
    box.id = 'punchVisibilityGuardStatus';
    box.className = 'status-box';
    box.style.marginBottom = '12px';
    const body = byId('livePunchBody');
    body?.closest('.mini-table-wrap, .table-wrap, section, .card')?.insertBefore(box, body?.closest('table') || body);
  }
  if (!box) return;
  box.textContent = message;
  box.style.borderColor = isError ? '#b42318' : '';
}

async function loadBranchPunches() {
  const apps = getApps();
  if (!apps.length) return [];
  const db = getFirestore(apps[0]);
  const siteId = currentManagerSiteId();
  const snap = await getDocs(query(
    collection(db, 'punches'),
    where('siteId', '==', siteId),
    limit(MAX_FALLBACK_ROWS)
  ));
  return snap.docs
    .map((record) => ({ id: record.id, ...record.data() }))
    .filter((row) => row.companyId === COMPANY_ID || !row.companyId)
    .filter(activePunch);
}

async function loadFallbackPunches() {
  const todayStart = new Date();
  todayStart.setHours(0, 0, 0, 0);
  const rows = await loadBranchPunches();
  return rows
    .filter((row) => punchTimestampMs(row) >= todayStart.getTime())
    .sort((a, b) => punchTimestampMs(b) - punchTimestampMs(a));
}

function renderFallbackPunches(rows) {
  const body = byId('livePunchBody');
  if (!body) return;
  if (!rows.length) {
    const existingText = body.textContent.trim().toLowerCase();
    if (!existingText || existingText.includes('loading') || existingText.includes('could not')) {
      body.innerHTML = '<tr><td colspan="5">No punches recorded for this branch today.</td></tr>';
    }
    return;
  }
  body.innerHTML = rows.map((row) => {
    const name = row.name || row.workerName || row.employeeName || 'Unknown worker';
    const agency = row.agencyName || row.agency || 'Direct';
    const site = row.siteId || row.branchId || currentManagerSiteId();
    return `<tr data-fallback-punch-id="${escapeHtml(row.id)}">
      <td>${escapeHtml(formatPunchTime(punchTimestampMs(row)))}</td>
      <td>${escapeHtml(name)}</td>
      <td>${escapeHtml(actionLabel(row.action))}</td>
      <td>${escapeHtml(agency)}</td>
      <td>${escapeHtml(site)}</td>
    </tr>`;
  }).join('');
}

async function ensurePunchesVisible({ announce = false } = {}) {
  const body = byId('livePunchBody');
  if (!body) return;
  try {
    const rows = await loadFallbackPunches();
    const hasPrimaryRows = body.querySelector('tr:not([data-fallback-punch-id])')
      && !/no punches|loading|could not/i.test(body.textContent);
    if (!hasPrimaryRows || announce) renderFallbackPunches(rows);
    showStatus(rows.length
      ? `${rows.length} punch${rows.length === 1 ? '' : 'es'} found for ${currentManagerSiteId()} today.`
      : `No punches found for ${currentManagerSiteId()} today.`);
  } catch (error) {
    console.warn('Punch visibility fallback failed:', error);
    showStatus(`Punch refresh failed: ${error.message || 'Unknown error'}`, true);
  }
}

function schedulePunchRefresh(options = {}) {
  window.clearTimeout(refreshTimer);
  refreshTimer = window.setTimeout(() => ensurePunchesVisible(options), REFRESH_DEBOUNCE_MS);
}

function textFromVisibleTimeArea() {
  const candidates = [
    byId('agencyPreview'),
    byId('timesheetBody')?.closest('table'),
    byId('managerTimeRangeResults')
  ].filter(Boolean);
  const visible = candidates.find((element) => !element.classList.contains('hidden') && element.offsetParent !== null)
    || candidates[0];
  return visible?.innerText?.trim() || '';
}

function selectedWeekLabel() {
  return byId('weekPicker')?.value || byId('agencyWeekPicker')?.value || 'selected week';
}

async function copyTimeSummary() {
  const text = textFromVisibleTimeArea();
  if (!text) {
    showStatus('Open Weekly Signoff or Agency Export first, then copy the time summary.', true);
    return;
  }
  await navigator.clipboard.writeText(text);
  showStatus('Time summary copied.');
}

function emailTimeSummary() {
  const text = textFromVisibleTimeArea();
  if (!text) {
    showStatus('Open Weekly Signoff or Agency Export first, then email the time summary.', true);
    return;
  }
  const subject = encodeURIComponent(`QRTimeclock hours — ${selectedWeekLabel()}`);
  const body = encodeURIComponent(`${text}\n\nSent from QRTimeclock.`);
  window.location.href = `mailto:?subject=${subject}&body=${body}`;
}

function clickTabButton(id) {
  const button = byId(id);
  if (button && !button.classList.contains('hidden')) button.click();
}

function managerNavigationIsAvailable() {
  return [byId('editPunchesTabBtn'), byId('timesheetsTabBtn'), byId('agencyTabBtn')]
    .filter(Boolean)
    .some((control) => !control.classList.contains('hidden') && !control.hidden && control.getAttribute('aria-hidden') !== 'true');
}

function lastWeekRange() {
  const now = new Date();
  const day = now.getDay();
  const thisMonday = new Date(now);
  thisMonday.setHours(0, 0, 0, 0);
  thisMonday.setDate(now.getDate() - ((day + 6) % 7));
  const start = new Date(thisMonday);
  start.setDate(start.getDate() - 7);
  const end = new Date(thisMonday);
  return { start, end };
}

function normalizeWorkerKey(row) {
  return String(row.employeeId || row.workerId || row.nameKey || row.name || row.workerName || 'unknown')
    .trim().toLowerCase();
}

async function runLastWeekPunchAudit() {
  const output = byId('lastWeekPunchAuditStatus');
  if (!output) return;
  output.textContent = 'Checking recorded punches from last week...';
  try {
    const { start, end } = lastWeekRange();
    const rows = (await loadBranchPunches())
      .filter((row) => {
        const ms = punchTimestampMs(row);
        return ms >= start.getTime() && ms < end.getTime();
      })
      .sort((a, b) => punchTimestampMs(a) - punchTimestampMs(b));

    const byWorkerDay = new Map();
    rows.forEach((row) => {
      const date = new Date(punchTimestampMs(row)).toLocaleDateString('en-CA');
      const key = `${normalizeWorkerKey(row)}|${date}`;
      if (!byWorkerDay.has(key)) byWorkerDay.set(key, []);
      byWorkerDay.get(key).push(row);
    });

    const warnings = [];
    byWorkerDay.forEach((dayRows) => {
      const actions = dayRows.map((row) => row.action);
      const name = dayRows[0]?.name || dayRows[0]?.workerName || 'Unknown worker';
      const date = new Date(punchTimestampMs(dayRows[0])).toLocaleDateString();
      if (!actions.includes('clock_in')) warnings.push(`${name} — ${date}: missing Clock In`);
      if (!actions.includes('clock_out')) warnings.push(`${name} — ${date}: missing Clock Out`);
      const hasOneLunchSide = actions.includes('start_lunch') !== actions.includes('end_lunch');
      if (hasOneLunchSide) warnings.push(`${name} — ${date}: incomplete lunch pair`);
    });

    const label = `${start.toLocaleDateString()}–${new Date(end.getTime() - 1).toLocaleDateString()}`;
    output.innerHTML = warnings.length
      ? `<strong>${rows.length} recorded punches checked for ${escapeHtml(label)}.</strong><br>${warnings.length} possible missing-punch problem${warnings.length === 1 ? '' : 's'}:<br>${warnings.slice(0, 40).map(escapeHtml).join('<br>')}`
      : `<strong>${rows.length} recorded punches checked for ${escapeHtml(label)}.</strong><br>No incomplete Clock In/Clock Out or one-sided lunch pairs were found in the records returned.`;
  } catch (error) {
    output.textContent = `Last-week audit failed: ${error.message || 'Unknown error'}`;
  }
}

function installSimpleAgencyMode() {
  if (agencySimpleModeInstalled) return;
  const tab = byId('agencyTab');
  const legacyPreview = tab?.querySelector('.agency-print-preview');
  if (!tab || !legacyPreview) return;
  agencySimpleModeInstalled = true;

  tab.querySelectorAll('.agency-sticky-tools, .agency-stats-grid, #agencyCoverageStatus, .agency-recovery-actions, .agency-layout')
    .forEach((element) => { element.style.display = 'none'; });
  tab.querySelector('.agency-workbench-card > .card-head')?.classList.remove('split-head');
  const heading = tab.querySelector('.agency-workbench-card > .card-head h2');
  const description = tab.querySelector('.agency-workbench-card > .card-head p');
  if (heading) heading.textContent = 'Temp Agency Export';
  if (description) description.textContent = 'Choose one worker, preview the weekly sheet, then print or save it as a PDF.';
  const topExportButtons = tab.querySelector('.agency-workbench-card > .card-head .form-actions');
  if (topExportButtons) topExportButtons.style.display = 'none';
  legacyPreview.open = true;
  const summary = legacyPreview.querySelector('summary');
  if (summary) summary.style.display = 'none';

  const auditBox = document.createElement('div');
  auditBox.className = 'card';
  auditBox.style.marginTop = '16px';
  auditBox.innerHTML = `
    <div class="card-head">
      <h3>Last Week Punch Check</h3>
      <p>Read-only check for recorded Clock In, Clock Out, and incomplete lunch pairs. It does not edit or delete payroll data.</p>
    </div>
    <button id="runLastWeekPunchAuditBtn" class="secondary-btn" type="button">Check Last Week's Recorded Punches</button>
    <div id="lastWeekPunchAuditStatus" class="status-box" style="margin-top:12px;">Not checked yet.</div>`;
  legacyPreview.insertAdjacentElement('afterend', auditBox);
  byId('runLastWeekPunchAuditBtn')?.addEventListener('click', runLastWeekPunchAudit);
}

function installManagerControls() {
  if (controlsInstalled || !byId('appShell') || !managerNavigationIsAvailable()) return false;
  controlsInstalled = true;
  const bar = document.createElement('div');
  bar.id = 'timeclockQuickActions';
  bar.className = 'card';
  bar.style.marginBottom = '14px';
  bar.innerHTML = `
    <div class="card-head">
      <h3>Quick Actions</h3>
      <p>Refresh punches, correct time, review the week, or send hours.</p>
    </div>
    <div class="form-actions" style="flex-wrap:wrap">
      <button id="quickRefreshPunches" class="primary-btn" type="button">Refresh Today's Punches</button>
      <button id="quickEditPunches" class="secondary-btn" type="button">Edit / Add Punch</button>
      <button id="quickWeeklyTime" class="secondary-btn" type="button">Weekly Signoff</button>
      <button id="quickAgencyExport" class="secondary-btn" type="button">Agency Export</button>
      <button id="quickCopyTime" class="ghost-btn" type="button">Copy Time Summary</button>
      <button id="quickEmailTime" class="ghost-btn" type="button">Email Time Summary</button>
    </div>`;
  const appShell = byId('appShell');
  const tabBar = byId('tabBar');
  appShell.insertBefore(bar, tabBar || appShell.firstChild);
  byId('quickRefreshPunches')?.addEventListener('click', () => ensurePunchesVisible({ announce: true }));
  byId('quickEditPunches')?.addEventListener('click', () => clickTabButton('editPunchesTabBtn'));
  byId('quickWeeklyTime')?.addEventListener('click', () => clickTabButton('timesheetsTabBtn'));
  byId('quickAgencyExport')?.addEventListener('click', () => clickTabButton('agencyTabBtn'));
  byId('quickCopyTime')?.addEventListener('click', () => copyTimeSummary().catch((error) => showStatus(error.message, true)));
  byId('quickEmailTime')?.addEventListener('click', emailTimeSummary);
  document.addEventListener('change', (event) => {
    if (event.target.matches('.manager-branch-select, [data-manager-branch-select], #managerBranchSelect')) {
      schedulePunchRefresh({ announce: true });
    }
  });
  schedulePunchRefresh();
  return true;
}

function scheduleManagerControlInstall() {
  MANAGER_CONTROL_RETRY_MS.forEach((delay) => {
    window.setTimeout(() => {
      installSimpleAgencyMode();
      if (!controlsInstalled) installManagerControls();
    }, delay);
  });
}

function watchWorkerPunchConfirmation() {
  const status = byId('workerStatusMessage');
  if (!status) return;
  const observer = new MutationObserver(() => {
    const text = status.textContent.toLowerCase();
    if (/saved|recorded|clocked|punch complete|success/.test(text)) {
      const viewButton = byId('workerViewTimeBtn');
      if (viewButton && !byId('workerMyTimePanel')?.classList.contains('hidden')) {
        window.setTimeout(() => viewButton.click(), 350);
      }
      schedulePunchRefresh();
    }
  });
  observer.observe(status, { childList: true, subtree: true, characterData: true });
}

function initialize() {
  watchWorkerPunchConfirmation();
  installSimpleAgencyMode();
  const apps = getApps();
  if (!apps.length) return;
  const auth = getAuth(apps[0]);
  onAuthStateChanged(auth, (user) => {
    if (!user) return;
    scheduleManagerControlInstall();
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initialize, { once: true });
} else {
  initialize();
}
