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
const MAX_FALLBACK_ROWS = 250;
const REFRESH_DEBOUNCE_MS = 800;
let refreshTimer = null;
let controlsInstalled = false;

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
  const site = String(value || '').trim().toUpperCase();
  return site === 'OHC' ? 'OHC' : 'OH01';
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

function startOfTodayMs() {
  const now = new Date();
  now.setHours(0, 0, 0, 0);
  return now.getTime();
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

async function loadFallbackPunches() {
  const apps = getApps();
  if (!apps.length) return [];
  const db = getFirestore(apps[0]);
  const siteId = currentManagerSiteId();

  const snap = await getDocs(query(
    collection(db, 'punches'),
    where('siteId', '==', siteId),
    limit(MAX_FALLBACK_ROWS)
  ));

  const todayStart = startOfTodayMs();
  return snap.docs
    .map((record) => ({ id: record.id, ...record.data() }))
    .filter((row) => row.companyId === COMPANY_ID || !row.companyId)
    .filter(activePunch)
    .filter((row) => punchTimestampMs(row) >= todayStart)
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
    byId('agencyPreviewPanel'),
    byId('agencyWorkbench'),
    byId('timesheetBody')?.closest('table'),
    byId('managerTimeRangeResults')
  ].filter(Boolean);

  const visible = candidates.find((element) => !element.classList.contains('hidden') && element.offsetParent !== null)
    || candidates[0];
  return visible?.innerText?.trim() || '';
}

function selectedWeekLabel() {
  return byId('weekPicker')?.value
    || byId('agencyWeekPicker')?.value
    || 'selected week';
}

async function copyTimeSummary() {
  const text = textFromVisibleTimeArea();
  if (!text) {
    showStatus('Open Weekly Signoff or Agency Export first, then copy the time summary.', true);
    return;
  }
  await navigator.clipboard.writeText(text);
  showStatus('Time summary copied. You can paste it into email, Slack, or payroll notes.');
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

function installManagerControls() {
  if (controlsInstalled || !byId('appShell')) return;
  controlsInstalled = true;

  const bar = document.createElement('div');
  bar.id = 'timeclockQuickActions';
  bar.className = 'card';
  bar.style.marginBottom = '14px';
  bar.innerHTML = `
    <div class="card-head">
      <h3>Quick Actions</h3>
      <p>Refresh punches, correct time, review the week, or send hours without hunting through every tab.</p>
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
  const apps = getApps();
  if (!apps.length) return;
  const auth = getAuth(apps[0]);
  onAuthStateChanged(auth, (user) => {
    if (!user) return;
    window.setTimeout(installManagerControls, 500);
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initialize, { once: true });
} else {
  initialize();
}
