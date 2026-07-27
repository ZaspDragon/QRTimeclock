from pathlib import Path

path = Path('app.js')
text = path.read_text(encoding='utf-8')

replacements = []

replacements.append((
"""function getWorkerProfileIds(worker) {
  return [
    worker?.id,
    worker?.employeeId,
    worker?.workerId,
    worker?.userId,
    worker?.uid
  ].map((value) => String(value || '').trim()).filter(Boolean);
}""",
"""function getWorkerProfileIds(worker) {
  const linkedIds = [
    ...(Array.isArray(worker?.linkedWorkerIds) ? worker.linkedWorkerIds : []),
    ...(Array.isArray(worker?.legacyWorkerIds) ? worker.legacyWorkerIds : []),
    ...(Array.isArray(worker?.identityAliases) ? worker.identityAliases : [])
  ];
  return [
    worker?.id,
    worker?.canonicalEmployeeId,
    worker?.employeeId,
    worker?.workerId,
    worker?.userId,
    worker?.uid,
    ...linkedIds
  ].map((value) => String(value || '').trim()).filter(Boolean);
}"""))

replacements.append((
"""function getRecordWorkerIds(row) {
  return [
    row?.employeeId,
    row?.workerId,
    row?.userId,
    row?.uid
  ].map((value) => String(value || '').trim()).filter(Boolean);
}""",
"""function getRecordWorkerIds(row) {
  return [
    row?.canonicalEmployeeId,
    row?.employeeId,
    row?.workerId,
    row?.userId,
    row?.uid
  ].map((value) => String(value || '').trim()).filter(Boolean);
}"""))

replacements.append((
"""function getWorkerIdentityKey(row, directory = buildCanonicalWorkerDirectory(), profileIndex = buildWorkerProfileIndex()) {
  const stableId = getDirectoryWorkerIds(row, profileIndex)[0] || '';
  const email = getRecordEmail(row);
  const signature = getWorkerSignature(row);
  const signaturePrimaryId = signature ? directory.signaturePrimary.get(signature) : '';
  const emailPrimaryId = email ? directory.emailPrimary.get(email) : '';

  if (signaturePrimaryId) return `worker:${signaturePrimaryId}`;
  if (stableId) return `worker:${stableId}`;
  if (emailPrimaryId) return `worker:${emailPrimaryId}`;
  if (email) return `email:${email}`;
  return signature ? `person:${signature}` : '';
}""",
"""function getWorkerIdentityKey(row, directory = buildCanonicalWorkerDirectory(), profileIndex = buildWorkerProfileIndex()) {
  const recordIds = getDirectoryWorkerIds(row, profileIndex);
  for (const recordId of recordIds) {
    const profile = profileIndex.get(recordId);
    if (profile) {
      const canonicalId = String(profile.id || profile.canonicalEmployeeId || profile.employeeId || '').trim();
      if (canonicalId) return `worker:${canonicalId}`;
    }
  }

  const stableId = recordIds[0] || '';
  if (stableId) return `worker:${stableId}`;

  const employeeNumber = normalizeWorkerNumber(row?.employeeNumber || row?.empNumber || '');
  if (employeeNumber) {
    const matchingProfile = (state.allEmployees || []).find((employee) =>
      normalizeWorkerNumber(employee.employeeNumber) === employeeNumber
    );
    if (matchingProfile?.id) return `worker:${matchingProfile.id}`;
  }

  const email = getRecordEmail(row);
  const emailPrimaryId = email ? directory.emailPrimary.get(email) : '';
  if (emailPrimaryId) return `worker:${emailPrimaryId}`;
  if (email) return `email:${email}`;

  const signature = getWorkerSignature(row);
  const signaturePrimaryId = signature ? directory.signaturePrimary.get(signature) : '';
  if (signaturePrimaryId) return `worker:${signaturePrimaryId}`;
  return signature ? `person:${signature}` : '';
}"""))

replacements.append((
"""  renderDuplicateWorkerWarning(employees);
""",
"""  const selectedEmployeeId = String(els.employeeDocId?.value || '').trim();
  const duplicateScope = selectedEmployeeId
    ? employees.filter((employee) => employee.id === selectedEmployeeId || getWorkerProfileIds(employee).includes(selectedEmployeeId))
    : filtered;
  renderDuplicateWorkerWarning(duplicateScope);
  renderIdentityAdminPanel();
"""))

for old, new in replacements:
    if old not in text:
        raise SystemExit(f'Missing expected block:\n{old[:180]}')
    text = text.replace(old, new, 1)

append_marker = "// CANONICAL_WORKER_IDENTITY_FIX_V1"
if append_marker not in text:
    text += r'''

// CANONICAL_WORKER_IDENTITY_FIX_V1
function getIdentityPanelHost() {
  if (!els.duplicateWorkerWarning?.parentElement) return null;
  let panel = document.getElementById('workerIdentityDiagnosticPanel');
  if (!panel) {
    panel = document.createElement('section');
    panel.id = 'workerIdentityDiagnosticPanel';
    panel.className = 'status-box hidden';
    panel.style.marginTop = '14px';
    els.duplicateWorkerWarning.insertAdjacentElement('afterend', panel);
  }
  return panel;
}

function identityPunchRows() {
  return dedupePunches([
    ...(state.allPunchRows || []),
    ...(state.selectedWeekPunchRows || []),
    ...(state.agencyReview?.rangePunchRows || [])
  ]).filter(isActivePunchRecord);
}

function getCanonicalEmployeeForIdentityPanel() {
  const selectedId = String(els.employeeDocId?.value || '').trim();
  if (selectedId) {
    const selected = (state.allEmployees || []).find((employee) => employee.id === selectedId);
    if (selected) return selected;
  }
  const filter = String(els.empFilterInput?.value || '').trim();
  if (!filter) return null;
  const normalizedFilter = normalizeName(filter);
  const matches = (state.allEmployees || []).filter((employee) => {
    const nameKey = normalizeName(getWorkerProfileName(employee));
    const numberKey = normalizeWorkerNumber(employee.employeeNumber);
    return nameKey.includes(normalizedFilter) || numberKey.includes(normalizeWorkerNumber(filter));
  });
  return matches.length === 1 ? matches[0] : null;
}

function identityRowsForEmployee(employee) {
  if (!employee) return [];
  const canonicalId = String(employee.id || '').trim();
  const ids = new Set(getWorkerProfileIds(employee));
  ids.add(canonicalId);
  return identityPunchRows().filter((row) => getRecordWorkerIds(row).some((id) => ids.has(id)));
}

function summarizeIdentityId(workerId, punches) {
  const rows = punches.filter((row) => getRecordWorkerIds(row).includes(workerId));
  const times = rows.map(punchTimestampMs).filter(Boolean).sort((a, b) => a - b);
  return {
    workerId,
    count: rows.length,
    earliest: times.length ? formatDateTime(times[0]) : '-',
    latest: times.length ? formatDateTime(times[times.length - 1]) : '-'
  };
}

function renderIdentityAdminPanel() {
  const panel = getIdentityPanelHost();
  if (!panel) return;
  const employee = getCanonicalEmployeeForIdentityPanel();
  if (!isAdmin() || !employee) {
    panel.classList.add('hidden');
    panel.innerHTML = '';
    return;
  }

  const canonicalId = String(employee.id || '').trim();
  const linkedIds = [...new Set([
    ...(Array.isArray(employee.linkedWorkerIds) ? employee.linkedWorkerIds : []),
    ...(Array.isArray(employee.legacyWorkerIds) ? employee.legacyWorkerIds : []),
    ...(Array.isArray(employee.identityAliases) ? employee.identityAliases : [])
  ].map((value) => String(value || '').trim()).filter(Boolean))];
  const punches = identityPunchRows();
  const summaries = [canonicalId, ...linkedIds].filter(Boolean).map((id) => summarizeIdentityId(id, punches));

  panel.classList.remove('hidden');
  panel.innerHTML = `
    <strong>Worker identity diagnostic</strong>
    <p class="tiny">Linking adds an alias to the canonical profile. It does not delete, move, or rewrite historical records.</p>
    <div class="stats-grid" style="margin-top:10px;">
      <div class="stat-card"><span>Profile ID</span><strong>${escapeHtml(canonicalId || '-')}</strong></div>
      <div class="stat-card"><span>Employee #</span><strong>${escapeHtml(employee.employeeNumber || '-')}</strong></div>
      <div class="stat-card"><span>Agency</span><strong>${escapeHtml(agencyLabel(employee.agencyId))}</strong></div>
      <div class="stat-card"><span>Branch</span><strong>${escapeHtml(getWorkerBranchId(employee) || '-')}</strong></div>
    </div>
    <div style="overflow:auto;margin-top:10px;">
      <table>
        <thead><tr><th>Worker ID</th><th>Punches</th><th>Earliest</th><th>Latest</th></tr></thead>
        <tbody>${summaries.length ? summaries.map((row) => `
          <tr><td>${escapeHtml(row.workerId)}</td><td>${row.count}</td><td>${escapeHtml(row.earliest)}</td><td>${escapeHtml(row.latest)}</td></tr>
        `).join('') : '<tr><td colspan="4">No identity IDs found.</td></tr>'}</tbody>
      </table>
    </div>
    <div class="grid-form" style="margin-top:12px;">
      <label class="full-width"><span>Legacy worker ID to link</span><input id="identityLegacyWorkerIdInput" type="text" placeholder="Paste the old worker ID" /></label>
      <div class="form-actions full-width">
        <button id="identityPreviewLinkBtn" class="secondary-btn" type="button">Preview Link</button>
        <button id="identityConfirmLinkBtn" class="primary-btn hidden" type="button">Confirm Safe Link</button>
      </div>
      <div id="identityLinkPreview" class="status-box full-width hidden"></div>
    </div>
  `;

  document.getElementById('identityPreviewLinkBtn')?.addEventListener('click', () => previewWorkerIdentityLink(employee));
  document.getElementById('identityConfirmLinkBtn')?.addEventListener('click', () => confirmWorkerIdentityLink(employee));
}

function previewWorkerIdentityLink(employee) {
  const legacyId = String(document.getElementById('identityLegacyWorkerIdInput')?.value || '').trim();
  const preview = document.getElementById('identityLinkPreview');
  const confirmBtn = document.getElementById('identityConfirmLinkBtn');
  if (!legacyId || !preview || !confirmBtn) {
    toast('Enter the legacy worker ID first.', true);
    return;
  }
  if (legacyId === employee.id || getWorkerProfileIds(employee).includes(legacyId)) {
    preview.textContent = 'That ID is already linked to this employee.';
    preview.classList.remove('hidden');
    confirmBtn.classList.add('hidden');
    return;
  }
  const rows = identityPunchRows().filter((row) => getRecordWorkerIds(row).includes(legacyId));
  const copiedNames = [...new Set(rows.map(getCopiedWorkerName).filter(Boolean))];
  const agencies = [...new Set(rows.map((row) => row.agencyId || row.agencyName || '').filter(Boolean))];
  const branches = [...new Set(rows.map(getWorkerBranchId).filter(Boolean))];
  preview.innerHTML = `
    <strong>Preview only — no data has changed.</strong><br>
    Link <strong>${escapeHtml(legacyId)}</strong> to <strong>${escapeHtml(getWorkerProfileName(employee))}</strong> (${escapeHtml(employee.employeeNumber || '-')}).<br>
    Matching punches currently loaded: <strong>${rows.length}</strong><br>
    Names on records: ${escapeHtml(copiedNames.join(', ') || '-')}<br>
    Agencies: ${escapeHtml(agencies.join(', ') || '-')} · Branches: ${escapeHtml(branches.join(', ') || '-')}
  `;
  preview.dataset.legacyId = legacyId;
  preview.classList.remove('hidden');
  confirmBtn.classList.remove('hidden');
}

async function confirmWorkerIdentityLink(employee) {
  if (!isAdmin()) return;
  const preview = document.getElementById('identityLinkPreview');
  const legacyId = String(preview?.dataset.legacyId || '').trim();
  if (!legacyId) {
    toast('Preview the link first.', true);
    return;
  }
  const existingOwner = (state.allEmployees || []).find((candidate) =>
    candidate.id !== employee.id && getWorkerProfileIds(candidate).includes(legacyId)
  );
  if (existingOwner) {
    toast(`That worker ID is already linked to ${getWorkerProfileName(existingOwner)}.`, true);
    return;
  }

  const linkedWorkerIds = [...new Set([
    ...(Array.isArray(employee.linkedWorkerIds) ? employee.linkedWorkerIds : []),
    legacyId
  ].map((value) => String(value || '').trim()).filter(Boolean))];

  try {
    await updateDoc(doc(db, 'employees', employee.id), {
      canonicalEmployeeId: employee.id,
      linkedWorkerIds,
      identityLinkUpdatedAt: serverTimestamp(),
      identityLinkUpdatedBy: state.profile?.name || state.me?.email || 'Admin'
    });
    await logAudit('worker_identity_linked', 'employee', employee.id, {
      linkedWorkerIds: employee.linkedWorkerIds || []
    }, {
      linkedWorkerIds,
      linkedLegacyWorkerId: legacyId
    }, 'Linked legacy worker ID without rewriting historical records');

    employee.canonicalEmployeeId = employee.id;
    employee.linkedWorkerIds = linkedWorkerIds;
    state.weeklyDataCache.clear();
    renderEmployeeList(state.allEmployees || []);
    attachTimesheetView({ force: true });
    toast('Worker profile linked safely. Weekly time is rebuilding now.');
  } catch (error) {
    console.error(error);
    toast(error.message || 'Could not link worker identity.', true);
  }
}
'''

path.write_text(text, encoding='utf-8')
print('Canonical worker identity fix applied to app.js')
