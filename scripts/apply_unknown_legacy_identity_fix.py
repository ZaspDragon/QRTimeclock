from pathlib import Path

app_path = Path('app.js')
text = app_path.read_text(encoding='utf-8')
old = '''function getWorkerIdentityKey(row, directory = buildCanonicalWorkerDirectory(), profileIndex = buildWorkerProfileIndex()) {
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
}
'''
new = '''function getWorkerIdentityKey(row, directory = buildCanonicalWorkerDirectory(), profileIndex = buildWorkerProfileIndex()) {
  const recordIds = getDirectoryWorkerIds(row, profileIndex);

  // 1. Canonical profile ID or an explicitly linked legacy ID.
  for (const recordId of recordIds) {
    const profile = profileIndex.get(recordId);
    if (profile) {
      const canonicalId = String(profile.id || profile.canonicalEmployeeId || profile.employeeId || '').trim();
      if (canonicalId) return `worker:${canonicalId}`;
    }
  }

  // 2. Employee number, but only when it identifies exactly one profile.
  const employeeNumber = normalizeWorkerNumber(row?.employeeNumber || row?.empNumber || '');
  if (employeeNumber) {
    const numberMatches = (state.allEmployees || []).filter((employee) =>
      normalizeWorkerNumber(employee.employeeNumber) === employeeNumber
    );
    if (numberMatches.length === 1 && numberMatches[0]?.id) {
      return `worker:${numberMatches[0].id}`;
    }
  }

  // 3. Controlled fallback: normalized name + agency + branch must identify one profile.
  const rowName = normalizeName(getCopiedWorkerName(row));
  const rowAgency = getRecordAgencyIdentity(row);
  const rowBranch = getRecordBranchIdentity(row);
  if (rowName && rowAgency && rowBranch) {
    const controlledMatches = (state.allEmployees || []).filter((employee) =>
      normalizeName(getWorkerProfileName(employee)) === rowName
      && getRecordAgencyIdentity(employee) === rowAgency
      && getRecordBranchIdentity(employee) === rowBranch
    );
    if (controlledMatches.length === 1 && controlledMatches[0]?.id) {
      return `worker:${controlledMatches[0].id}`;
    }
  }

  // Email remains a safer fallback than an unrecognized historical worker ID.
  const email = getRecordEmail(row);
  const emailPrimaryId = email ? directory.emailPrimary.get(email) : '';
  if (emailPrimaryId) return `worker:${emailPrimaryId}`;
  if (email) return `email:${email}`;

  // Keep an unknown legacy ID separate when no controlled match exists.
  // This prevents workers with the same name from being merged accidentally.
  const stableId = recordIds[0] || '';
  if (stableId) return `worker:${stableId}`;

  const signature = getWorkerSignature(row);
  return signature ? `person:${signature}` : '';
}
'''
if old not in text:
    raise SystemExit('Expected identity function was not found; refusing to modify app.js')
text = text.replace(old, new, 1)
app_path.write_text(text, encoding='utf-8')

index_path = Path('index.html')
index = index_path.read_text(encoding='utf-8')
index = index.replace('app.js?v=20260717-2', 'app.js?v=20260727-identity2')
index_path.write_text(index, encoding='utf-8')
