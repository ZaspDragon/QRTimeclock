from pathlib import Path

app_path = Path('app.js')
text = app_path.read_text(encoding='utf-8')
old = """  // Keep an unknown legacy ID separate when no controlled match exists.\n  // This prevents workers with the same name from being merged accidentally.\n  const stableId = recordIds[0] || '';\n  if (stableId) return `worker:${stableId}`;\n\n  const signature = getWorkerSignature(row);\n  return signature ? `person:${signature}` : '';\n"""
new = """  // Legacy records may be missing agency/branch fields even though the active roster\n  // supplies that scope. Resolve only when one active profile has this exact name and\n  // its branch matches the selected branch. This does not use a same-name merge when\n  // multiple profiles exist.\n  if (rowName) {\n    const selectedBranch = normalizeIdentityToken(getCurrentSiteId());\n    const uniqueRosterMatches = (state.allEmployees || []).filter((employee) =>\n      isActiveEmployee(employee)\n      && normalizeName(getWorkerProfileName(employee)) === rowName\n      && (!selectedBranch || getRecordBranchIdentity(employee) === selectedBranch)\n    );\n    if (uniqueRosterMatches.length === 1 && uniqueRosterMatches[0]?.id) {\n      return `worker:${uniqueRosterMatches[0].id}`;\n    }\n  }\n\n  // Keep an unknown legacy ID separate when no controlled match exists.\n  // This prevents workers with the same name from being merged accidentally.\n  const stableId = recordIds[0] || '';\n  if (stableId) return `worker:${stableId}`;\n\n  const signature = getWorkerSignature(row);\n  return signature ? `person:${signature}` : '';\n"""
if old not in text:
    raise SystemExit('Expected fallback block not found; refusing to modify app.js')
text = text.replace(old, new, 1)
app_path.write_text(text, encoding='utf-8')

index_path = Path('index.html')
index = index_path.read_text(encoding='utf-8')
index = index.replace('app.js?v=20260727-identity2', 'app.js?v=20260727-identity3')
index_path.write_text(index, encoding='utf-8')
