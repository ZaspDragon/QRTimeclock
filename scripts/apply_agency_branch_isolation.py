from pathlib import Path

APP_PATH = Path("app.js")
text = APP_PATH.read_text(encoding="utf-8")

old_source = """function getAgencySourcePunches() {
  return Array.isArray(state.agencyReview.rangePunchRows)
    ? state.agencyReview.rangePunchRows
    : state.selectedWeekPunchRows;
}
"""

new_source = """function getAgencyRecordSiteId(record) {
  return String(record?.siteId || record?.assignedSiteId || record?.branch || '').trim();
}

function agencyExportAllowsMultipleSites() {
  return normalizedRole() === 'agency_admin';
}

function agencyExportRecordInCurrentSite(record) {
  if (agencyExportAllowsMultipleSites()) return true;
  const recordSiteId = getAgencyRecordSiteId(record);
  return recordSiteId === getCurrentSiteId();
}

function getAgencySourcePunches() {
  const rows = Array.isArray(state.agencyReview.rangePunchRows)
    ? state.agencyReview.rangePunchRows
    : state.selectedWeekPunchRows;
  return (Array.isArray(rows) ? rows : []).filter(agencyExportRecordInCurrentSite);
}
"""

old_rows = """  rows.forEach((row) => {
    const key = agencyRosterRecordKey(row);
    if (key && !unique.has(key)) unique.set(key, row);
  });
  return [...unique.values()];
}
"""

new_rows = """  rows.forEach((row) => {
    if (!agencyExportRecordInCurrentSite(row)) return;
    const key = agencyRosterRecordKey(row);
    if (key && !unique.has(key)) unique.set(key, row);
  });
  return [...unique.values()];
}
"""

permissive_scope = "return !recordSiteId || recordSiteId === getCurrentSiteId();"
strict_scope = "return recordSiteId === getCurrentSiteId();"

patched = text
if old_source in patched:
    patched = patched.replace(old_source, new_source, 1)
elif permissive_scope in patched:
    patched = patched.replace(permissive_scope, strict_scope, 1)
elif strict_scope not in patched:
    raise SystemExit("Expected agency export source block was not found; no file changed.")

if old_rows in patched:
    patched = patched.replace(old_rows, new_rows, 1)
elif "if (!agencyExportRecordInCurrentSite(row)) return;" not in patched:
    raise SystemExit("Expected agency employee-row block was not found; no file changed.")

if patched == text:
    print("Strict agency branch isolation is already applied.")
else:
    APP_PATH.write_text(patched, encoding="utf-8")
    print("Applied strict agency export branch isolation without changing stored data.")
