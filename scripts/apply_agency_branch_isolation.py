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
  return !recordSiteId || recordSiteId === getCurrentSiteId();
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

if old_source not in text:
    raise SystemExit("Expected getAgencySourcePunches block was not found; no file changed.")
if old_rows not in text:
    raise SystemExit("Expected agency employee-row block was not found; no file changed.")

patched = text.replace(old_source, new_source, 1).replace(old_rows, new_rows, 1)
if patched == text:
    raise SystemExit("Patch produced no change.")

APP_PATH.write_text(patched, encoding="utf-8")
print("Applied agency export branch isolation without changing stored data.")
