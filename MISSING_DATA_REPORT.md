# QRTimeclock Missing Data Report

Date: 2026-07-13

Evidence source: compared the read-only and post-backfill Firestore snapshots from 2026-06-24 in `C:\Users\ileva\Downloads\qrtimeclock-firestore-backups`.

## Snapshot Preservation Check

All document IDs present in the read-only snapshot were still present in the post-backfill snapshot:

| Collection | Before | After | Missing After |
| --- | ---: | ---: | ---: |
| employees | 456 | 456 | 0 |
| punches | 1122 | 1123 | 0 |
| timesheets | 56 | 56 | 0 |
| missedPunchRequests | 5 | 5 | 0 |
| punch_edits | 441 | 441 | 0 |
| auditLogs | 59 | 59 | 0 |
| mergeLogs | 0 | 0 | 0 |
| users | 9 | 9 | 0 |

## Post-Backfill Data Signals

| Check | Result |
| --- | ---: |
| Employees | 456 |
| Active employees | 26 |
| Punches | 1123 |
| Active punches | 1123 |
| Timesheets | 56 |
| Punch edit history records | 441 |
| Audit logs | 59 |
| Merge logs | 0 |
| Merge-related audit logs | 42 |
| Active punches with invalid timestamp/action | 0 |
| Active punches with blank employee and worker id | 943 |
| Active punches with employee id not found in employee ids | 3 |
| Duplicate-looking active punch signatures | 7 |
| Timesheets without matching active punch rows | 4 |
| Active employees without matching punch history | 5 |
| Duplicate active worker groups by name + agency + branch | 0 |

## Important Interpretation

Blank `employeeId` punches are historical records. They should not be deleted or rewritten automatically because they are payroll evidence. Current lookup code keeps legacy name fallback, but new punch creation now requires a verified employee id for public, manager manual, missed-punch approval, and Agency Export insert paths.

The three orphan-looking employee id punches use `employee_direct_oh01_emp-1001` in the backup. They should be reviewed manually against employee profile history before any production correction.

Four saved timesheets did not match active punch rows by the simple snapshot comparison:

- `2026-04-13_alex_morales`
- `2026-04-13_brandon_evanshine`
- `2026-05-18_christopher_stone`
- `2026-05-18_noah_rinehart`

This branch does not delete or overwrite those timesheets. Weekly Signoff now derives rows from loaded punch data and reconciles saved signoff state separately.

## Risk Controls Added

- New payroll regression script fails if a future snapshot comparison loses existing payroll documents.
- New save-time validation blocks malformed new punch writes.
- Branch-scoped history lookups keep Worker My Time, Manager lookup, Weekly Signoff, and Agency Export aligned on company/site.
