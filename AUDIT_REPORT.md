# QRTimeclock Repository Audit

Date: 2026-07-13

Scope audited: `app.js`, `index.html`, `firestore.rules`, `firestore.indexes.json`, `scripts/branch-backfill-preview.js`, and the available Firestore backup snapshots in `C:\Users\ileva\Downloads\qrtimeclock-firestore-backups`.

Production safety: no Firebase deploy was run, no live Firestore data was read or modified, and no collections were renamed or deleted.

## Payroll Data Flow

### Punch creation

- Public worker punches are created by `handleWorkerPunch` in `app.js`.
- Manager manual punches are created by `handleManualPunchSubmit`.
- Approved missed punch requests create punches in `approveRequest`.
- Agency Export missing-punch insert creates punches in `addMissingAgencyPunch`.
- All punch writes target the existing `punches` collection.

### Punch reads

- Public Worker My Time uses `loadPunchesForEmployeeRange` and `fetchPunchesWithRange`.
- Manager Edit Punches and live dashboard use `attachManagerLiveViews`.
- Weekly Signoff uses `attachTimesheetView`, `loadCompatibleWeeklyPunchRows`, `getDerivedTimesheetRows`, and `buildWeekTotals`.
- Employee self-service My Timecard uses `attachMyTimecardView`.
- Agency Export uses `handleAgencyDateRangeChange`, `getAgencySourcePunches`, `buildAgencyReviewRows`, and `buildAgencyPunchTotals`.
- Duplicate/merged worker lookup uses `buildCanonicalWorkerDirectory`, `getWorkerIdentityKey`, and related employee profile indexes.

### Punch edits and soft deletion

- Manager punch edits are handled by `editPunch`.
- Manager delete is a soft delete in `deletePunchRecord`; it updates `status: "deleted"` and `active: false`.
- Agency edit/delete/restore paths write through `writeAgencyPunchChange` and `restoreAgencySoftDeletedPunches`.
- Edit history is appended to `punch_edits`.
- `firestore.rules` blocks physical deletes for `punches`, `punch_edits`, `timesheets`, `employees`, `workers`, `missedPunchRequests`, `auditLogs`, and `mergeLogs`.

### Timesheets and signoff

- Weekly Signoff is derived from current punch rows rather than relying only on saved timesheet docs.
- Saved signoff state is reconciled by `findSavedTimesheetForGroup`.
- Sign and reopen operations update the existing `timesheets` collection through `signTimesheet` and `reopenTimesheet`.

### Agency Export

- Agency Export combines active/inactive employee rows, worker profile rows, public employee cache rows, and selected punch rows.
- It groups by canonical identity and warns if loaded punches are not represented.
- Export output is generated from `agencyExportRows`.

## Findings

### Fixed in this branch

- Date-range punch history lookup could query by employee/name without company and branch constraints. This could create missing history, permission failures, or cross-branch results depending on rules and indexes.
- Merged employee lookup could run without branch scope. This could fail for managers or miss relevant branch-scoped records.
- Public punch save could fall back to a name-only worker if employee auto-create failed, creating new punches without a verified employee id.
- Manager manual punch creation wrote `employeeId: ""`, creating new orphan-prone payroll records.
- Approved missed-punch requests updated request status before validating that the resulting punch could be created.
- Agency Export missing-punch insert did not share a central punch validation guard.
- Firestore indexes were missing the new branch-scoped arbitrary date-range history query shapes.

### Existing safety already present

- Physical deletes are blocked by Firestore rules.
- Punch delete UI performs soft delete.
- Employee merge flow marks duplicate employees merged and writes `mergeLogs` for new merges.
- Agency Export performs loaded-punch coverage logging.
- Duplicate-looking active employees are detected by normalized name, agency, and branch.

### Residual historical issues

- The June 24 post-backfill snapshot contains legacy active punches with blank `employeeId` or duplicate-looking signatures. This branch prevents new malformed saves but does not rewrite historical data.
- The post-backfill snapshot has zero `mergeLogs` but 42 merge-related `auditLogs`, so older merges predate append-only merge log coverage.
- Some historical timesheets do not match active punch rows in the backup pair. These are reported in `MISSING_DATA_REPORT.md`.

## Collections Preserved

The audit and repair keep the existing collection names:

- `employees`
- `workers`
- `punches`
- `timesheets`
- `missedPunchRequests`
- `punch_edits`
- `auditLogs`
- `mergeLogs`
- `users`

No destructive migration or collection rename was introduced.
