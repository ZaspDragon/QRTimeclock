# QRTimeclock Repair Summary

Date: 2026-07-13

## Repairs Made

- Added `scopedPunchHistoryConstraints` so date-range punch history queries include `companyId`, `siteId`, and agency scope when applicable.
- Updated `loadPunchesForEmployeeRange` so public Worker My Time, Manager worker lookup, Weekly Signoff compatibility lookup, and Agency Export date range lookup use branch-scoped punch reads.
- Scoped merged employee lookup by `siteId` and `assignedSiteId` so merged profile history can still roll up without broad unscoped reads.
- Added `validatePunchPayloadForSave` to block new punches with missing/invalid name, action, timestamp, date, week, company, site, or employee identity.
- Updated public worker punch save to require a verified employee profile id before saving.
- Updated manager manual punch save to resolve one active employee profile before creating the punch.
- Updated missed-punch approval to validate the punch before marking the request approved.
- Updated Agency Export missing-punch insert to use the same punch validation.
- Added Firestore composite indexes for branch-scoped date-range history and merged employee lookup queries.
- Added `scripts/payroll-snapshot-regression.mjs`.
- Added `package.json` scripts for `check`, `test:payroll`, and `build`.

## Data Safety

- No collection was renamed.
- No delete path was added.
- No live Firebase data was touched.
- No migration was run.
- No historical punch, employee, worker, timesheet, approval, signature, edit, audit, or merge record was overwritten.

## Remaining Manual Follow-Up

- Deploy or paste the updated `firestore.indexes.json` indexes through Firebase before relying on the new arbitrary date-range branch-scoped queries in production.
- Review legacy blank-id and duplicate-looking punch signatures from the backup before any manual production cleanup.
- If old merge audit records need canonical `mergeLogs`, do that as a separate reviewed append-only backfill, not as an automatic rewrite.
