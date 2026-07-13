# QRTimeclock Test Results

Date: 2026-07-13

## Commands Run

```powershell
npm.cmd run check
```

Result: passed.

```powershell
npm.cmd run test:payroll
```

Result: passed.

Output:

```text
payroll snapshot regression passed
WARN: snapshot contains 7 duplicate-looking active punch signature(s); current app validation prevents new malformed saves but does not rewrite history
```

```powershell
node -e "JSON.parse(require('fs').readFileSync('firestore.indexes.json','utf8')); console.log('firestore.indexes.json OK')"
```

Result: passed.

```powershell
git diff --check
```

Result: passed with Windows line-ending warnings only.

## What The Regression Test Proves

- Existing read-only snapshot employee, punch, timesheet, missed punch request, punch edit, audit log, merge log, and user document IDs still exist in the post-backfill snapshot.
- Active punches in the post-backfill snapshot have valid actions, timestamps, date keys, week keys, company ids, and site ids.
- Historical duplicate-looking punch signatures are reported as warnings, not rewritten.

## Not Run

- No Firebase deploy.
- No live Firestore reads or writes.
- No browser E2E test, because this repository is a static Firebase app and no local Firebase emulator configuration is present.
