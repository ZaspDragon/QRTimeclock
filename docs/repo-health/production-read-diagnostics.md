# Production Firestore read diagnostics

## Finding

`app.js` currently enables Firestore read-counter console logging when the hostname ends in `.github.io`:

```js
const DEV_READ_COUNTERS_ENABLED = location.hostname === 'localhost'
  || location.hostname === '127.0.0.1'
  || location.hostname.endsWith('.github.io');
```

Because the live application is hosted on GitHub Pages, this treats production as a development environment. The wrappers still perform the same Firestore reads, but every counted read also updates in-memory counters and emits `console.info` messages during worker, manager, payroll, and agency flows.

## Safe corrective change

In a focused follow-up code PR, change the condition to localhost-only:

```js
const DEV_READ_COUNTERS_ENABLED = location.hostname === 'localhost'
  || location.hostname === '127.0.0.1';
```

This is backward-compatible and does not alter:

- Firestore queries or writes
- punches or timesheets
- employee or user records
- branch selection
- authentication or permissions
- saved production data

## Validation checklist

- Run `node --check app.js`.
- Load the app on GitHub Pages and confirm no `[QRTimeclock Firestore reads]` messages appear.
- Load from localhost and confirm the diagnostics still work.
- Verify worker clock-in, lunch, clock-out, manager timesheet lookup, and agency review behavior are unchanged.

## Why this PR is documentation-only

The active `app.js` is a large permission-sensitive module. This audit records the exact one-line fix separately so it can be applied and reviewed without mixing it with punch, payroll, or role-navigation changes. No production behavior or stored data is changed by this document.