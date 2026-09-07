# Worker hours lookup repair

## What is fixed

- Lookup errors, invalid responses and timeouts display unavailable hours (`—`), not a false zero.
- Lookup requests time out after 15 seconds and can be retried.
- Changing the worker, branch, agency or dates discards an in-flight result.
- Lookup updates only the hours panel, leaving clock status and punch buttons alone.
- The legacy past-time click callback no longer receives a click event as its employee.
- Last 2 Weeks selects this week plus the previous week (14 days).
- Weekly overtime is calculated separately for each Monday-start week in the selected range.
- The unsupported Overall Hours shortcut is replaced with This Month. Custom ranges allow up to 31 days.
- The server pages identity queries so recent punches after the first 400 historical records are included. It reports excessive histories/results instead of returning a truncated total.
- Missing legacy date keys use Eastern time for OH01 and OHC.

## Confirmed deployment dependency

On September 7, 2026, HEAD and browser preflight (OPTIONS) requests to
`https://us-central1-qrtimeclock-42764.cloudfunctions.net/publicWorkerTimeLookup`
returned HTML HTTP 404. This is not a successful worker lookup.

The lookup module **is** already loaded indirectly by `firebase-config.js`.
Updating GitHub files alone does not deploy the Firebase function. Deploy and
verify this single function before considering employee self-service restored.

## Scope and access

Clock-in/out, lunch punch writers, writer priorities, employee creation, existing
PIN behavior, Firestore rules, manager edits and payroll exports are unchanged.
The function only reads Firestore. No data migration or rule deployment is needed.

This retains the existing public exact-name + branch + agency lookup contract.
Those fields identify a worker; they do not authenticate the requester. This
patch does not introduce login/PIN verification for hours viewing. Do not describe
the name-based view as a private authenticated employee portal.

## Deployment from an authorized Firebase environment

Use this repair's checkout with Node 20 or 22, Firebase CLI, and an account
authorized to deploy to `qrtimeclock-42764`. Do not paste account credentials or
service-account keys into chat. If Firebase requires billing or account access,
resolve that in the Firebase console before deployment.

Run the regression checks from the repository root:

```sh
npm run check
```

Install the existing function dependencies, then deploy **only** the scoped
lookup function:

```sh
npm install --prefix functions
firebase deploy --project qrtimeclock-42764 --only functions:publicWorkerTimeLookup
```

Do not use a blanket `firebase deploy` or deploy `publicWorkerTimeLookupByName` as
a shortcut. The alternate name-only endpoint does not apply the same agency filter.
No change to Firestore rules is part of this repair.

[Firebase's documentation for deploying specific functions](https://firebase.google.com/docs/functions/manage-functions#deploy_functions)

Check preflight without requesting any employee data:

```sh
curl -i -X OPTIONS \
  -H 'Origin: https://zaspdragon.github.io' \
  -H 'Access-Control-Request-Method: POST' \
  https://us-central1-qrtimeclock-42764.cloudfunctions.net/publicWorkerTimeLookup
```

Expect HTTP 204 and `Access-Control-Allow-Origin: https://zaspdragon.github.io`.
This verifies routing/CORS only, not real worker data access.

Publish the reviewed frontend changes using the existing GitHub Pages deployment.
If employees instead use Firebase Hosting, publish that frontend separately with
`firebase deploy --project qrtimeclock-42764 --only hosting` from the reviewed checkout.

## Acceptance before calling this live

1. On a phone, select an existing worker, correct branch and agency, then My Hours.
2. Compare daily punches and totals with the manager's saved records.
3. Check Last Week and a custom range. Check a worker with linked historical IDs.
4. Confirm a legitimate manager correction is reflected after refreshing hours.
5. Verify another branch/agency's explicitly scoped records are excluded.
6. During a normal scheduled punch, verify Clock In still records once. Do not
   create test punches in production merely to exercise buttons.

Automated coverage runs the actual lookup module and HTTP function with mocked
DOM/network/Firestore adapters, and the unchanged punch dispatcher. It tests all
four punch actions while lookup is waiting, lookup timeout/failure/retry, stale
responses, past dates, weekly totals, pagination, linked IDs, explicit scope
exclusion, corrections and ambiguous identities. It makes no production writes.

Live Firebase deployment and real worker-record reconciliation remain separate
acceptance steps. Existing name normalization, missing-scope legacy data and
overnight-shift handling may require follow-up after real-data verification.

## Rollback

Revert this repair's frontend commit using the normal hosting release process.
The clock-in writers and database rules are unchanged throughout. The deployed
lookup function is read-only; rolling back the frontend does not require any
employee, punch, guard, state or timesheet deletion.
