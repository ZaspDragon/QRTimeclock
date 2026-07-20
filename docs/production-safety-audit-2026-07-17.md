# QRTimeclock production safety audit

Date reviewed: 2026-07-17

## Scope

This is a documentation-only safety review. No punch creation, editing, payroll calculations, employee matching, permissions, Firestore schemas, or production data behavior were changed.

## Confirmed findings

### 1. Firestore read diagnostics run on the production GitHub Pages site

`DEV_READ_COUNTERS_ENABLED` currently evaluates to true for every hostname ending in `.github.io`.

That includes the production deployment. The wrappers do not add Firestore reads by themselves, but they keep production-only counters and console logging active for normal workers and managers. This adds noise during troubleshooting and makes it harder to distinguish real production errors from diagnostic output.

Safe follow-up:

- Restrict read diagnostics to `localhost` and `127.0.0.1`.
- Optionally allow an explicit query-string or local-storage flag for a temporary administrator diagnostic session.
- Do not change the underlying `getDoc` or `getDocs` calls in the same pull request.

### 2. Site selection is hard-coded in application source

The active site defaults to `OH01`, and the source comment instructs a developer to edit the JavaScript when deploying an OHC-specific manager application.

Risk:

- A deployment can accidentally point managers at the wrong branch.
- A source edit for deployment can be committed unintentionally.
- Troubleshooting branch visibility becomes harder because build configuration and application logic are mixed together.

Safe follow-up:

- Keep `OH01` as the backward-compatible default.
- Read an optional site value from a small deployment configuration object or validated URL parameter.
- Accept only known values from `BRANCH_OPTIONS`.
- Never use an unvalidated parameter directly in Firestore paths or permission decisions.

### 3. Role navigation has become difficult to reason about

The startup code binds worker, employee, manager, timesheet, edit-punch, reports, admin, approval, employee-roster, and agency controls in one large module.

This is not automatically a bug, but it raises regression risk. A navigation change can accidentally expose a tab, skip a role guard, or initialize a workflow before the signed-in profile is ready.

Safe follow-up:

- Document one role-to-tab matrix before changing navigation.
- Add a small shared function that returns visible tabs for a role without changing existing permissions.
- Test worker, employee, manager, admin, owner, and agency accounts separately.
- Keep permission enforcement in Firestore rules and write guards; hiding a tab is not authorization.

## Recommended order of work

1. Disable production read-counter logging only.
2. Add tests or a manual validation checklist for each role's visible navigation.
3. Introduce validated deployment configuration for site selection.
4. Refactor navigation in small pull requests only after the role matrix is verified.

## Validation notes

Reviewed the current `main` version of `app.js` and confirmed:

- Firebase reads are wrapped by counted helpers.
- `.github.io` currently enables the diagnostic counter.
- `CURRENT_SITE_ID` is hard-coded to `OH01`.
- Multiple role-specific UI elements are initialized in the same application module.

## Data safety

This audit does not:

- delete or edit punches
- delete or edit employees
- alter timesheets or approvals
- change user accounts or roles
- change Firestore collections, indexes, or rules
- run migrations
- modify production behavior
