# Why the app wasn't running smoothly — and what changed

Date: 2026-08-15

## Root cause #1 (the big one): five punch writers racing each other

`.worker-action-btn` (Clock In / Start Lunch / End Lunch / Clock Out) was being
handled by **five** different modules, each of which installed its own
capture-phase `document` click listener and called `stopImmediatePropagation()`:

| Module | Added |
|---|---|
| `public-clock-document-id-fix.js` | 2026-08-11 |
| `public-clock-permission-hotfix.js` | 2026-08-11 |
| `canonical-public-clock.js` | 2026-08-10 |
| `stable-public-clock-handler.js` | 2026-08-02 |
| `new-worker-first-punch-hotfix.js` | 2026-08-01 |

Capture listeners on the same node fire in **registration order**, and the first
one to call `stopImmediatePropagation()` silently kills the other four plus
`app.js`'s own handler. Registration order was decided by:

- static vs. lazy `import()` in `firebase-config.js`,
- whether a module installed immediately or waited for `DOMContentLoaded`,
- and how fast each of ~21 separate module downloads finished.

On a fast desktop the newest writer usually won. On a warehouse phone on weak
Wi-Fi, an **older** writer often won — which is exactly why the same tap
produced different results on different phones and different refreshes:
permission-denied errors, duplicate worker records, "punch didn't save",
inconsistent lunch behavior. Each past fix added another writer instead of
removing the old one, so the race got worse with every fix.

### Fix

New module `punch-writer-lock.js` owns the **only** punch click listener. Each
writer now registers itself with an explicit priority:

```
public-clock-document-id-fix     100   (newest, rules-safe, document-ID-based)
public-clock-permission-hotfix    90
canonical-public-clock            80
stable-public-clock-handler       70
new-worker-first-punch-hotfix     60   (only when its own canHandle() applies)
```

The highest-priority writer that can handle the tap runs — **every time,
regardless of load order or network speed**. If no hotfix module loads at all,
the lock stays out of the way and `app.js`'s built-in handler runs, so a failed
download can no longer leave the punch buttons dead.

The lock also adds a hard one-at-a-time guard, so repeated taps on a laggy
screen can't start two punch writes.

No punch-saving logic, Firestore collection, rule, or historical record was
changed. Only *which* writer runs, and *that it is the same one every time*.

## Root cause #2: DOM observers thrashing the manager views

- `agency-export-dropdown-dedupe.js` kept a `MutationObserver` on
  `document.documentElement` **for the life of the page** and re-ran on every
  single DOM mutation. Live punches and timesheet tables mutate constantly, so
  this ran thousands of times per session. It now debounces to one pass per
  animation frame and disconnects as soon as the Agency Export select exists.
- The same file's per-select observer re-triggered itself (dedupe removes
  options → observer fires → dedupe...). Now debounced.
- `lunch-labels.js` ran a full `querySelectorAll` sweep for **each inserted
  node** on a body-wide observer. Rendering a few hundred timesheet rows meant a
  few hundred sweeps. Now batched into one pass per frame, skipping detached
  nodes.

## Root cause #3: dead code shipped to phones

Four modules were unreachable (imported by nothing) but still living in the
hosting root and cluttering the codebase during debugging. Moved to `archive/`
and excluded from deploy via `firebase.json`:

- `agency-export-saved-timesheet-fallback.js`
- `firestore-compatible-public-punch.js`
- `punch-exceptions-dashboard.js` (superseded by `-v2`)
- `secure-temp-time-lookup.js`

## Verified

- All modules pass `node --check` (ES module syntax).
- No writer body references the click `event` object any more.
- Firestore rules, indexes, collection names, and `app.js` payroll logic: untouched.

## Still recommended (not done here — needs your call)

1. **Every punch does a full `window.location.reload()`** ~0.9s after saving.
   That re-downloads and re-parses a 345 KB `app.js` plus ~20 modules on a
   phone. Replacing it with an in-place confirmation screen would be the single
   biggest remaining speed win on the clock-in screen.
2. **`app.js` is 8,658 lines / 345 KB, unbundled,** and is loaded on the public
   clock screen even though workers use ~5% of it. Splitting the public clock
   into its own small entry point would cut phone load time dramatically.
3. **Consolidate the 5 punch writers into 1.** The lock makes behavior
   deterministic today; deleting the 4 losers (after a week of clean punches)
   removes the confusion permanently.
4. The Firebase Hosting site `qrtimeclock-42764.web.app` currently returns
   **404 / Site Not Found**, so the deployed app could not be smoke-tested.
