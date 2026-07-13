# Payroll-Critical Phone Reliability Fix Plan

## Confirmed failure mode

The public punch flow currently treats any typed name with no selected employee as a new worker. On slow phones, the employee list may still be loading or the autocomplete selection may not register before the punch button is tapped. This can create another employee record for an existing worker.

The current reuse lookup also supplies blank agency/site values while existing records are branch-scoped. That can prevent a legitimate existing employee from being reused and explains repeated records such as multiple Don Morrison entries.

## Required behavior before production trust

1. Never auto-create a worker from a punch button.
2. Require the active employee roster to finish loading before enabling punch buttons.
3. Require an existing, uniquely resolved employee ID before saving a punch.
4. Resolve workers using normalized name plus company, branch, and agency scope.
5. If several legacy records match, select one canonical active record for new punches while preserving every historical record.
6. Use an idempotency key for every punch attempt so double taps, retries, refreshes, and weak connections cannot create duplicates.
7. Wait for Firestore confirmation, then read the saved punch back before showing success.
8. Show a large confirmation screen with employee name, action, time, and confirmation ID.
9. On failure, show “PUNCH NOT RECORDED — SEE MANAGER” and retain a local recovery attempt.
10. Do not delete, rename, merge, migrate, or overwrite historical employee, worker, punch, timesheet, approval, signature, or audit records.

## Duplicate employee display

Existing duplicate worker documents must remain untouched. All worker pickers, employee rosters, weekly signoff, editing, and agency export should display a canonical grouped worker while collecting punches tied to every known legacy employee ID for that person.

## Required regression tests

- Slow network during roster load
- Double tap and repeated tap
- Browser refresh immediately after tap
- Offline-to-online transition
- Two phones punching simultaneously
- Exact duplicate names across different agencies
- Legacy blank agency/site record plus current branch-scoped record
- Clock In, Start Lunch, End Lunch, and Clock Out
- Weekly signoff and agency export include the saved punch
