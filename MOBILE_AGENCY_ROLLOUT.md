# Mobile and agency rollout validation

This change removes the OH01/OHC iframe wrappers, locks branch context inside the real app, and requires every public temp to choose a staffing agency.

## Data safety

- No punches, timesheets, approvals, signatures, edits, requests, or audit records are deleted or rewritten.
- Existing blank-agency employee profiles are updated only after that worker chooses an agency.
- Existing profiles already assigned to another agency are rejected rather than silently reassigned.
- New worker IDs include branch, agency, and normalized name to reduce collisions.

## Validation checklist

- OH01 and OHC QR links open the native app page instead of an iframe.
- Mobile inputs remain tappable and editable.
- Branch is locked from the URL context.
- Agency selection is required for every public punch.
- Sterling, Excel, and Lifestyle are available.
- Agency choice is saved locally for the worker's next visit.
- Existing blank-agency profiles receive only the selected agency metadata.
- Existing punches remain untouched.
