# Guarded, non-destructive repository rules patch; this note retriggers validation.
from pathlib import Path

RULES_PATH = Path('firestore.rules')
text = RULES_PATH.read_text(encoding='utf-8')


def replace_once(old: str, new: str, label: str) -> None:
    global text
    count = text.count(old)
    if count != 1:
        raise SystemExit(f'{label}: expected exactly one match, found {count}; no rules changed.')
    text = text.replace(old, new, 1)


replace_once(
"""    function validCompanySite(data) {
      return validCompany(data)
        && (
          documentSiteId(data) in ['OH01', 'OHC']
          || documentSiteId(data) == ''
        );
    }

    function sameCompanyAndSite(data) {
""",
"""    function validCompanySite(data) {
      return validCompany(data)
        && (
          documentSiteId(data) in ['OH01', 'OHC']
          || documentSiteId(data) == ''
        );
    }

    function validAgencyId(agencyId) {
      return agencyId is string
        && agencyId in [
          'sterling_staffing',
          'excel_staffing',
          'lifestyle_staffing'
        ];
    }

    function publicEmployeeAgencyMatches(employeeId, agencyId) {
      return validAgencyId(agencyId)
        && publicEmployeeIsActive(employeeId)
        && get(
          /databases/$(database)/documents/employees/$(employeeId)
        ).data.get('agencyId', '') == agencyId;
    }

    function sameCompanyAndSite(data) {
""",
'insert agency validation helpers',
)

replace_once(
"""    function publicEmployeeWrite() {
      return validCompanySite(request.resource.data)
        && request.resource.data.source == 'auto_created'
        && request.resource.data.status == 'active'
        && request.resource.data.active == true
        && request.resource.data.name is string
        && request.resource.data.name.size() >= 2
        && request.resource.data.name.size() <= 80
        && request.resource.data.nameKey is string
        && request.resource.data.employeeNumber is string;
    }

    function publicPunchCreate() {
""",
"""    function publicEmployeeWrite() {
      return validCompanySite(request.resource.data)
        && validAgencyId(request.resource.data.agencyId)
        && request.resource.data.source == 'auto_created'
        && request.resource.data.status == 'active'
        && request.resource.data.active == true
        && request.resource.data.name is string
        && request.resource.data.name.size() >= 2
        && request.resource.data.name.size() <= 80
        && request.resource.data.nameKey is string
        && request.resource.data.employeeNumber is string;
    }

    function publicEmployeeAgencyAssignment(employeeId) {
      return !signedIn()
        && activeEmployee(resource.data)
        && validCompanySite(request.resource.data)
        && (
          !('agencyId' in resource.data)
          || resource.data.agencyId == null
          || resource.data.agencyId == ''
        )
        && validAgencyId(request.resource.data.agencyId)
        && request.resource.data.agencyAssignmentSource
          == 'worker_public_selection'
        && request.resource.data
          .diff(resource.data)
          .affectedKeys()
          .hasOnly([
            'agencyId',
            'agencyAssignedAt',
            'agencyAssignmentSource',
            'updatedAt'
          ]);
    }

    function publicPunchCreate() {
""",
'add safe blank-agency assignment rule',
)

replace_once(
"""    function publicPunchCreate() {
      return validCompanySite(request.resource.data)
        && request.resource.data.source == 'public_qr'
        && request.resource.data.name is string
        && request.resource.data.name.size() >= 2
        && request.resource.data.name.size() <= 80
        && request.resource.data.nameKey is string
        && request.resource.data.nameKey.size() >= 2
        && request.resource.data.employeeId is string
        && request.resource.data.employeeId.size() > 0
        && publicEmployeeIsActive(request.resource.data.employeeId)
        && request.resource.data.action in [
          'clock_in',
          'start_lunch',
          'end_lunch',
          'clock_out'
        ]
        && request.resource.data.dateKey is string
        && request.resource.data.weekKey is string
        && request.resource.data.timestampMs is number;
    }
""",
"""    function publicPunchCreate() {
      return validCompanySite(request.resource.data)
        && request.resource.data.source == 'public_qr'
        && request.resource.data.name is string
        && request.resource.data.name.size() >= 2
        && request.resource.data.name.size() <= 80
        && request.resource.data.nameKey is string
        && request.resource.data.nameKey.size() >= 2
        && request.resource.data.employeeId is string
        && request.resource.data.employeeId.size() > 0
        && validAgencyId(request.resource.data.agencyId)
        && publicEmployeeAgencyMatches(
          request.resource.data.employeeId,
          request.resource.data.agencyId
        )
        && request.resource.data.action in [
          'clock_in',
          'start_lunch',
          'end_lunch',
          'clock_out'
        ]
        && request.resource.data.dateKey is string
        && request.resource.data.weekKey is string
        && request.resource.data.timestampMs is number;
    }
""",
'bind only publicPunchCreate to employee agency',
)

replace_once(
"""      allow update: if (
          publicEmployeeWrite()
          && activeEmployee(resource.data)
        )
        || (
""",
"""      allow update: if (
          publicEmployeeWrite()
          && activeEmployee(resource.data)
        )
        || publicEmployeeAgencyAssignment(employeeId)
        || (
""",
'allow only guarded public agency assignment',
)

replace_once(
"""      allow read: if isOwnerOrSuperAdmin()
        || (
          validCompanySite(resource.data)
          && (
            publicEmployeeIsActive(
              resource.data.get('employeeId', '')
            )
            || resource.data.get('employeeId', '') == ''
          )
        )
        || (
          activeProfile()
          && (
            (
              (isAdmin() || isManager())
              && inScope(resource.data)
            )
            || workerOwnsRecord(resource.data)
          )
        );
""",
"""      allow read: if isOwnerOrSuperAdmin()
        || (
          activeProfile()
          && (
            (
              (isAdmin() || isManager())
              && inScope(resource.data)
            )
            || workerOwnsRecord(resource.data)
          )
        );
""",
'remove unauthenticated punch-history reads',
)

punch_block = text.split('match /punches/{punchId} {', 1)[1].split(
    'match /punchGuards/{guardId} {', 1
)[0]
for forbidden in [
    "resource.data.get('employeeId', '') == ''",
    'validCompanySite(resource.data)\n          && (\n            publicEmployeeIsActive',
]:
    if forbidden in punch_block:
        raise SystemExit(f'Unsafe public punch read marker remains: {forbidden!r}')

required = [
    'function publicEmployeeAgencyAssignment(employeeId)',
    'function publicEmployeeAgencyMatches(employeeId, agencyId)',
    'function publicPunchCreate()',
    'validAgencyId(request.resource.data.agencyId)',
    'publicEmployeeAgencyAssignment(employeeId)',
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit(f'Missing required secure-rule markers: {missing}')

RULES_PATH.write_text(text, encoding='utf-8')
print('Applied issue #33 secure punch-history rules without modifying stored data.')
