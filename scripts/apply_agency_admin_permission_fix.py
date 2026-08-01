from pathlib import Path

path = Path('app.js')
text = path.read_text(encoding='utf-8')

old = """const PUNCH_EDIT_ROLES = new Set([
  'admin',
  'manager',
  'supervisor',
  'super_admin',
  'superadmin',
  'owner',
]);
"""
new = """const PUNCH_EDIT_ROLES = new Set([
  'admin',
  'agency_admin',
  'manager',
  'supervisor',
  'super_admin',
  'superadmin',
  'owner',
]);
"""

count = text.count(old)
if count != 1:
    raise SystemExit(f'Expected one PUNCH_EDIT_ROLES block, found {count}; app.js was not changed.')
text = text.replace(old, new, 1)

required = [
    "'agency_admin',\n  'manager'",
    "canEditPunches: fullAccess || ['admin', 'agency_admin', 'manager', 'supervisor'].includes(normalized)",
    "function canEditPunches()",
    "function isManager()",
]
missing = [marker for marker in required if marker not in text]
if missing:
    raise SystemExit(f'Permission validation failed; missing: {missing}')

path.write_text(text, encoding='utf-8')
print('Agency admins now receive manager/edit UI access while retaining agency and branch scoping.')
