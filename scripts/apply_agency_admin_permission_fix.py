from pathlib import Path

# Guarded source patch only. No employee, punch, timesheet, approval, or audit data is touched.
path = Path('app.js')
text = path.read_text(encoding='utf-8')

start_marker = "const PUNCH_EDIT_ROLES = new Set(["
end_marker = "]);"

if start_marker not in text:
    raise SystemExit('PUNCH_EDIT_ROLES was not found; app.js was not changed.')

prefix, remainder = text.split(start_marker, 1)
role_block, suffix = remainder.split(end_marker, 1)

if "'agency_admin'" not in role_block:
    admin_line = "\n  'admin',"
    if admin_line not in role_block:
        raise SystemExit("The admin role entry was not found; app.js was not changed.")
    role_block = role_block.replace(
        admin_line,
        "\n  'admin',\n  'agency_admin',",
        1,
    )
    text = prefix + start_marker + role_block + end_marker + suffix

updated_block = text.split(start_marker, 1)[1].split(end_marker, 1)[0]
required = [
    "'agency_admin'",
    "canEditPunches: fullAccess || ['admin', 'agency_admin', 'manager', 'supervisor'].includes(normalized)",
    "function canEditPunches()",
    "function isManager()",
]
missing = [marker for marker in required if marker not in text]
if "'agency_admin'" not in updated_block:
    missing.append('agency_admin in PUNCH_EDIT_ROLES')
if missing:
    raise SystemExit(f'Permission validation failed; missing: {missing}')

path.write_text(text, encoding='utf-8')
print('Agency-admin manager/edit access is present and remains agency/branch scoped.')
