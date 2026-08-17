import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const rules = readFileSync(new URL('../firestore.rules', import.meta.url), 'utf8');

assert(/function hasEditPunchPermission\(\)[\s\S]*isOwnerOrSuperAdmin\(\)[\s\S]*isAdmin\(\)[\s\S]*isManager\(\)[\s\S]*hasPermission\('canEditPunches'\)/.test(rules), 'existing edit-capable roles and canEditPunches remain authorized');
assert(/function isManager\(\)[\s\S]*manager\|supervisor/.test(rules), 'manager and supervisor role aliases remain authorized');
assert(/match \/punches\/\{punchId\}[\s\S]*allow update:[\s\S]*hasEditPunchPermission\(\)[\s\S]*inScope\(resource\.data\)[\s\S]*inScope\(request\.resource\.data\)/.test(rules), 'authenticated punch editing remains available in scope');
assert(/match \/employees\/\{employeeId\}[\s\S]*allow update: if publicEmployeeAgencyAssignment\(employeeId\)/.test(rules), 'anonymous employee updates are limited to blank-agency assignment');
assert(!/allow update: if \([\s\S]{0,100}publicEmployeeWrite\(\)/.test(rules), 'anonymous public worker creation cannot overwrite an existing profile');
assert(/match \/punches\/\{punchId\}[\s\S]*allow delete: if false/.test(rules), 'physical punch deletion remains prohibited');

console.log('edit permission contract regression passed');
