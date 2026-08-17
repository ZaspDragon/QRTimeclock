import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const source = readFileSync(new URL('../mobile-punch-editor-actions.js', import.meta.url), 'utf8');

assert(/Reason required for this punch edit/.test(source), 'active punch editor asks for a reason');
assert(/editReason\.length < 2/.test(source), 'blank edit reasons are rejected');
assert(/const batch = writeBatch\(db\)/.test(source), 'punch and edit history use a batch');
assert(/batch\.update\(doc\(db, 'punches', punchId\), updated\)/.test(source), 'batch updates the punch');
assert(/batch\.set\(editRef,[\s\S]*original:[\s\S]*updated:/.test(source), 'batch preserves original and updated values');
assert(/editedByUid: user\.uid/.test(source), 'editor identity is recorded');
assert(/reason: editReason[\s\S]*editReason/.test(source), 'edit reason is recorded in compatible fields');
assert(/await batch\.commit\(\)/.test(source), 'edit is committed atomically');

console.log('edit audit regression passed');
