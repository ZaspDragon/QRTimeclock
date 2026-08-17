import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';

const writers = [
  'public-clock-document-id-fix.js',
  'public-clock-permission-hotfix.js',
  'canonical-public-clock.js',
  'stable-public-clock-handler.js',
  'new-worker-first-punch-hotfix.js',
];

const lock = readFileSync(new URL('../punch-writer-lock.js', import.meta.url), 'utf8');
assert(/document\.addEventListener\('click'/.test(lock), 'lock owns the public click listener');
assert(/if \(busy\) return/.test(lock), 'lock prevents concurrent punch attempts');
assert(/writers\.sort\(\(left, right\) => right\.priority - left\.priority\)/.test(lock), 'highest-priority writer wins deterministically');

for (const file of writers) {
  const source = readFileSync(new URL(`../${file}`, import.meta.url), 'utf8');
  assert(/registerPunchWriter\(/.test(source), `${file} registers with the lock`);
  assert(!/document\.addEventListener\('click',[\s\S]{0,220}worker-action-btn/.test(source), `${file} does not install a competing public click listener`);
}

console.log('punch writer lock regression passed');
