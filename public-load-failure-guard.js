// Keeps a denied public live-history read from blocking clock buttons.
// Firestore intentionally denies unauthenticated punch-history reads; this guard
// converts that expected listener error into a neutral status so the secured
// one-click punch fallback can verify the employee and save normally.

const STATUS_ID = 'workerLookupStatus';

function repairPublicStatus() {
  const status = document.getElementById(STATUS_ID);
  if (!status) return;

  const message = String(status.textContent || '').trim().toLowerCase();
  if (message === 'load failed' || message.includes('permission-denied')) {
    status.textContent = 'Could not load live status. Your name will be verified when you tap a punch button.';
    status.style.borderColor = '';
  }
}

function installGuard() {
  repairPublicStatus();
  const status = document.getElementById(STATUS_ID);
  if (!status || status.dataset.loadFailureGuard === 'true') return;

  status.dataset.loadFailureGuard = 'true';
  const observer = new MutationObserver(() => repairPublicStatus());
  observer.observe(status, { childList: true, characterData: true, subtree: true });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', installGuard, { once: true });
} else {
  installGuard();
}

window.setTimeout(installGuard, 250);
window.setTimeout(installGuard, 1000);
