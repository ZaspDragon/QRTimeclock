import { firebaseConfig } from './firebase-config.js';
import { getApp, getApps, initializeApp } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-app.js';
import { getAuth } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-auth.js';
import { doc, getFirestore, serverTimestamp, updateDoc } from 'https://www.gstatic.com/firebasejs/10.12.5/firebase-firestore.js';

const app = getApps().length ? getApp() : initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);
let busy = false;

function showMessage(message, isError = false) {
  const status = document.getElementById('timesheetStatus')
    || document.getElementById('workerStatusMessage');
  if (status) {
    status.textContent = message;
    if (isError) status.style.borderColor = 'rgba(255,90,90,.6)';
  }
  if (typeof window.toast === 'function') window.toast(message, isError);
  else if (isError) window.alert(message);
}

function findDeleteButton(target) {
  return target instanceof Element
    ? target.closest('.manager-delete-punch-btn')
    : null;
}

async function mobileSoftDelete(button) {
  const punchId = String(button?.dataset?.id || '').trim();
  if (!punchId) throw new Error('This punch could not be identified. Refresh and try again.');

  const user = auth.currentUser;
  if (!user) throw new Error('Your manager session expired. Sign in again before deleting a punch.');

  const reason = window.prompt('Reason required for deleting this punch:');
  if (!String(reason || '').trim()) return;
  if (!window.confirm('Mark this punch deleted while preserving its history?')) return;

  button.disabled = true;
  button.setAttribute('aria-busy', 'true');

  await updateDoc(doc(db, 'punches', punchId), {
    status: 'deleted',
    active: false,
    deletedAt: serverTimestamp(),
    deletedBy: user.email || user.uid,
    deleteReason: String(reason).trim().slice(0, 300),
    updatedAt: serverTimestamp(),
  });

  const row = button.closest('tr');
  if (row) row.remove();
  showMessage('Punch marked deleted. Historical data was preserved.');
}

function install() {
  const handle = async (event) => {
    const button = findDeleteButton(event.target);
    if (!button || busy) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    busy = true;
    try {
      await mobileSoftDelete(button);
    } catch (error) {
      console.error('[mobile-punch-editor-actions]', error);
      showMessage(error?.message || 'Could not delete punch.', true);
      button.disabled = false;
      button.removeAttribute('aria-busy');
    } finally {
      busy = false;
    }
  };

  document.addEventListener('click', handle, true);
  document.addEventListener('touchend', handle, { capture: true, passive: false });
  console.info('[QRTimeclock] Mobile punch editor actions installed.');
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', install, { once: true });
} else {
  install();
}
