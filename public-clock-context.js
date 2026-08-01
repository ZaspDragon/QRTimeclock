const VALID_SITES = new Set(['OH01', 'OHC']);
const AGENCY_LABELS = new Map([
  ['sterling_staffing', 'Sterling Staffing'],
  ['excel_staffing', 'Excel Staffing'],
  ['lifestyle_staffing', 'Lifestyle Staffing'],
]);

const INITIALIZATION_DELAYS_MS = [0, 100, 250, 500, 1000, 2000, 4000];

function requestedSiteId() {
  const value = String(new URLSearchParams(window.location.search).get('site') || '')
    .trim()
    .toUpperCase();
  return VALID_SITES.has(value) ? value : '';
}

function ensureAgencyControl() {
  let agencySelect = document.getElementById('workerAgencySelect');
  if (!agencySelect) {
    const branchSelect = document.getElementById('workerBranchSelect');
    const branchLabel = branchSelect?.closest('label');
    if (!branchLabel) return null;

    const agencyField = document.createElement('label');
    agencyField.id = 'workerAgencyField';
    agencyField.innerHTML = `
      <span>Staffing agency <small>(required for every temp)</small></span>
      <select id="workerAgencySelect" required aria-required="true">
        <option value="">Choose your staffing agency</option>
        ${[...AGENCY_LABELS.entries()]
          .map(([value, label]) => `<option value="${value}">${label}</option>`)
          .join('')}
      </select>
    `;
    branchLabel.insertAdjacentElement('afterend', agencyField);
    agencySelect = agencyField.querySelector('select');
  }

  return agencySelect;
}

function applyBranchContext() {
  const siteId = requestedSiteId();
  const branchSelect = document.getElementById('workerBranchSelect');
  if (!branchSelect) return false;

  if (siteId) {
    if (branchSelect.value !== siteId) {
      branchSelect.value = siteId;
      branchSelect.dispatchEvent(new Event('change', { bubbles: true }));
    }
    if (!branchSelect.disabled) branchSelect.disabled = true;
    branchSelect.dataset.urlLocked = 'true';

    const ariaLabel = `Branch locked to ${siteId}`;
    if (branchSelect.getAttribute('aria-label') !== ariaLabel) {
      branchSelect.setAttribute('aria-label', ariaLabel);
    }

    const labelText = branchSelect.closest('label')?.querySelector('span');
    const desiredLabel = `Branch (${siteId} — automatically selected)`;
    if (labelText && labelText.textContent !== desiredLabel) {
      labelText.textContent = desiredLabel;
    }

    const signupSite = document.getElementById('signupSiteInput');
    if (signupSite && signupSite.value !== siteId) signupSite.value = siteId;
  }

  return true;
}

function applyAgencyContext() {
  const agencySelect = ensureAgencyControl();
  if (!agencySelect) return false;

  const firstOption = agencySelect.options[0];
  if (firstOption && firstOption.textContent !== 'Choose your staffing agency') {
    firstOption.textContent = 'Choose your staffing agency';
  }
  agencySelect.required = true;
  agencySelect.setAttribute('aria-required', 'true');

  AGENCY_LABELS.forEach((label, value) => {
    if (![...agencySelect.options].some((option) => option.value === value)) {
      agencySelect.add(new Option(label, value));
    }
  });

  const fieldLabel = agencySelect.closest('label')?.querySelector('span');
  const desiredLabelHtml = 'Staffing agency <small>(required for every temp)</small>';
  if (fieldLabel && fieldLabel.innerHTML !== desiredLabelHtml) {
    fieldLabel.innerHTML = desiredLabelHtml;
  }

  const savedAgency = String(localStorage.getItem('workerPunchAgency') || '').trim();
  if (!agencySelect.value && AGENCY_LABELS.has(savedAgency)) {
    agencySelect.value = savedAgency;
  }

  if (agencySelect.dataset.agencyContextBound !== 'true') {
    agencySelect.dataset.agencyContextBound = 'true';
    agencySelect.addEventListener('change', () => {
      if (AGENCY_LABELS.has(agencySelect.value)) {
        localStorage.setItem('workerPunchAgency', agencySelect.value);
      } else {
        localStorage.removeItem('workerPunchAgency');
      }
    });
  }

  return true;
}

function applyPublicClockContext() {
  applyBranchContext();
  applyAgencyContext();
}

function initializePublicClockContext() {
  // Bounded retries handle controls created by other modules without observing and
  // rewriting the same DOM nodes indefinitely. The previous MutationObserver loop
  // could starve mobile click and input events.
  INITIALIZATION_DELAYS_MS.forEach((delay) => {
    window.setTimeout(applyPublicClockContext, delay);
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializePublicClockContext, { once: true });
} else {
  initializePublicClockContext();
}
