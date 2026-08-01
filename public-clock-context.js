const VALID_SITES = new Set(['OH01', 'OHC']);
const AGENCY_LABELS = new Map([
  ['sterling_staffing', 'Sterling Staffing'],
  ['excel_staffing', 'Excel Staffing'],
  ['lifestyle_staffing', 'Lifestyle Staffing'],
]);

function requestedSiteId() {
  const value = String(new URLSearchParams(window.location.search).get('site') || '').trim().toUpperCase();
  return VALID_SITES.has(value) ? value : '';
}

function applyBranchContext() {
  const siteId = requestedSiteId();
  if (!siteId) return;

  const branchSelect = document.getElementById('workerBranchSelect');
  if (branchSelect) {
    if (branchSelect.value !== siteId) {
      branchSelect.value = siteId;
      branchSelect.dispatchEvent(new Event('change', { bubbles: true }));
    }
    branchSelect.disabled = true;
    branchSelect.dataset.urlLocked = 'true';
    branchSelect.setAttribute('aria-label', `Branch locked to ${siteId}`);
    const labelText = branchSelect.closest('label')?.querySelector('span');
    if (labelText) labelText.textContent = `Branch (${siteId} — automatically selected)`;
  }

  const signupSite = document.getElementById('signupSiteInput');
  if (signupSite) signupSite.value = siteId;
}

function applyAgencyContext() {
  const agencySelect = document.getElementById('workerAgencySelect');
  if (!agencySelect) return;

  const firstOption = agencySelect.options[0];
  if (firstOption) firstOption.textContent = 'Choose your staffing agency';
  agencySelect.required = true;
  agencySelect.setAttribute('aria-required', 'true');

  AGENCY_LABELS.forEach((label, value) => {
    if (![...agencySelect.options].some((option) => option.value === value)) {
      agencySelect.add(new Option(label, value));
    }
  });

  const fieldLabel = agencySelect.closest('label')?.querySelector('span');
  if (fieldLabel) fieldLabel.innerHTML = 'Staffing agency <small>(required for every temp)</small>';

  const savedAgency = String(localStorage.getItem('workerPunchAgency') || '').trim();
  if (!agencySelect.value && AGENCY_LABELS.has(savedAgency)) agencySelect.value = savedAgency;

  agencySelect.addEventListener('change', () => {
    if (AGENCY_LABELS.has(agencySelect.value)) {
      localStorage.setItem('workerPunchAgency', agencySelect.value);
    }
  });
}

function applyPublicClockContext() {
  applyBranchContext();
  applyAgencyContext();
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', applyPublicClockContext, { once: true });
} else {
  applyPublicClockContext();
}

const observer = new MutationObserver(() => applyPublicClockContext());
observer.observe(document.documentElement, { childList: true, subtree: true });
window.setTimeout(() => observer.disconnect(), 15000);
