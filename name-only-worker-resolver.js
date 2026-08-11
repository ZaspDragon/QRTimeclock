import { findPublicWorkerMatches, chooseCanonicalPublicWorker, employeeName, employeeSite, employeeAgency } from './public-worker-lookup-v3.js';

let requestId = 0;

function selectedSite() {
  const querySite = String(new URLSearchParams(location.search).get('site') || '').trim().toUpperCase();
  if (querySite === 'OH01' || querySite === 'OHC') return querySite;
  return String(document.getElementById('workerBranchSelect')?.value || 'OH01').trim().toUpperCase();
}

function selectedAgency() {
  return String(document.getElementById('workerAgencySelect')?.value || localStorage.getItem('workerPunchAgency') || '').trim();
}

function setStatus(message, isError = false) {
  const status = document.getElementById('workerLookupStatus');
  if (!status) return;
  status.textContent = message;
  status.style.borderColor = isError ? 'rgba(255,90,90,.55)' : 'rgba(43,213,118,.45)';
}

function hideCreateNewSuggestion() {
  document.querySelectorAll('#workerAutocompleteList button, #workerAutocompleteList [role="option"]').forEach((node) => {
    if (/create new/i.test(node.textContent || '')) node.remove();
  });
}

async function resolveTypedName() {
  const input = document.getElementById('workerNameInput');
  const typed = String(input?.value || '').trim();
  const currentRequest = ++requestId;
  if (typed.length < 2) return;

  window.setTimeout(async () => {
    if (currentRequest !== requestId) return;
    try {
      const siteId = selectedSite();
      const agencyId = selectedAgency();
      const matches = await findPublicWorkerMatches(typed, siteId, agencyId);
      if (currentRequest !== requestId) return;
      hideCreateNewSuggestion();

      const match = chooseCanonicalPublicWorker(matches);
      if (match) {
        const branch = employeeSite(match);
        const agency = employeeAgency(match);
        const branchSelect = document.getElementById('workerBranchSelect');
        const agencySelect = document.getElementById('workerAgencySelect');

        if (branch && branchSelect && [...branchSelect.options].some((option) => option.value === branch)) {
          branchSelect.disabled = false;
          branchSelect.value = branch;
        }
        if (agency && agencySelect && [...agencySelect.options].some((option) => option.value === agency)) {
          agencySelect.value = agency;
          localStorage.setItem('workerPunchAgency', agency);
        }

        input.value = employeeName(match) || typed;
        localStorage.setItem('workerPunchName', input.value);
        setStatus(`✓ Found: ${input.value}. Ready to punch.`);
        return;
      }

      if (matches.length > 1) {
        setStatus('More than one separate worker uses that exact name. Ask a manager to select the correct profile before punching.', true);
        return;
      }

      // A failed pre-check must not block a legitimate new temp from punching.
      setStatus('Ready to punch. If this is your first punch, your worker profile will be created automatically.');
    } catch (error) {
      console.warn('[name-only-worker-resolver]', error);
      // The punch handler performs its own final verification. Keep this lookup
      // informational so a temporary directory read problem does not disable all temps.
      setStatus('Ready to punch. Worker identity will be verified when you tap a punch button.');
    }
  }, 300);
}

function install() {
  const input = document.getElementById('workerNameInput');
  if (!input || input.dataset.nameOnlyResolverV3 === 'true') return;
  input.dataset.nameOnlyResolverV3 = 'true';
  input.addEventListener('input', resolveTypedName);
  input.addEventListener('change', resolveTypedName);
  input.addEventListener('blur', resolveTypedName);
  if (input.value.trim()) resolveTypedName();
}

if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', install, { once: true });
else install();

console.info('[QRTimeclock] Rules-safe name resolver installed.');
