// Safely collapse duplicate Agency Export options without deleting employee or punch data.
// Identity is based on the rendered worker label: name + agency + branch.

const SELECT_ID = 'agencyLegacyWorkerSelect';

function normalizeLabel(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

function dedupeAgencyWorkerOptions(select) {
  if (!select || select.dataset.dedupeRunning === 'true') return;
  select.dataset.dedupeRunning = 'true';

  try {
    const selectedValue = select.value;
    const seen = new Map();
    const duplicates = [];

    [...select.options].forEach((option) => {
      if (!option.value) return;
      const key = normalizeLabel(option.textContent);
      if (!key) return;

      if (!seen.has(key)) {
        seen.set(key, option);
        return;
      }

      const kept = seen.get(key);
      const aliases = JSON.parse(kept.dataset.duplicateWorkerValues || '[]');
      aliases.push(option.value);
      kept.dataset.duplicateWorkerValues = JSON.stringify([...new Set(aliases)]);
      duplicates.push(option);
    });

    duplicates.forEach((option) => option.remove());

    if (selectedValue && [...select.options].some((option) => option.value === selectedValue)) {
      select.value = selectedValue;
    }

    if (duplicates.length) {
      console.info(`[QRTimeclock] Collapsed ${duplicates.length} duplicate Agency Export worker option(s).`);
    }
  } finally {
    delete select.dataset.dedupeRunning;
  }
}

function attachAgencyExportDedupe() {
  const select = document.getElementById(SELECT_ID);
  if (!select || select.dataset.dedupeAttached === 'true') return;

  select.dataset.dedupeAttached = 'true';
  dedupeAgencyWorkerOptions(select);

  // Debounced: the dedupe itself removes options, which would otherwise
  // re-trigger this observer on every render.
  let queued = false;
  const observer = new MutationObserver(() => {
    if (queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      dedupeAgencyWorkerOptions(select);
    });
  });
  observer.observe(select, { childList: true, subtree: true });

  // The select exists now, so stop scanning the whole page.
  pageObserver.disconnect();
}

// Waits for the Agency Export select to appear, then disconnects. Previously
// this stayed connected for the life of the page and ran on every single DOM
// mutation, which made manager tables (live punches, timesheets) stutter.
let pageScanQueued = false;
const pageObserver = new MutationObserver(() => {
  if (pageScanQueued) return;
  pageScanQueued = true;
  requestAnimationFrame(() => {
    pageScanQueued = false;
    attachAgencyExportDedupe();
  });
});
pageObserver.observe(document.documentElement, { childList: true, subtree: true });

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', attachAgencyExportDedupe, { once: true });
} else {
  attachAgencyExportDedupe();
}
