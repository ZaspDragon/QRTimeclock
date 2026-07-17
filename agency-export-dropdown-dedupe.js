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

  const observer = new MutationObserver(() => dedupeAgencyWorkerOptions(select));
  observer.observe(select, { childList: true, subtree: true });
}

const pageObserver = new MutationObserver(attachAgencyExportDedupe);
pageObserver.observe(document.documentElement, { childList: true, subtree: true });

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', attachAgencyExportDedupe, { once: true });
} else {
  attachAgencyExportDedupe();
}
