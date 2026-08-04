function loadAgencyExportPreviewFix() {
  import('./agency-export-saved-timesheet-fallback.js?v=20260804-3').catch((error) => {
    console.warn('Agency export preview fix failed to load:', error.message);
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', loadAgencyExportPreviewFix, { once: true });
} else {
  loadAgencyExportPreviewFix();
}
