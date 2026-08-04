function loadAgencyExportPreviewFix() {
  import('./agency-export-live-preview-v2.js?v=20260804-1').catch((error) => {
    console.warn('Agency export live preview failed to load:', error.message);
  });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', loadAgencyExportPreviewFix, { once: true });
} else {
  loadAgencyExportPreviewFix();
}
