function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}

function showStatus(message, isError = false) {
  const preview = document.getElementById('agencyPreview');
  if (!preview) return;
  const box = document.createElement('div');
  box.className = 'empty-state';
  box.textContent = message;
  if (isError) box.setAttribute('role', 'alert');
  preview.replaceChildren(box);
}

function printableSheet() {
  return document.getElementById('agencyPrintableSheet');
}

function printLiveSheet(event) {
  const button = event.target instanceof Element
    ? event.target.closest('#agencyLegacyPrintBtn')
    : null;
  if (!button) return;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();

  const sheet = printableSheet();
  if (!sheet) {
    showStatus('Preview the worker sheet first, then choose Print / Save PDF.', true);
    return;
  }

  const worker = sheet.querySelector('h2')?.textContent || 'Weekly Time Sheet';
  const win = window.open('', '_blank', 'width=1100,height=800');
  if (!win) {
    showStatus('Pop-up blocked. Allow pop-ups, then choose Print / Save PDF again.', true);
    return;
  }

  win.document.write(`<!doctype html><html><head><meta charset="utf-8"><title>${escapeHtml(worker)}</title><style>
    @page { size: auto; margin: 0.5in; }
    body { font-family: Arial, sans-serif; margin: 0; color: #111; background: #fff; }
    #agencyPrintableSheet { box-shadow: none !important; border-radius: 0 !important; padding: 0 !important; }
    table { page-break-inside: avoid; }
    button { display: none !important; }
  </style></head><body>${sheet.outerHTML}<script>
    window.addEventListener('load', function () {
      window.focus();
      window.print();
    });
  <\/script></body></html>`);
  win.document.close();
}

document.addEventListener('click', printLiveSheet, true);
console.info('[QRTimeclock] Live agency sheet print handler installed.');
