const LABELS = new Map([
  ['Lunch Out', 'Start Lunch'],
  ['Lunch In', 'End Lunch']
]);

function replaceLunchText(root = document) {
  root.querySelectorAll('button, option, th, td, span, strong, label, p').forEach((element) => {
    const current = element.textContent?.trim();
    if (LABELS.has(current)) {
      element.textContent = LABELS.get(current);
    }
  });
}

function initializeLunchLabels() {
  replaceLunchText();

  const observer = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType !== Node.ELEMENT_NODE) return;
        const element = /** @type {Element} */ (node);
        replaceLunchText(element);
        const current = element.textContent?.trim();
        if (LABELS.has(current) && !element.children.length) {
          element.textContent = LABELS.get(current);
        }
      });
    }
  });

  observer.observe(document.body, { childList: true, subtree: true });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeLunchLabels, { once: true });
} else {
  initializeLunchLabels();
}
