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

  document.querySelectorAll('.worker-action-btn[data-action="start_lunch"]').forEach((button) => {
    button.textContent = 'Start Lunch';
  });
  document.querySelectorAll('.worker-action-btn[data-action="end_lunch"]').forEach((button) => {
    button.textContent = 'End Lunch';
  });
  document.querySelectorAll('option[value="start_lunch"]').forEach((option) => {
    option.textContent = 'Start Lunch';
  });
  document.querySelectorAll('option[value="end_lunch"]').forEach((option) => {
    option.textContent = 'End Lunch';
  });
}

function initializeLunchLabels() {
  replaceLunchText();
  [0, 100, 300, 750, 1500, 3000].forEach((delay) => {
    window.setTimeout(() => replaceLunchText(), delay);
  });

  // Batched to one pass per animation frame. Previously every inserted node
  // triggered its own full querySelectorAll sweep, which is expensive when the
  // manager tables render hundreds of rows at once.
  let pending = [];
  let queued = false;
  const observer = new MutationObserver((mutations) => {
    for (const mutation of mutations) {
      mutation.addedNodes.forEach((node) => {
        if (node.nodeType === Node.ELEMENT_NODE) pending.push(node);
      });
    }
    if (!pending.length || queued) return;
    queued = true;
    requestAnimationFrame(() => {
      queued = false;
      const nodes = pending;
      pending = [];
      nodes.forEach((node) => {
        if (node.isConnected) replaceLunchText(node);
      });
    });
  });

  observer.observe(document.body, { childList: true, subtree: true });
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', initializeLunchLabels, { once: true });
} else {
  initializeLunchLabels();
}

import('./correction-dashboard.js?v=20260717-1').catch((error) => {
  console.warn('Trial correction dashboard failed to load:', error.message);
});
