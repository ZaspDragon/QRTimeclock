// Single deterministic owner for public clock (.worker-action-btn) taps.
//
// Why this exists:
// Five separate hotfix modules each installed their own capture-phase
// document click listener and called stopImmediatePropagation(). Whichever
// module happened to finish loading first won the tap. Because those modules
// are loaded through async dynamic imports (and some wait for DOMContentLoaded
// while others install immediately), the winner changed from page load to page
// load — especially on slow phones. That is why punches behaved differently on
// different devices/refreshes.
//
// Now every writer registers itself here with an explicit priority and this
// module owns the only click listener. The highest-priority registered writer
// that can handle the tap runs, every time, regardless of load order.

const writers = [];
let busy = false;

export const PUNCH_WRITER_PRIORITY = {
  DOCUMENT_ID_FIX: 100,      // newest rules-safe writer (2026-08-11)
  PERMISSION_HOTFIX: 90,
  CANONICAL: 80,
  STABLE: 70,
  NEW_WORKER_FALLBACK: 60    // only applies when its own canHandle() says so
};

/**
 * @param {string} name        diagnostic label
 * @param {number} priority    higher wins
 * @param {(action: string, button: Element) => Promise<void>} handler
 * @param {(button: Element) => boolean} [canHandle]
 */
export function registerPunchWriter(name, priority, handler, canHandle) {
  writers.push({ name, priority, handler, canHandle });
  writers.sort((a, b) => b.priority - a.priority);
  console.info(`[punch-writer-lock] registered "${name}" (priority ${priority})`);
}

export function activePunchWriterName() {
  return writers[0]?.name || null;
}

function pickWriter(button) {
  return writers.find((writer) => {
    try {
      return typeof writer.canHandle === 'function' ? writer.canHandle(button) : true;
    } catch (error) {
      console.warn(`[punch-writer-lock] canHandle failed for ${writer.name}:`, error?.message);
      return false;
    }
  }) || null;
}

document.addEventListener('click', async (event) => {
  const button = event.target?.closest?.('.worker-action-btn');
  if (!button) return;

  // No hotfix writer loaded (yet): let the built-in app.js handler run.
  if (!writers.length) return;

  const writer = pickWriter(button);
  // No hotfix writer applies to this tap (e.g. the worker card is hidden):
  // leave the event alone so app.js's built-in handler still works.
  if (!writer) return;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();

  // Hard double-tap guard: one punch attempt at a time, no matter how many
  // times a worker taps a laggy phone screen.
  if (busy) return;

  busy = true;
  try {
    await writer.handler(String(button.dataset.action || ''), button);
  } catch (error) {
    console.error(`[punch-writer-lock] ${writer.name} failed:`, error);
  } finally {
    busy = false;
  }
}, true);
