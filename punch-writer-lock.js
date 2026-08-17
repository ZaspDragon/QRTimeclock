// Deterministic owner for public clock taps. Legacy writers register here
// instead of competing with capture-phase document listeners.
const writers = [];
let busy = false;

export const PUNCH_WRITER_PRIORITY = {
  DOCUMENT_ID_FIX: 100,
  PERMISSION_HOTFIX: 90,
  CANONICAL: 80,
  STABLE: 70,
  NEW_WORKER_FALLBACK: 60,
};

export function registerPunchWriter(name, priority, handler, canHandle) {
  const existing = writers.findIndex((writer) => writer.name === name);
  const registration = { name, priority, handler, canHandle };
  if (existing >= 0) writers.splice(existing, 1, registration);
  else writers.push(registration);
  writers.sort((left, right) => right.priority - left.priority);
  console.info(`[punch-writer-lock] registered "${name}" (priority ${priority})`);
}

export function activePunchWriterName(button = null) {
  return pickWriter(button)?.name || null;
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
  if (!button || !writers.length) return;
  const writer = pickWriter(button);
  if (!writer) return;

  event.preventDefault();
  event.stopPropagation();
  event.stopImmediatePropagation();
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
