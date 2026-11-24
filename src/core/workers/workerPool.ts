/**
 * Lightweight Web Worker task registry for Phase 1–4 heavy client-side jobs
 * (PDF rendering/merging, OCR via Tesseract.js, OOXML parsing, and ZIP/conversion tasks).
 */
export interface WorkerTaskHandle {
  terminate: () => void;
}

export function isWebWorkerSupported(): boolean {
  return typeof Worker !== 'undefined';
}

export function spawnModuleWorker(workerFactory: () => Worker): Worker | null {
  if (!isWebWorkerSupported()) return null;
  try {
    return workerFactory();
  } catch {
    return null;
  }
}
