/**
 * Manages a single shared WebLLM engine instance for the whole web part.
 *
 * Two lifecycles are deliberately kept separate:
 *  - Model WEIGHTS live in the browser's own model cache (WebLLM handles
 *    this internally via the Cache API), persisting across sessions once
 *    downloaded. We never touch that ourselves.
 *  - The in-memory ENGINE (loaded weights ready for inference) is unloaded
 *    after a period of inactivity to free RAM/VRAM, without re-downloading
 *    anything - reloading from the on-disk cache afterward is fast.
 */
import type { MLCEngine } from '@mlc-ai/web-llm';

export const SUMMARIZATION_MODEL_ID = 'Llama-3.2-3B-Instruct-q4f16_1-MLC';

/** Conservative usable context window, leaving room for the prompt template
 *  and the model's own output. The model's actual context window is 4096
 *  tokens; see chunker.ts for how this bounds chunk sizing. */
export const MODEL_CONTEXT_WINDOW_TOKENS = 4096;

const IDLE_UNLOAD_MS = 5 * 60 * 1000; // 5 minutes

export type ModelLoadProgressHandler = (report: { progress: number; text: string }) => void;

let engineInstance: MLCEngine | undefined;
let loadInFlight: Promise<MLCEngine> | undefined;
let idleTimer: ReturnType<typeof setTimeout> | undefined;

function resetIdleTimer(): void {
  if (idleTimer) clearTimeout(idleTimer);
  idleTimer = setTimeout(() => {
    unloadEngine().catch(() => undefined);
  }, IDLE_UNLOAD_MS);
}

/**
 * Returns a ready-to-use engine, loading it if necessary. Safe to call
 * repeatedly - concurrent callers share the same in-flight load rather than
 * triggering duplicate downloads.
 */
export async function getEngine(onProgress?: ModelLoadProgressHandler): Promise<MLCEngine> {
  resetIdleTimer();

  if (engineInstance) {
    return engineInstance;
  }

  if (loadInFlight) {
    return loadInFlight;
  }

  loadInFlight = (async () => {
    const { CreateMLCEngine } = await import(/* webpackChunkName: 'webllm' */ '@mlc-ai/web-llm');
    const engine = await CreateMLCEngine(SUMMARIZATION_MODEL_ID, {
      initProgressCallback: (report) => {
        onProgress?.({ progress: report.progress, text: report.text });
      },
    });
    engineInstance = engine;
    loadInFlight = undefined;
    return engine;
  })();

  try {
    return await loadInFlight;
  } catch (error) {
    loadInFlight = undefined;
    throw error;
  }
}

/** Frees the in-memory engine (RAM/VRAM) without touching the on-disk model
 *  weight cache. Called automatically after IDLE_UNLOAD_MS of inactivity,
 *  and can be called directly (e.g. on web part disposal). */
export async function unloadEngine(): Promise<void> {
  if (idleTimer) {
    clearTimeout(idleTimer);
    idleTimer = undefined;
  }
  const instanceToUnload = engineInstance;
  if (instanceToUnload) {
    await instanceToUnload.unload();
    // Only clear the shared reference if it still points at the instance we
    // just unloaded - avoids clobbering a newer instance that may have been
    // loaded concurrently while this await was in flight.
    if (engineInstance === instanceToUnload) {
      engineInstance = undefined;
    }
  }
}

/** Testing/advanced use only. */
export function resetEngineStateForTests(): void {
  engineInstance = undefined;
  loadInFlight = undefined;
  if (idleTimer) clearTimeout(idleTimer);
  idleTimer = undefined;
}
