/**
 * Detects whether this browser can run the in-browser summarization model.
 * WebLLM requires WebGPU. Roughly Chrome/Edge 113+ (2023); not available by
 * default in Firefox or older Safari as of this writing. Callers should
 * check this before offering the Summarize button, or at minimum before
 * attempting to load the model, and show a clear fallback message rather
 * than a confusing crash.
 */
export interface IWebGpuAvailability {
  available: boolean;
  reason?: string;
}

export async function checkWebGpuAvailability(): Promise<IWebGpuAvailability> {
  const nav = navigator as Navigator & { gpu?: { requestAdapter: () => Promise<unknown> } };

  if (!nav.gpu) {
    return {
      available: false,
      reason:
        'Your browser doesn\u2019t support WebGPU, which this feature needs to run. Try a recent version of Chrome or Edge.',
    };
  }

  try {
    const adapter = await nav.gpu.requestAdapter();
    if (!adapter) {
      return {
        available: false,
        reason: 'No compatible graphics device was found for WebGPU on this machine.',
      };
    }
  } catch (error) {
    return {
      available: false,
      reason: `Could not initialize WebGPU on this device: ${error instanceof Error ? error.message : 'unknown error'}.`,
    };
  }

  return { available: true };
}
