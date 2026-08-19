/**
 * Ties together the cache, parsers, and summarizer for a single file:
 * check cache -> (on miss) parse -> summarize -> cache the result -> return.
 * This is the one function the UI layer needs to call; it doesn't need to
 * know about any of the pieces underneath.
 */
import { getCachedSummary, setCachedSummary } from './cache/summaryCache';
import { checkWebGpuAvailability } from './model/webGpuCheck';
import { DocumentParseError, isSummarizableExtension, parseFile } from './parsers/index';

export type SummarizeFileProgress =
  | { stage: 'cache-hit' }
  | { stage: 'checking-webgpu' }
  | { stage: 'parsing' }
  | { stage: 'loading-model'; modelLoadProgress?: number }
  | { stage: 'summarizing-chunk'; chunkIndex: number; chunkCount: number }
  | { stage: 'combining' };

export interface ISummarizeFileResult {
  summary: string;
  fromCache: boolean;
}

/** Thrown for any expected, user-facing failure condition (unsupported file
 *  type, no WebGPU, empty document, etc). The UI layer should catch this
 *  specifically and show `message` directly - it's already been written to
 *  be shown to the person, not just logged. */
export class SummarizationUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'SummarizationUnavailableError';
    // See DocumentParseError in parsers/types.ts for why this line is
    // necessary under this project's ES5 TypeScript build target.
    Object.setPrototypeOf(this, SummarizationUnavailableError.prototype);
  }
}

export interface ISummarizeFileOptions {
  fileName: string;
  fileId: string;
  fileVersion: string;
  fetchFileBytes: () => Promise<ArrayBuffer>;
  pdfWorkerSrc?: string;
  onProgress?: (progress: SummarizeFileProgress) => void;
}

export async function summarizeFile(options: ISummarizeFileOptions): Promise<ISummarizeFileResult> {
  const { fileName, fileId, fileVersion, fetchFileBytes, onProgress } = options;

  if (!isSummarizableExtension(fileName)) {
    throw new SummarizationUnavailableError(
      'This file type can\u2019t be summarized. Summarization works for PDF, Word, Excel, and PowerPoint files.'
    );
  }

  const cached = await getCachedSummary(fileId, fileVersion);
  if (cached !== undefined) {
    onProgress?.({ stage: 'cache-hit' });
    return { summary: cached, fromCache: true };
  }

  onProgress?.({ stage: 'checking-webgpu' });
  const gpuCheck = await checkWebGpuAvailability();
  if (!gpuCheck.available) {
    throw new SummarizationUnavailableError(
      gpuCheck.reason || 'This device can\u2019t run the summarization feature.'
    );
  }

  onProgress?.({ stage: 'parsing' });
  const fileBytes = await fetchFileBytes();

  let parsed;
  try {
    parsed = await parseFile(fileName, fileBytes, { pdfWorkerSrc: options.pdfWorkerSrc });
  } catch (error) {
    if (error instanceof DocumentParseError) {
      throw new SummarizationUnavailableError(error.message);
    }
    throw error;
  }

  if (parsed.isEffectivelyEmpty) {
    throw new SummarizationUnavailableError(
      'This document doesn\u2019t appear to contain readable text to summarize (it may be a scanned document with no text layer).'
    );
  }

  // Lazy-import: keeps the summarizer (and everything it pulls in, like
  // WebLLM) out of the bundle's critical path until summarization is
  // actually attempted - consistent with how the parsers module already
  // lazy-loads the PDF parser for the same reason.
  const { summarizeDocumentText } = await import(/* webpackChunkName: 'webllm' */ './model/summarizer');

  const summary = await summarizeDocumentText(parsed.text, {
    onProgress: (status) => {
      if (status.stage === 'loading-model') {
        onProgress?.({ stage: 'loading-model', modelLoadProgress: status.modelLoadProgress });
      } else if (status.stage === 'summarizing-chunk') {
        onProgress?.({
          stage: 'summarizing-chunk',
          chunkIndex: status.chunkIndex || 1,
          chunkCount: status.chunkCount || 1,
        });
      } else if (status.stage === 'combining') {
        onProgress?.({ stage: 'combining' });
      }
    },
  });

  await setCachedSummary(fileId, fileVersion, summary);

  return { summary, fromCache: false };
}
