/**
 * Splits document text into chunks that fit the summarization model's
 * context window, for the map-reduce summarization strategy (see
 * summarizer.ts). Splits on paragraph/section boundaries where possible
 * rather than mid-sentence, so each chunk reads coherently on its own.
 */

/** Rough, fast token estimate (no model/tokenizer dependency needed at
 *  chunking time). English prose averages ~4 characters per token; this
 *  intentionally errs slightly conservative (estimates a bit high) so we
 *  don't overflow the model's real context window on the boundary case. */
function estimateTokens(text: string): number {
  return Math.ceil(text.length / 3.5);
}

export interface IChunkOptions {
  /** Max tokens per chunk. Should leave headroom below the model's real
   *  context window for the prompt template and generated output. */
  maxTokensPerChunk: number;
}

export function chunkText(text: string, options: IChunkOptions): string[] {
  const trimmed = text.trim();
  if (trimmed.length === 0) return [];

  if (estimateTokens(trimmed) <= options.maxTokensPerChunk) {
    return [trimmed];
  }

  // Split on paragraph boundaries first (handles the common case of
  // prose/docx/pdf text cleanly), falling back to sentence boundaries for
  // any single paragraph that's still too large on its own (e.g. a dense
  // spreadsheet row dump with no paragraph breaks).
  const paragraphs = trimmed.split(/\n{2,}/);
  const chunks: string[] = [];
  let current = '';

  const flush = (): void => {
    if (current.trim().length > 0) {
      chunks.push(current.trim());
      current = '';
    }
  };

  for (const paragraph of paragraphs) {
    const candidate = current.length > 0 ? `${current}\n\n${paragraph}` : paragraph;

    if (estimateTokens(candidate) <= options.maxTokensPerChunk) {
      current = candidate;
      continue;
    }

    // Adding this whole paragraph would overflow the current chunk.
    flush();

    if (estimateTokens(paragraph) <= options.maxTokensPerChunk) {
      current = paragraph;
    } else {
      // Even a single paragraph is too large on its own - break it up by
      // sentence instead so we never emit a chunk the model can't accept.
      const sentences = paragraph.split(/(?<=[.!?])\s+/);
      let sentenceBuffer = '';
      for (const sentence of sentences) {
        const sentenceCandidate = sentenceBuffer.length > 0 ? `${sentenceBuffer} ${sentence}` : sentence;
        if (estimateTokens(sentenceCandidate) <= options.maxTokensPerChunk) {
          sentenceBuffer = sentenceCandidate;
        } else {
          if (sentenceBuffer.trim().length > 0) chunks.push(sentenceBuffer.trim());
          sentenceBuffer = sentence;
        }
      }
      if (sentenceBuffer.trim().length > 0) chunks.push(sentenceBuffer.trim());
    }
  }

  flush();
  return chunks;
}
