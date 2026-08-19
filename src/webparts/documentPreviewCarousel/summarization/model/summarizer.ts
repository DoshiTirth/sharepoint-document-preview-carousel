/**
 * Map-reduce document summarization: summarize each chunk individually
 * ("map"), then summarize the combined chunk-summaries into one final
 * structured summary ("reduce"). Necessary because the model's context
 * window (4096 tokens) is far smaller than many real documents.
 *
 * Output length scales with the original document's size: short documents
 * get a brief summary, long documents get a fuller one with more key points
 * - both as structured headers + bullet points, per the agreed output style.
 */
import type { MLCEngine } from '@mlc-ai/web-llm';
import { chunkText } from './chunker';
import { getEngine, MODEL_CONTEXT_WINDOW_TOKENS, ModelLoadProgressHandler } from './modelEngine';

// Reserve headroom below the model's real context window for the prompt
// template and the model's generated output, for both the map and reduce steps.
const PROMPT_OVERHEAD_TOKENS = 250;
const MAP_OUTPUT_RESERVE_TOKENS = 400;

const MAX_CHUNK_TOKENS = MODEL_CONTEXT_WINDOW_TOKENS - PROMPT_OVERHEAD_TOKENS - MAP_OUTPUT_RESERVE_TOKENS;

export type SummarizationProgressHandler = (status: {
  stage: 'loading-model' | 'summarizing-chunk' | 'combining';
  /** Only present during 'loading-model'. */
  modelLoadProgress?: number;
  /** Only present during 'summarizing-chunk'. */
  chunkIndex?: number;
  chunkCount?: number;
}) => void;

export interface ISummarizeOptions {
  onProgress?: SummarizationProgressHandler;
  /** Injected for testing; defaults to the real shared engine. */
  engineProvider?: (onProgress?: ModelLoadProgressHandler) => Promise<MLCEngine>;
}

async function runCompletion(engine: MLCEngine, systemPrompt: string, userContent: string): Promise<string> {
  const response = await engine.chat.completions.create({
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userContent },
    ],
    temperature: 0.3,
  });

  const text = response.choices[0]?.message?.content;
  return (text || '').trim();
}

function targetKeyPointCount(documentLength: number): number {
  // Scales the requested level of detail with document size, per the
  // agreed design ("short document -> short summary, long -> fuller one").
  if (documentLength < 2000) return 3;
  if (documentLength < 10000) return 5;
  if (documentLength < 40000) return 8;
  return 12;
}

const MAP_STEP_SYSTEM_PROMPT =
  'You are summarizing one section of a larger document. Write a concise, factual summary of ' +
  'the key information in this section, in a few sentences. Do not add commentary, opinions, or ' +
  'information not present in the text.';

function buildDirectSummaryPrompt(documentLength: number): string {
  const points = targetKeyPointCount(documentLength);
  return (
    'Summarize the following document. Respond in Markdown with a short "## Summary" heading ' +
    `followed by a brief overview paragraph, then a "## Key Points" heading with up to ${points} ` +
    'bullet points covering the most important information. Be factual and concise; do not add ' +
    'information not present in the text.'
  );
}

function buildReduceStepPrompt(documentLength: number): string {
  const points = targetKeyPointCount(documentLength);
  return (
    'You are combining section summaries from a larger document into one final summary. The ' +
    'sections are provided in their original document order. Respond in Markdown with a short ' +
    `"## Summary" heading (one short paragraph capturing the overall document), followed by a ` +
    `"## Key Points" heading with up to ${points} bullet points covering the most important ` +
    'information across all sections. Merge related points from different sections rather than ' +
    'listing them separately. Be factual and concise; do not add information not present in the ' +
    'provided section summaries.'
  );
}

export async function summarizeDocumentText(
  documentText: string,
  options: ISummarizeOptions = {}
): Promise<string> {
  const getEngineFn = options.engineProvider || getEngine;

  options.onProgress?.({ stage: 'loading-model' });
  const engine = await getEngineFn((report) => {
    options.onProgress?.({ stage: 'loading-model', modelLoadProgress: report.progress });
  });

  const chunks = chunkText(documentText, { maxTokensPerChunk: MAX_CHUNK_TOKENS });

  if (chunks.length === 0) {
    return '## Summary\n\nThis document doesn\u2019t contain enough readable text to summarize.';
  }

  // Single chunk: skip the map step entirely and go straight to a
  // structured summary in one model call.
  if (chunks.length === 1) {
    options.onProgress?.({ stage: 'summarizing-chunk', chunkIndex: 1, chunkCount: 1 });
    return runCompletion(engine, buildDirectSummaryPrompt(documentText.length), chunks[0]);
  }

  // Map step: summarize each chunk independently.
  const chunkSummaries: string[] = [];
  for (let i = 0; i < chunks.length; i++) {
    options.onProgress?.({ stage: 'summarizing-chunk', chunkIndex: i + 1, chunkCount: chunks.length });
    const chunkSummary = await runCompletion(engine, MAP_STEP_SYSTEM_PROMPT, chunks[i]);
    chunkSummaries.push(chunkSummary);
  }

  // Reduce step: combine chunk summaries into one final structured summary.
  options.onProgress?.({ stage: 'combining' });
  const combinedInput = chunkSummaries
    .map((summary, index) => `[Section ${index + 1}]\n${summary}`)
    .join('\n\n');

  return runCompletion(engine, buildReduceStepPrompt(documentText.length), combinedInput);
}
