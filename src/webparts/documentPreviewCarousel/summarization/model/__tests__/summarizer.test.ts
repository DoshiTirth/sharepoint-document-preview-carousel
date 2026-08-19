/**
 * The WebLLM engine is mocked here rather than run for real - it needs an
 * actual WebGPU device and a multi-GB model download, neither available in
 * a unit test environment. What we verify is OUR orchestration logic: that
 * short text skips the map step, long text gets chunked and each chunk
 * summarized before a final reduce pass, and progress callbacks fire with
 * the right stages. Whether the model's actual output is a good summary is
 * a model-quality question to validate manually, not something a unit test
 * can meaningfully assert on.
 */
import type { MLCEngine } from '@mlc-ai/web-llm';
import { summarizeDocumentText, SummarizationProgressHandler } from '../summarizer';

function makeMockEngine(replyFn: (userContent: string) => string): {
  engine: MLCEngine;
  calls: { system: string; user: string }[];
} {
  const calls: { system: string; user: string }[] = [];

  const engine = {
    chat: {
      completions: {
        create: jest.fn((request: { messages: { role: string; content: string }[] }) => {
          const system = request.messages.find((m) => m.role === 'system')?.content || '';
          const user = request.messages.find((m) => m.role === 'user')?.content || '';
          calls.push({ system, user });
          return Promise.resolve({
            choices: [{ message: { content: replyFn(user) } }],
          });
        }),
      },
    },
  } as unknown as MLCEngine;

  return { engine, calls };
}

describe('summarizeDocumentText', () => {
  it('short text: skips the map step and makes exactly one model call', async () => {
    const { engine, calls } = makeMockEngine(() => '## Summary\n\nA short document about testing.');

    const result = await summarizeDocumentText('A short bit of document text.', {
      engineProvider: () => Promise.resolve(engine),
    });

    expect(calls).toHaveLength(1);
    expect(result).toContain('A short document about testing.');
  });

  it('long text: chunks, summarizes each chunk (map), then combines (reduce)', async () => {
    const longText = Array.from(
      { length: 20 },
      (_, i) => `Section ${i}: ${'Padding content to make this chunk substantial. '.repeat(100)}`
    ).join('\n\n');

    let callCount = 0;
    const { engine, calls } = makeMockEngine(() => {
      callCount += 1;
      return `Chunk summary #${callCount}`;
    });

    const result = await summarizeDocumentText(longText, {
      engineProvider: () => Promise.resolve(engine),
    });

    // More than one call means the map step ran per chunk, plus one final reduce call.
    expect(calls.length).toBeGreaterThan(1);
    // The reduce call's input should reference the map step's chunk summaries.
    const reduceCall = calls[calls.length - 1];
    expect(reduceCall.user).toContain('Chunk summary #1');
    expect(result).toBe(`Chunk summary #${callCount}`); // last call's output is the final result
  });

  it('reports progress through loading, per-chunk, and combining stages', async () => {
    const longText = Array.from(
      { length: 20 },
      (_, i) => `Section ${i}: ${'Padding content to make this chunk substantial. '.repeat(100)}`
    ).join('\n\n');

    const { engine } = makeMockEngine(() => 'summary');
    const stages: string[] = [];
    const onProgress: SummarizationProgressHandler = (status: { stage: string }) => stages.push(status.stage);

    await summarizeDocumentText(longText, {
      engineProvider: () => Promise.resolve(engine),
      onProgress,
    });

    expect(stages[0]).toBe('loading-model');
    expect(stages).toContain('summarizing-chunk');
    expect(stages).toContain('combining');
  });

  it('handles empty/unreadable text without crashing', async () => {
    const { engine, calls } = makeMockEngine(() => 'unused');

    const result = await summarizeDocumentText('   ', {
      engineProvider: () => Promise.resolve(engine),
    });

    expect(calls).toHaveLength(0);
    expect(result).toContain('doesn\u2019t contain enough readable text');
  });
});
