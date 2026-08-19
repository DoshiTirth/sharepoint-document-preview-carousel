import { chunkText } from '../chunker';

describe('chunkText', () => {
  it('returns a single chunk for short text', () => {
    const result = chunkText('A short paragraph of text.', { maxTokensPerChunk: 100 });
    expect(result).toEqual(['A short paragraph of text.']);
  });

  it('returns an empty array for empty/whitespace-only text', () => {
    expect(chunkText('', { maxTokensPerChunk: 100 })).toEqual([]);
    expect(chunkText('   \n\n  ', { maxTokensPerChunk: 100 })).toEqual([]);
  });

  it('splits long text into multiple chunks, none exceeding the token budget', () => {
    const paragraph = 'This is a reasonably long sentence used to pad out the paragraph. '.repeat(20);
    const longText = Array.from({ length: 10 }, (_, i) => `Section ${i}: ${paragraph}`).join('\n\n');

    const maxTokensPerChunk = 50;
    const chunks = chunkText(longText, { maxTokensPerChunk });

    expect(chunks.length).toBeGreaterThan(1);
    for (const chunk of chunks) {
      // Same estimate formula as the implementation (chars / 3.5), with a
      // small tolerance since sentence-splitting doesn't hit the boundary exactly.
      expect(chunk.length / 3.5).toBeLessThanOrEqual(maxTokensPerChunk * 1.2);
    }
  });

  it('preserves all non-whitespace content across chunks (nothing silently dropped)', () => {
    const longText = Array.from({ length: 5 }, (_, i) => `Paragraph number ${i} with some content here.`).join(
      '\n\n'
    );
    const chunks = chunkText(longText, { maxTokensPerChunk: 15 });

    for (let i = 0; i < 5; i++) {
      expect(chunks.some((c: string) => c.includes(`Paragraph number ${i}`))).toBe(true);
    }
  });

  it('breaks up a single oversized paragraph by sentence rather than failing', () => {
    const hugeParagraph =
      'First sentence here. Second sentence follows. Third sentence continues. Fourth sentence too. Fifth and final sentence.';
    const chunks = chunkText(hugeParagraph, { maxTokensPerChunk: 10 });

    expect(chunks.length).toBeGreaterThan(1);
    expect(chunks.join(' ')).toContain('First sentence');
    expect(chunks.join(' ')).toContain('Fifth and final sentence');
  });
});
