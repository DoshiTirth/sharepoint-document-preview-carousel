/**
 * Uses fake-indexeddb (a spec-compliant in-memory IndexedDB implementation)
 * rather than mocking our own db.ts. jsdom doesn't implement IndexedDB at
 * all - another gap in this jsdom version, consistent with the Blob gap
 * found in PR #2 - so without this, there'd be nothing to test against.
 * Unlike the mocking approach used for third-party model/parsing libraries
 * elsewhere in this project, here we WANT real IndexedDB transaction/index
 * semantics exercised, since the eviction logic's correctness depends on them.
 */
import 'fake-indexeddb/auto';
import { getCachedSummary, setCachedSummary } from '../summaryCache';
import { resetDatabaseConnectionForTests } from '../db';

describe('summaryCache', () => {
  beforeEach(async () => {
    // fake-indexeddb persists across tests within a module unless reset;
    // close any open connection first (a lingering connection blocks
    // deleteDatabase), then delete the whole database for isolation.
    await resetDatabaseConnectionForTests();
    await new Promise<void>((resolve) => {
      const req = indexedDB.deleteDatabase('ctxpack-carousel-summaries');
      req.onsuccess = () => resolve();
      req.onerror = () => resolve();
      req.onblocked = () => resolve();
    });
  });

  it('returns undefined for a key that was never cached', async () => {
    const result = await getCachedSummary('file-1', 'v1');
    expect(result).toBeUndefined();
  });

  it('stores and retrieves a summary', async () => {
    await setCachedSummary('file-1', 'v1', '## Summary\n\nA test summary.');
    const result = await getCachedSummary('file-1', 'v1');
    expect(result).toBe('## Summary\n\nA test summary.');
  });

  it('treats a different file version as a cache miss (no stale summaries served)', async () => {
    await setCachedSummary('file-1', 'v1', 'Summary for v1');
    const result = await getCachedSummary('file-1', 'v2');
    expect(result).toBeUndefined();
  });

  it('keeps different files with the same version independent', async () => {
    await setCachedSummary('file-1', 'v1', 'Summary for file 1');
    await setCachedSummary('file-2', 'v1', 'Summary for file 2');

    expect(await getCachedSummary('file-1', 'v1')).toBe('Summary for file 1');
    expect(await getCachedSummary('file-2', 'v1')).toBe('Summary for file 2');
  });

  it('evicts least-recently-used entries once the size budget is exceeded', async () => {
    const bigSummary = 'x'.repeat(1000);
    const tinyBudget = 2500; // fits roughly 2 of these entries, not 3

    await setCachedSummary('file-1', 'v1', bigSummary, { maxCacheBytes: tinyBudget });
    await setCachedSummary('file-2', 'v1', bigSummary, { maxCacheBytes: tinyBudget });
    // Access file-1 again so it's more recently used than file-2 at this point.
    await getCachedSummary('file-1', 'v1');
    await setCachedSummary('file-3', 'v1', bigSummary, { maxCacheBytes: tinyBudget });

    // file-2 was least-recently-used at the time the budget was exceeded, so
    // it should have been evicted; file-1 (recently touched) and file-3
    // (just written) should survive.
    expect(await getCachedSummary('file-2', 'v1')).toBeUndefined();
    expect(await getCachedSummary('file-1', 'v1')).toBe(bigSummary);
    expect(await getCachedSummary('file-3', 'v1')).toBe(bigSummary);
  });

  it('does not evict anything while under budget', async () => {
    await setCachedSummary('file-1', 'v1', 'small summary', { maxCacheBytes: 50 * 1024 * 1024 });
    await setCachedSummary('file-2', 'v1', 'another small summary', { maxCacheBytes: 50 * 1024 * 1024 });

    expect(await getCachedSummary('file-1', 'v1')).toBe('small summary');
    expect(await getCachedSummary('file-2', 'v1')).toBe('another small summary');
  });
});

