/**
 * Client-side cache for generated document summaries, keyed to a specific
 * file + version so a re-summarize only happens when the document actually
 * changes. Per the agreed design: caching lives entirely in the user's
 * browser (IndexedDB), nothing shared or server-side, with automatic
 * eviction once the cache grows past a size cap - no manual cleanup needed.
 */
import { deleteRecord, getAllRecordsByLastAccessed, getRecord, putRecord } from './db';

const DEFAULT_MAX_CACHE_BYTES = 50 * 1024 * 1024; // 50MB - see design discussion; summaries
// are just text, so this comfortably holds hundreds of documents' worth.

function byteLength(text: string): number {
  // Summaries are plain text/Markdown; this is an accurate-enough size
  // estimate for eviction purposes without pulling in a UTF-8 byte-counting
  // dependency for what's ultimately just a soft cap.
  return new Blob([text]).size;
}

export function makeCacheKey(fileId: string, fileVersion: string): string {
  return `${fileId}:${fileVersion}`;
}

export async function getCachedSummary(fileId: string, fileVersion: string): Promise<string | undefined> {
  const key = makeCacheKey(fileId, fileVersion);
  const record = await getRecord(key);
  if (!record) return undefined;

  // Touch the access time so this entry looks "recently used" to the LRU
  // eviction pass, without waiting on the write before returning to the caller.
  putRecord({ ...record, lastAccessedAt: Date.now() }).catch(() => undefined);

  return record.summary;
}

export async function setCachedSummary(
  fileId: string,
  fileVersion: string,
  summary: string,
  options: { maxCacheBytes?: number } = {}
): Promise<void> {
  const key = makeCacheKey(fileId, fileVersion);
  const sizeBytes = byteLength(summary);

  await putRecord({ key, summary, sizeBytes, lastAccessedAt: Date.now() });
  await evictIfOverBudget(options.maxCacheBytes ?? DEFAULT_MAX_CACHE_BYTES);
}

async function evictIfOverBudget(maxCacheBytes: number): Promise<void> {
  const records = await getAllRecordsByLastAccessed();
  let totalBytes = records.reduce((sum, r) => sum + r.sizeBytes, 0);

  if (totalBytes <= maxCacheBytes) return;

  // getAllRecordsByLastAccessed returns oldest-first (per the index), so
  // evicting from the front of the array removes least-recently-used
  // entries first, exactly the LRU behavior described in the design.
  for (const record of records) {
    if (totalBytes <= maxCacheBytes) break;
    await deleteRecord(record.key);
    totalBytes -= record.sizeBytes;
  }
}
