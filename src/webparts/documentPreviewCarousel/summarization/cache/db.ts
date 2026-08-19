/**
 * Thin IndexedDB wrapper for storing cached document summaries client-side.
 * Deliberately minimal - no ORM, just the handful of operations the cache
 * layer above (summaryCache.ts) actually needs.
 */

const DB_NAME = 'ctxpack-carousel-summaries';
const DB_VERSION = 1;
const STORE_NAME = 'summaries';

export interface ICachedSummaryRecord {
  /** Composite key: `${fileId}:${fileVersion}`. A new file version gets a
   *  different key, so stale summaries are naturally never served - they
   *  just become unreachable and get cleaned up by eviction over time. */
  key: string;
  summary: string;
  sizeBytes: number;
  lastAccessedAt: number;
}

// Reuse a single open connection across all operations rather than opening
// a new one per call - IndexedDB connections left open concurrently can
// block each other, and there's no reason to pay the open-connection cost
// repeatedly for what's effectively a long-lived, page-session cache.
let dbPromise: Promise<IDBDatabase> | undefined;

function openDatabase(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(STORE_NAME)) {
        const store = db.createObjectStore(STORE_NAME, { keyPath: 'key' });
        store.createIndex('lastAccessedAt', 'lastAccessedAt', { unique: false });
      }
    };

    request.onsuccess = () => resolve(request.result);
    request.onerror = () => {
      dbPromise = undefined;
      reject(request.error);
    };
  });

  return dbPromise;
}

/** Testing/advanced use only - closes and forgets the shared connection so
 *  a fresh one is opened on next use (e.g. after deleting the DB in tests). */
export async function resetDatabaseConnectionForTests(): Promise<void> {
  const promiseToClose = dbPromise;
  if (promiseToClose) {
    const db = await promiseToClose;
    db.close();
    if (dbPromise === promiseToClose) {
      dbPromise = undefined;
    }
  }
}

export async function getRecord(key: string): Promise<ICachedSummaryRecord | undefined> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const request = tx.objectStore(STORE_NAME).get(key);
    request.onsuccess = () => resolve(request.result as ICachedSummaryRecord | undefined);
    request.onerror = () => reject(request.error);
  });
}

export async function putRecord(record: ICachedSummaryRecord): Promise<void> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).put(record);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

export async function deleteRecord(key: string): Promise<void> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readwrite');
    tx.objectStore(STORE_NAME).delete(key);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

/** Returns all records ordered oldest-accessed-first, for LRU eviction. */
export async function getAllRecordsByLastAccessed(): Promise<ICachedSummaryRecord[]> {
  const db = await openDatabase();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(STORE_NAME, 'readonly');
    const index = tx.objectStore(STORE_NAME).index('lastAccessedAt');
    const request = index.getAll();
    request.onsuccess = () => resolve(request.result as ICachedSummaryRecord[]);
    request.onerror = () => reject(request.error);
  });
}
