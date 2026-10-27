/**
 * IndexedDB implementation of StorageBackend.
 * Provides multi-store key-value storage with transactions and quota management.
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 *
 * Schema compatibility invariant: this backend must open the existing
 * "pi-for-excel" database (same store names, key paths, and indices) so
 * pre-migration user data restores unchanged.
 */

import type { IndexedDBConfig, StorageTransaction } from "./types.js";
import type { StorageBackend } from "./types.js";

function toIdbError(error: DOMException | null): Error {
  return error ?? new Error("IndexedDB request failed");
}

/** All stores in this app use string (or numeric) keys; other key shapes are skipped. */
function normalizeIdbKey(key: IDBValidKey): string[] {
  if (typeof key === "string") return [key];
  if (typeof key === "number") return [key.toString()];
  return [];
}

/** Fired on `window` when this pane's database open waits for another pane to close it. */
export const STORAGE_UPGRADE_BLOCKED_EVENT = "pi:storage-upgrade-blocked";
/** Fired on `window` after this pane closed its database for a newer version's upgrade. */
export const STORAGE_VERSION_CHANGED_EVENT = "pi:storage-version-changed";

function dispatchStorageEvent(type: string): void {
  if (typeof window !== "undefined") window.dispatchEvent(new Event(type));
}

export class IndexedDBStorageBackend implements StorageBackend {
  private dbPromise: Promise<IDBDatabase> | null = null;
  private readonly config: IndexedDBConfig;

  constructor(config: IndexedDBConfig) {
    this.config = config;
  }

  private async getDB(): Promise<IDBDatabase> {
    if (!this.dbPromise) {
      this.dbPromise = new Promise((resolve, reject) => {
        const request = indexedDB.open(this.config.dbName, this.config.version);
        request.onerror = () => reject(toIdbError(request.error));
        // An older pane still holds the database open; the open continues once it closes.
        request.onblocked = () => dispatchStorageEvent(STORAGE_UPGRADE_BLOCKED_EVENT);
        request.onsuccess = () => {
          const db = request.result;
          // A newer version wants to upgrade the database: step aside so its
          // upgrade is not blocked. This pane's storage stops working until reload.
          db.onversionchange = () => {
            db.close();
            dispatchStorageEvent(STORAGE_VERSION_CHANGED_EVENT);
          };
          resolve(db);
        };
        request.onupgradeneeded = () => {
          const db = request.result;
          // Create object stores from config
          for (const storeConfig of this.config.stores) {
            if (db.objectStoreNames.contains(storeConfig.name)) continue;

            const params: IDBObjectStoreParameters = {};
            if (storeConfig.keyPath !== undefined) params.keyPath = storeConfig.keyPath;
            if (storeConfig.autoIncrement !== undefined) {
              params.autoIncrement = storeConfig.autoIncrement;
            }
            const store = db.createObjectStore(storeConfig.name, params);

            for (const indexConfig of storeConfig.indices ?? []) {
              const indexParams: IDBIndexParameters = {};
              if (indexConfig.unique !== undefined) indexParams.unique = indexConfig.unique;
              store.createIndex(indexConfig.name, indexConfig.keyPath, indexParams);
            }
          }
        };
      });
    }
    return this.dbPromise;
  }

  private promisifyRequest<T>(request: IDBRequest<T>): Promise<T> {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(toIdbError(request.error));
    });
  }

  /** Put honoring in-line (keyPath) vs out-of-line keys. */
  private putIntoStore(store: IDBObjectStore, key: string, value: unknown): Promise<IDBValidKey> {
    if (store.keyPath) {
      return this.promisifyRequest(store.put(value));
    }
    return this.promisifyRequest(store.put(value, key));
  }

  async get<T = unknown>(storeName: string, key: string): Promise<T | null> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const result = await this.promisifyRequest<T | undefined>(
      store.get(key) as IDBRequest<T | undefined>,
    );
    return result ?? null;
  }

  async set<T = unknown>(storeName: string, key: string, value: T): Promise<void> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readwrite").objectStore(storeName);
    await this.putIntoStore(store, key, value);
  }

  async delete(storeName: string, key: string): Promise<void> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readwrite").objectStore(storeName);
    await this.promisifyRequest(store.delete(key));
  }

  async keys(storeName: string, prefix?: string): Promise<string[]> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    if (prefix) {
      // Use IDBKeyRange for efficient prefix filtering
      const range = IDBKeyRange.bound(prefix, `${prefix}\uffff`, false, false);
      const keys = await this.promisifyRequest(store.getAllKeys(range));
      return keys.flatMap(normalizeIdbKey);
    }
    const keys = await this.promisifyRequest(store.getAllKeys());
    return keys.flatMap(normalizeIdbKey);
  }

  async getAllFromIndex<T = unknown>(
    storeName: string,
    indexName: string,
    direction: "asc" | "desc" = "asc",
  ): Promise<T[]> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const index = store.index(indexName);
    return new Promise((resolve, reject) => {
      const results: T[] = [];
      const request = index.openCursor(null, direction === "desc" ? "prev" : "next");
      request.onsuccess = () => {
        const cursor = request.result;
        if (cursor) {
          results.push(cursor.value as T);
          cursor.continue();
        } else {
          resolve(results);
        }
      };
      request.onerror = () => reject(toIdbError(request.error));
    });
  }

  async clear(storeName: string): Promise<void> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readwrite").objectStore(storeName);
    await this.promisifyRequest(store.clear());
  }

  async has(storeName: string, key: string): Promise<boolean> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    const result = await this.promisifyRequest(store.getKey(key));
    return result !== undefined;
  }

  async transaction<T>(
    storeNames: string[],
    mode: "readonly" | "readwrite",
    operation: (tx: StorageTransaction) => Promise<T>,
  ): Promise<T> {
    return this.indexedTransaction(storeNames, mode, async (idbTx) => {
      const storageTx: StorageTransaction = {
        get: async <V = unknown>(storeName: string, key: string): Promise<V | null> => {
          const result = await this.promisifyRequest<V | undefined>(
            idbTx.objectStore(storeName).get(key) as IDBRequest<V | undefined>,
          );
          return result ?? null;
        },
        set: async <V = unknown>(storeName: string, key: string, value: V): Promise<void> => {
          await this.putIntoStore(idbTx.objectStore(storeName), key, value);
        },
        delete: async (storeName: string, key: string): Promise<void> => {
          await this.promisifyRequest(idbTx.objectStore(storeName).delete(key));
        },
      };
      return operation(storageTx);
    });
  }

  /** Native indexes/cursors for stores with bounded, indexed queries. */
  async indexedTransaction<T>(
    storeNames: string[],
    mode: IDBTransactionMode,
    operation: (tx: IDBTransaction) => Promise<T>,
  ): Promise<T> {
    const db = await this.getDB();
    const idbTx = db.transaction(storeNames, mode);
    // Settle only once the transaction commits, so a write that fails at
    // commit time (for example QuotaExceededError) rejects instead of
    // reporting success.
    const committed = new Promise<void>((resolve, reject) => {
      idbTx.oncomplete = () => resolve();
      idbTx.onabort = () => reject(toIdbError(idbTx.error));
      idbTx.onerror = () => reject(toIdbError(idbTx.error));
    });
    // Callers that fail the operation get that error, not the abort it causes.
    void committed.catch(() => {});
    let result: T;
    try {
      // A transaction can abort without delivering the pending request's
      // error (Chromium's oversized structured-clone IPC is one example).
      result = await Promise.race([
        operation(idbTx),
        committed.then(() => new Promise<never>(() => {})),
      ]);
    } catch (error) {
      try {
        idbTx.abort();
      } catch {
        // Already finished or aborted.
      }
      throw error;
    }
    await committed;
    return result;
  }

  async getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    if (navigator.storage?.estimate) {
      const estimate = await navigator.storage.estimate();
      const usage = estimate.usage ?? 0;
      const quota = estimate.quota ?? 0;
      return {
        usage,
        quota,
        percent: quota ? (usage / quota) * 100 : 0,
      };
    }
    return { usage: 0, quota: 0, percent: 0 };
  }

  async requestPersistence(): Promise<boolean> {
    if (navigator.storage?.persist) {
      return await navigator.storage.persist();
    }
    return false;
  }
}
