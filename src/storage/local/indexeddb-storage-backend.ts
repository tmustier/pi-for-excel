/**
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono).
 */

import type { IndexedDBConfig, StorageTransaction } from "./types.js";
import type { StorageBackend } from "./types.js";

function toIdbError(error: DOMException | null): Error {
  return error ?? new Error("IndexedDB request failed");
}

function normalizeIdbKey(key: IDBValidKey): string[] {
  if (typeof key === "string") return [key];
  if (typeof key === "number") return [key.toString()];
  return [];
}

export const STORAGE_UPGRADE_BLOCKED_EVENT = "pi:storage-upgrade-blocked";
export const STORAGE_VERSION_CHANGED_EVENT = "pi:storage-version-changed";

function dispatchStorageEvent(type: string): void {
  if (typeof window !== "undefined") window.dispatchEvent(new Event(type));
}

export class IndexedDBStorageBackend implements StorageBackend {
  private state:
    | { kind: "unopened" }
    | { kind: "opening"; promise: Promise<IDBDatabase> }
    | { kind: "open"; db: IDBDatabase }
    | { kind: "closed-for-upgrade" } = { kind: "unopened" };
  private readonly config: IndexedDBConfig;

  constructor(config: IndexedDBConfig) {
    this.config = config;
  }

  private async getDB(): Promise<IDBDatabase> {
    if (this.state.kind === "closed-for-upgrade") throw new Error("Storage changed in another pane; reload required.");
    if (this.state.kind === "open") return this.state.db;
    if (this.state.kind === "opening") return this.state.promise;

    const promise = new Promise<IDBDatabase>((resolve, reject) => {
      const request = indexedDB.open(this.config.dbName, this.config.version);
      request.onerror = () => reject(toIdbError(request.error));
      request.onblocked = () => dispatchStorageEvent(STORAGE_UPGRADE_BLOCKED_EVENT);
      request.onsuccess = () => {
        const db = request.result;
        db.onversionchange = () => {
          db.close();
          this.state = { kind: "closed-for-upgrade" };
          dispatchStorageEvent(STORAGE_VERSION_CHANGED_EVENT);
        };
        resolve(db);
      };
      request.onupgradeneeded = () => {
        const db = request.result;
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
    }).then((db) => {
      this.state = { kind: "open", db };
      return db;
    }, (error) => {
      this.state = { kind: "unopened" };
      throw error;
    });
    this.state = { kind: "opening", promise };
    return promise;
  }

  private promisifyRequest<T>(request: IDBRequest<T>): Promise<T> {
    return new Promise((resolve, reject) => {
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(toIdbError(request.error));
    });
  }

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
    await this.indexedTransaction([storeName], "readwrite", async (tx) => {
      await this.putIntoStore(tx.objectStore(storeName), key, value);
    });
  }

  async delete(storeName: string, key: string): Promise<void> {
    await this.indexedTransaction([storeName], "readwrite", async (tx) => {
      await this.promisifyRequest(tx.objectStore(storeName).delete(key));
    });
  }

  async keys(storeName: string, prefix?: string): Promise<string[]> {
    const db = await this.getDB();
    const store = db.transaction(storeName, "readonly").objectStore(storeName);
    if (prefix) {
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
    await this.indexedTransaction([storeName], "readwrite", async (tx) => {
      await this.promisifyRequest(tx.objectStore(storeName).clear());
    });
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

  async indexedTransaction<T>(
    storeNames: string[],
    mode: IDBTransactionMode,
    operation: (tx: IDBTransaction) => Promise<T>,
  ): Promise<T> {
    const db = await this.getDB();
    const idbTx = db.transaction(storeNames, mode);
    const committed = new Promise<void>((resolve, reject) => {
      idbTx.oncomplete = () => resolve();
      idbTx.onabort = () => reject(toIdbError(idbTx.error));
      idbTx.onerror = () => reject(toIdbError(idbTx.error));
    });
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
