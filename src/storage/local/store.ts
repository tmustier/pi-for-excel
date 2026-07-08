/**
 * Base class for all storage stores.
 * Each store defines its IndexedDB schema and provides domain-specific methods.
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import type { StorageBackend, StoreConfig } from "./types.js";

export abstract class Store {
  private backend: StorageBackend | null = null;

  /**
   * Returns the IndexedDB configuration for this store.
   * Defines store name, key path, and indices.
   */
  abstract getConfig(): StoreConfig;

  /**
   * Sets the storage backend. Called by AppStorage after backend creation.
   */
  setBackend(backend: StorageBackend): void {
    this.backend = backend;
  }

  /**
   * Gets the storage backend. Throws if backend not set.
   * Concrete stores must use this to access the backend.
   */
  protected getBackend(): StorageBackend {
    if (!this.backend) {
      throw new Error(`Backend not set on ${this.constructor.name}`);
    }
    return this.backend;
  }
}
