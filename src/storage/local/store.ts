/**
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono).
 */

import type { StorageBackend, StoreConfig } from "./types.js";

export abstract class Store {
  private backend: StorageBackend | null = null;

  abstract getConfig(): StoreConfig;

  setBackend(backend: StorageBackend): void {
    this.backend = backend;
  }

  protected getBackend(): StorageBackend {
    if (!this.backend) {
      throw new Error(`Backend not set on ${this.constructor.name}`);
    }
    return this.backend;
  }
}
