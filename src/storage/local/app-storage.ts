/**
 * High-level storage API providing access to all storage operations.
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import type { CustomProvidersStore } from "./custom-providers-store.js";
import type { ProviderKeysStore } from "./provider-keys-store.js";
import type { SessionsStore } from "./sessions-store.js";
import type { SettingsStore } from "./settings-store.js";
import type { StorageBackend } from "./types.js";

export class AppStorage {
  readonly settings: SettingsStore;
  readonly providerKeys: ProviderKeysStore;
  readonly sessions: SessionsStore;
  readonly customProviders: CustomProvidersStore;
  readonly backend: StorageBackend;

  constructor(
    settings: SettingsStore,
    providerKeys: ProviderKeysStore,
    sessions: SessionsStore,
    customProviders: CustomProvidersStore,
    backend: StorageBackend,
  ) {
    this.settings = settings;
    this.providerKeys = providerKeys;
    this.sessions = sessions;
    this.customProviders = customProviders;
    this.backend = backend;
  }

  async getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return this.backend.getQuotaInfo();
  }

  async requestPersistence(): Promise<boolean> {
    return this.backend.requestPersistence();
  }
}

// Global instance management
let globalAppStorage: AppStorage | null = null;

/**
 * Get the global AppStorage instance.
 * Throws if not initialized.
 */
export function getAppStorage(): AppStorage {
  if (!globalAppStorage) {
    throw new Error("AppStorage not initialized. Call setAppStorage() first.");
  }
  return globalAppStorage;
}

/**
 * Set the global AppStorage instance.
 */
export function setAppStorage(storage: AppStorage): void {
  globalAppStorage = storage;
}
