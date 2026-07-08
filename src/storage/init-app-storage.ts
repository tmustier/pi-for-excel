/**
 * Shared app storage initialization for taskpane + dialog.
 */

import { AppStorage, setAppStorage } from "./local/app-storage.js";
import { IndexedDBStorageBackend } from "./local/indexeddb-storage-backend.js";
import { CustomProvidersStore } from "./local/custom-providers-store.js";
import { ProviderKeysStore } from "./local/provider-keys-store.js";
import { SessionsStore } from "./local/sessions-store.js";
import { SettingsStore } from "./local/settings-store.js";

type InitializedAppStorage = {
  storage: AppStorage;
  settings: SettingsStore;
  providerKeys: ProviderKeysStore;
  sessions: SessionsStore;
  customProviders: CustomProvidersStore;
  backend: IndexedDBStorageBackend;
};

export function initAppStorage(dbName = "pi-for-excel"): InitializedAppStorage {
  const settings = new SettingsStore();
  const providerKeys = new ProviderKeysStore();
  const sessions = new SessionsStore();
  const customProviders = new CustomProvidersStore();

  const backend = new IndexedDBStorageBackend({
    dbName,
    version: 1,
    stores: [
      settings.getConfig(),
      providerKeys.getConfig(),
      sessions.getConfig(),
      SessionsStore.getMetadataConfig(),
      customProviders.getConfig(),
    ],
  });

  settings.setBackend(backend);
  providerKeys.setBackend(backend);
  sessions.setBackend(backend);
  customProviders.setBackend(backend);

  const storage = new AppStorage(settings, providerKeys, sessions, customProviders, backend);
  setAppStorage(storage);

  return { storage, settings, providerKeys, sessions, customProviders, backend };
}
