/**
 * Shared app storage initialization for taskpane + dialog.
 */

import { AppStorage, setAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";
import { IndexedDBStorageBackend } from "@earendil-works/pi-web-ui/dist/storage/backends/indexeddb-storage-backend.js";
import { CustomProvidersStore } from "@earendil-works/pi-web-ui/dist/storage/stores/custom-providers-store.js";
import { ProviderKeysStore } from "@earendil-works/pi-web-ui/dist/storage/stores/provider-keys-store.js";
import { SessionsStore } from "@earendil-works/pi-web-ui/dist/storage/stores/sessions-store.js";
import { SettingsStore } from "@earendil-works/pi-web-ui/dist/storage/stores/settings-store.js";

export function initAppStorage(dbName = "pi-for-excel") {
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
