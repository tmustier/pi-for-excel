/**
 * Shared app storage initialization for taskpane + dialog.
 */

import {
  AppStorage,
  IndexedDBStorageBackend,
  ProviderKeysStore,
  CustomProvidersStore,
  SessionsStore,
  SettingsStore,
  setAppStorage,
} from "@mariozechner/pi-web-ui";

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
