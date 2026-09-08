import assert from "node:assert/strict";
import { test } from "node:test";

import { saveOpenAiGatewayConfig } from "../src/auth/custom-gateways.ts";
import { buildCustomGatewaySection } from "../src/commands/builtins/custom-gateway-settings.ts";
import { AppStorage, setAppStorage } from "../src/storage/local/app-storage.ts";
import { CustomProvidersStore, type CustomProvider } from "../src/storage/local/custom-providers-store.ts";
import { IndexedDBStorageBackend } from "../src/storage/local/indexeddb-storage-backend.ts";
import { ModelCatalogsStore } from "../src/storage/local/model-catalogs-store.ts";
import { ProviderKeysStore } from "../src/storage/local/provider-keys-store.ts";
import { SessionsStore } from "../src/storage/local/sessions-store.ts";
import { SettingsStore } from "../src/storage/local/settings-store.ts";
import { CONFIRM_DIALOG_OVERLAY_ID } from "../src/ui/overlay-ids.ts";
import { installFakeDom } from "./fake-dom.test.ts";

class MemoryCustomProvidersStore extends CustomProvidersStore {
  private readonly providers = new Map<string, CustomProvider>();

  override get(id: string): Promise<CustomProvider | null> {
    return Promise.resolve(this.providers.get(id) ?? null);
  }

  override set(provider: CustomProvider): Promise<void> {
    this.providers.set(provider.id, provider);
    return Promise.resolve();
  }

  override delete(id: string): Promise<void> {
    this.providers.delete(id);
    return Promise.resolve();
  }

  override getAll(): Promise<CustomProvider[]> {
    return Promise.resolve(Array.from(this.providers.values()));
  }
}

function findButtonByText(root: HTMLElement, text: string): HTMLElement | null {
  for (const button of root.querySelectorAll("button")) {
    if (button instanceof HTMLElement && button.textContent === text) {
      return button;
    }
  }
  return null;
}

function installTestStorage(customProviders: CustomProvidersStore): void {
  const backend = new IndexedDBStorageBackend({
    dbName: "custom-gateway-settings-test",
    version: 1,
    stores: [],
  });
  setAppStorage(new AppStorage(
    new SettingsStore(),
    new ProviderKeysStore(),
    new SessionsStore(),
    customProviders,
    new ModelCatalogsStore(),
    backend,
  ));
}

void test("custom gateway deletion persists only after overlay confirmation", async () => {
  const { document, restore } = installFakeDom();
  const customProviders = new MemoryCustomProvidersStore();
  installTestStorage(customProviders);

  try {
    await saveOpenAiGatewayConfig(customProviders, {
      displayName: "Test gateway",
      endpointUrl: "https://gateway.example.com/v1",
      modelId: "test-model",
      apiKey: "test-key",
    });

    let providersChangedCount = 0;
    const section = await buildCustomGatewaySection({
      onProvidersChanged: () => {
        providersChangedCount += 1;
      },
    });

    const deleteButton = findButtonByText(section, "Delete");
    assert.ok(deleteButton);
    deleteButton.dispatchEvent(new Event("click"));

    const firstDialog = document.getElementById(CONFIRM_DIALOG_OVERLAY_ID);
    assert.ok(firstDialog);
    const cancelButton = findButtonByText(firstDialog, "Cancel");
    assert.ok(cancelButton);
    cancelButton.dispatchEvent(new Event("click"));
    await new Promise<void>((resolve) => setTimeout(resolve, 0));

    assert.equal((await customProviders.getAll()).length, 1);
    assert.equal(providersChangedCount, 0);

    deleteButton.dispatchEvent(new Event("click"));
    const secondDialog = document.getElementById(CONFIRM_DIALOG_OVERLAY_ID);
    assert.ok(secondDialog);
    const confirmButton = findButtonByText(secondDialog, "Delete");
    assert.ok(confirmButton);
    confirmButton.dispatchEvent(new Event("click"));
    await new Promise<void>((resolve) => setTimeout(resolve, 0));

    assert.equal((await customProviders.getAll()).length, 0);
    assert.equal(providersChangedCount, 1);
  } finally {
    restore();
  }
});
