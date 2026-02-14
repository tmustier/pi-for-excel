/**
 * Provider picker overlay.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  closeOverlayById,
  createOverlayCloseButton,
  createOverlayDialog,
} from "../../ui/overlay-dialog.js";
import { PROVIDER_PICKER_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";

export async function showProviderPicker(): Promise<void> {
  if (closeOverlayById(PROVIDER_PICKER_OVERLAY_ID)) {
    return;
  }

  const { ALL_PROVIDERS, buildProviderRow } = await import("../../ui/provider-login.js");
  const storage = getAppStorage();
  const configuredKeys = await storage.providerKeys.list();
  const configuredSet = new Set(configuredKeys);

  const dialog = createOverlayDialog({
    overlayId: PROVIDER_PICKER_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-provider-picker-card",
  });

  const closeOverlay = dialog.close;

  const header = document.createElement("div");
  header.className = "pi-overlay-header";

  const titleWrap = document.createElement("div");
  titleWrap.className = "pi-overlay-title-wrap";

  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Providers";

  const subtitle = document.createElement("p");
  subtitle.className = "pi-overlay-subtitle";
  subtitle.textContent = "Connect providers to use their models.";

  const closeButton = createOverlayCloseButton({
    onClose: closeOverlay,
    label: "Close providers",
  });

  titleWrap.append(title, subtitle);
  header.append(titleWrap, closeButton);

  const list = document.createElement("div");
  list.className = "pi-welcome-providers pi-provider-picker-list";

  dialog.card.append(header, list);

  const expandedRef: { current: HTMLElement | null } = { current: null };

  for (const provider of ALL_PROVIDERS) {
    const isActive = configuredSet.has(provider.id);
    const row = buildProviderRow(provider, {
      isActive,
      expandedRef,
      onConnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} connected`);
      },
      onDisconnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} disconnected`);
      },
    });
    list.appendChild(row);
  }

  dialog.mount();
}
