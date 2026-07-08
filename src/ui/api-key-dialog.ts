/**
 * Provider connect dialog — first-party replacement for pi-web-ui's
 * `ApiKeyPromptDialog` (docs/ui-ownership.md step 6).
 *
 * Shown when a request needs credentials for a provider that has none.
 * Reuses the welcome-overlay provider row, so OAuth logins work here too
 * (upstream only supported raw API keys).
 *
 * Resolves `true` once the provider is connected, `false` if dismissed.
 */

import { t } from "../language/index.js";
import { API_KEY_OVERLAY_ID } from "./overlay-ids.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { ALL_PROVIDERS, buildProviderRow, type ProviderDef } from "./provider-login.js";

function resolveProviderDef(provider: string): ProviderDef {
  return (
    ALL_PROVIDERS.find((def) => def.id === provider) ?? { id: provider, label: provider }
  );
}

export function promptForProviderConnection(provider: string): Promise<boolean> {
  closeOverlayById(API_KEY_OVERLAY_ID);

  const def = resolveProviderDef(provider);

  return new Promise<boolean>((resolve) => {
    let settled = false;

    const settle = (connected: boolean): void => {
      if (settled) return;
      settled = true;
      resolve(connected);
    };

    const dialog = createOverlayDialog({
      overlayId: API_KEY_OVERLAY_ID,
      cardClassName: "pi-api-key-card",
      restoreFocusOnClose: true,
    });

    dialog.addCleanup(() => {
      settle(false);
    });

    const { header } = createOverlayHeader({
      title: t("apiKeyDialog.title", { label: def.label }),
      subtitle: t("apiKeyDialog.subtitle"),
      onClose: dialog.close,
      closeLabel: t("dialog.close"),
    });

    const expandedRef: { current: HTMLElement | null } = { current: null };
    const row = buildProviderRow(def, {
      isActive: false,
      expandedRef,
      onConnected: () => {
        settle(true);
        dialog.close();
      },
    });

    dialog.card.append(header, row);
    dialog.mount();

    // Single-provider dialog: open the connect detail immediately.
    row.querySelector<HTMLButtonElement>(".pi-login-trigger")?.click();
  });
}
