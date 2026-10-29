import { t } from "../language/index.js";
import { showToast } from "./toast.js";
import { API_KEY_OVERLAY_ID } from "./overlay-ids.js";
import {
  closeOverlayById,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { ALL_PROVIDERS, buildProviderRow, type ProviderDef } from "./provider-login.js";

function resolveProviderDef(provider: string): ProviderDef {
  return (
    ALL_PROVIDERS.find((def) => def.id === provider) ?? { id: provider, brand: provider }
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
      title: t("apiKeyDialog.title", { label: def.brand }),
      subtitle: t("apiKeyDialog.subtitle"),
      onClose: dialog.close,
      closeLabel: t("dialog.close"),
    });

    const expandedRef: { current: HTMLElement | null } = { current: null };
    const row = buildProviderRow(def, {
      isActive: false,
      expandedRef,
      onConnected: (_row, _id, label) => {
        settle(true);
        dialog.close();
        showToast(t("welcome.toast.connected", { label }), 3200);
      },
    });

    dialog.card.append(header, row);
    dialog.mount();

    row.querySelector<HTMLButtonElement>(".pi-login-trigger")?.click();
  });
}
