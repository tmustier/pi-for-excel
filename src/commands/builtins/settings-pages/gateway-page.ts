/**
 * Custom gateway page — OpenAI-compatible corporate/local gateways.
 */

import { t } from "../../../language/index.js";
import { requestConfirmationDialog } from "../../../ui/confirm-dialog.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import { buildCustomGatewaySection } from "../custom-gateway-settings.js";
import { getSettingsPagesDependencies } from "./dependencies.js";

export function createGatewayPage(): SettingsShellPage {
  return {
    id: "gateway",
    parentId: "root",
    title: () => t("custom-gateway.title"),
    subtitle: () => t("custom-gateway.hint"),
    render: async (ctx) => {
      const models = getSettingsPagesDependencies().models;
      const { section, hasUnsavedChanges } = await buildCustomGatewaySection({
        ...(models !== undefined ? { models } : {}),
        onProvidersChanged: () => {
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        },
        includeHeading: false,
      });

      // A half-filled gateway is not dropped silently on Back or close.
      ctx.setBeforeLeave(async () => {
        if (!hasUnsavedChanges()) return true;
        return requestConfirmationDialog({
          title: t("custom-gateway.discard.title"),
          message: t("custom-gateway.discard.message"),
          confirmLabel: t("custom-gateway.discard.confirm"),
          cancelLabel: t("custom-gateway.discard.keep"),
          confirmButtonTone: "danger",
          restoreFocusOnClose: false,
        });
      });

      ctx.body.appendChild(section);
    },
  };
}
