/**
 * Custom gateway page — OpenAI-compatible corporate/local gateways.
 */

import { t } from "../../../language/index.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import { buildCustomGatewaySection } from "../custom-gateway-settings.js";

export function createGatewayPage(): SettingsShellPage {
  return {
    id: "gateway",
    parentId: "root",
    title: () => t("custom-gateway.title"),
    subtitle: () => t("custom-gateway.hint"),
    render: async (ctx) => {
      const section = await buildCustomGatewaySection({
        onProvidersChanged: () => {
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        },
        includeHeading: false,
      });

      ctx.body.appendChild(section);
    },
  };
}
