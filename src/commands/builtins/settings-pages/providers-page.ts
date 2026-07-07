/**
 * Model providers page — connect/disconnect provider logins.
 */

import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import { t } from "../../../language/index.js";
import { VISIBLE_PROVIDERS, buildProviderRow } from "../../../ui/provider-login.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import { showToast } from "../../../ui/toast.js";

export function createProvidersPage(): SettingsShellPage {
  return {
    id: "providers",
    parentId: "root",
    title: () => t("settings.page.providers"),
    subtitle: () => t("settings.section.providers.hint"),
    render: async (ctx) => {
      const providerList = document.createElement("div");
      providerList.className = "pi-welcome-providers pi-provider-picker-list pi-settings-provider-list";

      const storage = getAppStorage();

      let configuredSet = new Set<string>();
      try {
        const configuredKeys = await storage.providerKeys.list();
        configuredSet = new Set(configuredKeys);
      } catch {
        const warning = document.createElement("p");
        warning.className = "pi-overlay-hint pi-overlay-text-warning";
        warning.textContent = t("settings.warning.provider_state");
        ctx.body.appendChild(warning);
      }

      const expandedRef: { current: HTMLElement | null } = { current: null };

      for (const provider of VISIBLE_PROVIDERS) {
        const row = buildProviderRow(provider, {
          isActive: configuredSet.has(provider.id),
          expandedRef,
          onConnected: (_row: HTMLElement, _id: string, label: string) => {
            document.dispatchEvent(new CustomEvent("pi:providers-changed"));
            showToast(t("settings.toast.connected", { label }));
          },
          onDisconnected: (_row: HTMLElement, _id: string, label: string) => {
            document.dispatchEvent(new CustomEvent("pi:providers-changed"));
            showToast(t("settings.toast.disconnected", { label }));
          },
        });

        providerList.appendChild(row);
      }

      ctx.body.appendChild(providerList);
    },
  };
}
