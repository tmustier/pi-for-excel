/**
 * Welcome/login overlay shown when no providers are configured.
 */

import type { ProviderKeysStore } from "@mariozechner/pi-web-ui";

import { showToast } from "../ui/toast.js";
import { setActiveProviders } from "./model-selector-patch.js";

export async function showWelcomeLogin(providerKeys: ProviderKeysStore): Promise<void> {
  const { ALL_PROVIDERS, buildProviderRow } = await import("../ui/provider-login.js");

  return new Promise<void>((resolve) => {
    const overlay = document.createElement("div");
    overlay.className = "pi-welcome-overlay";
    overlay.innerHTML = `
      <div class="pi-welcome-card" style="text-align: left;">
        <div class="pi-welcome-logo" style="text-align: center;">π</div>
        <h2 class="pi-welcome-title" style="text-align: center;">Pi for Excel</h2>
        <p class="pi-welcome-subtitle" style="text-align: center;">Connect a provider to get started</p>
        <div class="pi-welcome-providers"></div>
      </div>
    `;

    const providerList = overlay.querySelector<HTMLDivElement>(".pi-welcome-providers");
    if (!providerList) {
      throw new Error("Welcome provider list not found");
    }

    const expandedRef = { current: null as HTMLElement | null };

    for (const provider of ALL_PROVIDERS) {
      const row = buildProviderRow(provider, {
        isActive: false,
        expandedRef,
        onConnected: async (_row, _id, label) => {
          const updated = await providerKeys.list();
          setActiveProviders(new Set(updated));
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
          showToast(`${label} connected`);
          overlay.remove();
          resolve();
        },
      });
      providerList.appendChild(row);
    }

    document.body.appendChild(overlay);
  });
}
