/**
 * Welcome/login overlay shown when no providers are configured.
 */

import type { ProviderKeysStore } from "@mariozechner/pi-web-ui/dist/storage/stores/provider-keys-store.js";
import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import { closeOverlayById, createOverlayDialog } from "../ui/overlay-dialog.js";
import { WELCOME_LOGIN_OVERLAY_ID } from "../ui/overlay-ids.js";
import { showToast } from "../ui/toast.js";
import { setActiveProviders } from "../compat/model-selector-patch.js";
import {
  DEFAULT_LOCAL_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
} from "../auth/proxy-validation.js";

function createElement<K extends keyof HTMLElementTagNameMap>(
  tagName: K,
  className?: string,
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tagName);
  if (className) {
    element.className = className;
  }
  return element;
}

async function testLocalHttpsProxy(proxyUrl: string): Promise<boolean> {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), 1200);

  try {
    const url = `${proxyUrl.replace(/\/+$/, "")}/?url=${encodeURIComponent("https://example.com")}`;
    const resp = await fetch(url, { signal: controller.signal });
    return resp.ok;
  } catch {
    return false;
  } finally {
    clearTimeout(timeout);
  }
}

export async function showWelcomeLogin(providerKeys: ProviderKeysStore): Promise<void> {
  const { ALL_PROVIDERS, buildProviderRow } = await import("../ui/provider-login.js");

  // Make OAuth flows usable even before the user can access /settings.
  try {
    const storage = getAppStorage();
    const enabled = await storage.settings.get("proxy.enabled");
    const url = await storage.settings.get("proxy.url");

    const currentUrl = typeof url === "string" && url.trim().length > 0 ? url.trim() : DEFAULT_LOCAL_PROXY_URL;

    if (url === null) {
      await storage.settings.set("proxy.url", currentUrl);
    }

    // Auto-enable if a local HTTPS proxy is actually reachable.
    if (!enabled) {
      const ok = await testLocalHttpsProxy(currentUrl);
      if (ok) {
        await storage.settings.set("proxy.enabled", true);
      }
    }
  } catch {
    // ignore — welcome overlay should still show
  }

  closeOverlayById(WELCOME_LOGIN_OVERLAY_ID);

  return new Promise<void>((resolve) => {
    const dialog = createOverlayDialog({
      overlayId: WELCOME_LOGIN_OVERLAY_ID,
      cardClassName: "pi-welcome-card",
    });

    let settled = false;
    dialog.addCleanup(() => {
      if (settled) {
        return;
      }

      settled = true;
      resolve();
    });

    const closeOverlay = dialog.close;

    const titleId = `${WELCOME_LOGIN_OVERLAY_ID}-title`;
    const subtitleId = `${WELCOME_LOGIN_OVERLAY_ID}-subtitle`;

    const logo = createElement("div", "pi-welcome-logo");
    logo.textContent = "π";

    const title = createElement("h2", "pi-welcome-title");
    title.id = titleId;
    title.textContent = "Pi for Excel";

    const subtitle = createElement("p", "pi-welcome-subtitle");
    subtitle.id = subtitleId;
    subtitle.textContent = "Connect an AI provider to get started";

    const intro = createElement("p", "pi-welcome-intro");
    intro.textContent = "Pi reads your workbook, explains formulas, and makes edits directly in Excel.";

    const providerSectionTitle = createElement("p", "pi-welcome-section-title");
    providerSectionTitle.textContent = "Choose a provider";

    const providerList = createElement("div", "pi-welcome-providers");

    const proxyToggle = createElement("button", "pi-welcome-proxy-toggle");
    proxyToggle.type = "button";
    proxyToggle.textContent = "Having login trouble? Configure local proxy";
    proxyToggle.setAttribute("aria-expanded", "false");

    const proxyPanel = createElement("section", "pi-welcome-proxy");
    proxyPanel.hidden = true;

    const proxyTopRow = createElement("div", "pi-welcome-proxy__row");

    const proxyTitle = createElement("div", "pi-welcome-proxy__title");
    proxyTitle.textContent = "Local HTTPS proxy";

    const proxyToggleLabel = createElement("label", "pi-welcome-proxy__toggle");
    const proxyEnabledEl = createElement("input", "pi-welcome-proxy__enabled");
    proxyEnabledEl.type = "checkbox";
    const proxyToggleText = createElement("span");
    proxyToggleText.textContent = "Enabled";
    proxyToggleLabel.append(proxyEnabledEl, proxyToggleText);

    proxyTopRow.append(proxyTitle, proxyToggleLabel);

    const proxyUrlRow = createElement("div", "pi-welcome-proxy__row pi-welcome-proxy__row--compact");
    const proxyUrlEl = createElement("input", "pi-welcome-proxy__url");
    proxyUrlEl.type = "text";
    proxyUrlEl.spellcheck = false;

    const proxySaveEl = createElement("button", "pi-welcome-proxy__save");
    proxySaveEl.type = "button";
    proxySaveEl.textContent = "Save";

    proxyUrlRow.append(proxyUrlEl, proxySaveEl);

    const proxyHint = createElement("p", "pi-welcome-proxy__hint");
    const proxyCode = createElement("code");
    proxyCode.textContent = DEFAULT_LOCAL_PROXY_URL;

    const proxyGuideLink = createElement("a");
    proxyGuideLink.href = PROXY_HELPER_DOCS_URL;
    proxyGuideLink.target = "_blank";
    proxyGuideLink.rel = "noopener noreferrer";
    proxyGuideLink.textContent = "Step-by-step guide";

    proxyHint.append(
      "Needed only when OAuth login is blocked by CORS. Keep this URL at ",
      proxyCode,
      ", run a local HTTPS proxy helper, then enable this toggle. ",
      proxyGuideLink,
      ".",
    );

    proxyPanel.append(proxyTopRow, proxyUrlRow, proxyHint);

    dialog.card.replaceChildren(
      logo,
      title,
      subtitle,
      intro,
      providerSectionTitle,
      providerList,
      proxyToggle,
      proxyPanel,
    );

    dialog.overlay.setAttribute("aria-labelledby", titleId);
    dialog.overlay.setAttribute("aria-describedby", subtitleId);

    proxyToggle.addEventListener("click", () => {
      const willOpen = proxyPanel.hidden;
      proxyPanel.hidden = !willOpen;
      proxyToggle.setAttribute("aria-expanded", willOpen ? "true" : "false");
      proxyToggle.textContent = willOpen
        ? "Hide local proxy settings"
        : "Having login trouble? Configure local proxy";

      if (willOpen) {
        proxyUrlEl.focus();
      }
    });

    const hydrateProxyUi = async () => {
      try {
        const storage = getAppStorage();
        const enabled = await storage.settings.get("proxy.enabled");
        const url = await storage.settings.get("proxy.url");
        proxyEnabledEl.checked = Boolean(enabled);
        proxyUrlEl.value = typeof url === "string" && url.trim().length > 0 ? url.trim() : DEFAULT_LOCAL_PROXY_URL;
      } catch {
        proxyEnabledEl.checked = false;
        proxyUrlEl.value = DEFAULT_LOCAL_PROXY_URL;
      }
    };

    const saveProxyUi = async () => {
      try {
        const storage = getAppStorage();
        await storage.settings.set("proxy.enabled", proxyEnabledEl.checked);
        await storage.settings.set("proxy.url", proxyUrlEl.value.trim());
        showToast("Proxy settings saved");
      } catch {
        showToast("Failed to save proxy settings");
      }
    };

    proxyEnabledEl.addEventListener("change", () => {
      void saveProxyUi();
    });
    proxySaveEl.addEventListener("click", () => {
      void saveProxyUi();
    });

    void hydrateProxyUi();

    const expandedRef: { current: HTMLElement | null } = { current: null };

    for (const provider of ALL_PROVIDERS) {
      const row = buildProviderRow(provider, {
        isActive: false,
        expandedRef,
        onConnected: (_row, _id, label) => {
          void (async () => {
            const updated = await providerKeys.list();
            setActiveProviders(new Set(updated));
            document.dispatchEvent(new CustomEvent("pi:providers-changed"));
            showToast(`${label} connected — try “Explain this workbook”.`, 3200);
            closeOverlay();
          })();
        },
        onDisconnected: (_row, _id, label) => {
          void (async () => {
            const updated = await providerKeys.list();
            setActiveProviders(new Set(updated));
            document.dispatchEvent(new CustomEvent("pi:providers-changed"));
            showToast(`${label} disconnected`);
          })();
        },
      });
      providerList.appendChild(row);
    }

    dialog.mount();
  });
}
