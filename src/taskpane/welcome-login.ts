/**
 * Welcome/login overlay shown when no providers are configured.
 */

import { t, initLanguage, getLanguage } from "../language/index.js";

import type { ProviderKeysStore } from "@earendil-works/pi-web-ui/dist/storage/stores/provider-keys-store.js";
import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import { closeOverlayById, createOverlayDialog } from "../ui/overlay-dialog.js";
import { WELCOME_LOGIN_OVERLAY_ID } from "../ui/overlay-ids.js";
import { showToast } from "../ui/toast.js";
import { setActiveProviders } from "../compat/model-selector-patch.js";
import {
  DEFAULT_PROXY_IS_REMOTE,
  DEFAULT_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  probeProxyReachability,
  resolveConfiguredProxyUrl,
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
  return probeProxyReachability(proxyUrl, 1200);
}

export async function showWelcomeLogin(providerKeys: ProviderKeysStore): Promise<void> {
  const { VISIBLE_PROVIDERS, buildProviderRow } = await import("../ui/provider-login.js");

  // Make OAuth flows usable even before the user can access /settings.
  try {
    const storage = getAppStorage();
    const enabled = await storage.settings.get("proxy.enabled");
    const url = await storage.settings.get("proxy.url");

    const currentUrl = resolveConfiguredProxyUrl(url);

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
    title.textContent = t("welcome.title");

    const subtitle = createElement("p", "pi-welcome-subtitle");
    subtitle.id = subtitleId;
    subtitle.textContent = t("welcome.subtitle");

    const intro = createElement("p", "pi-welcome-intro");
    intro.textContent = t("welcome.intro");

    const providerSectionTitle = createElement("p", "pi-welcome-section-title");
    providerSectionTitle.textContent = t("welcome.select_provider");

    const providerList = createElement("div", "pi-welcome-providers");

    const customGatewayButton = createElement("button", "pi-welcome-custom-gateway");
    customGatewayButton.type = "button";
    customGatewayButton.textContent = t("welcome.custom_gateway");

    const proxyToggle = createElement("button", "pi-welcome-proxy-toggle");
    proxyToggle.type = "button";
    const proxyToggleClosedLabel = DEFAULT_PROXY_IS_REMOTE
      ? t("welcome.proxy.toggle_show_remote")
      : t("welcome.proxy.toggle_show");
    proxyToggle.textContent = proxyToggleClosedLabel;
    proxyToggle.setAttribute("aria-expanded", "false");

    const proxyPanel = createElement("section", "pi-welcome-proxy");
    proxyPanel.hidden = true;

    const proxyTopRow = createElement("div", "pi-welcome-proxy__row");

    const proxyTitle = createElement("div", "pi-welcome-proxy__title");
    proxyTitle.textContent = DEFAULT_PROXY_IS_REMOTE
      ? t("welcome.proxy.title_remote")
      : t("welcome.proxy.title");

    const proxyToggleLabel = createElement("label", "pi-welcome-proxy__toggle");
    const proxyEnabledEl = createElement("input", "pi-welcome-proxy__enabled");
    proxyEnabledEl.type = "checkbox";
    const proxyToggleText = createElement("span");
    proxyToggleText.textContent = t("welcome.proxy.enabled");
    proxyToggleLabel.append(proxyEnabledEl, proxyToggleText);

    proxyTopRow.append(proxyTitle, proxyToggleLabel);

    const proxyUrlRow = createElement("div", "pi-welcome-proxy__row pi-welcome-proxy__row--compact");
    const proxyUrlEl = createElement("input", "pi-welcome-proxy__url");
    proxyUrlEl.type = "text";
    proxyUrlEl.spellcheck = false;

    const proxySaveEl = createElement("button", "pi-welcome-proxy__save");
    proxySaveEl.type = "button";
    proxySaveEl.textContent = t("welcome.proxy.save");

    proxyUrlRow.append(proxyUrlEl, proxySaveEl);

    const proxyHint = createElement("p", "pi-welcome-proxy__hint");
    const proxyCode = createElement("code");
    proxyCode.textContent = DEFAULT_PROXY_URL;

    const proxyGuideLink = createElement("a");
    proxyGuideLink.href = PROXY_HELPER_DOCS_URL;
    proxyGuideLink.target = "_blank";
    proxyGuideLink.rel = "noopener noreferrer";
    proxyGuideLink.textContent = t("welcome.proxy.guide");

    proxyHint.append(
      t("welcome.proxy.hint.prefix"),
      proxyCode,
      DEFAULT_PROXY_IS_REMOTE
        ? t("welcome.proxy.hint.suffix_remote")
        : t("welcome.proxy.hint.suffix"),
      proxyGuideLink,
      t("welcome.proxy.hint.end"),
    );

    proxyPanel.append(proxyTopRow, proxyUrlRow, proxyHint);


    // Language bar at the top
    const langBar = createElement("div", "pi-welcome-lang-bar");
    langBar.style.cssText = "display:flex;justify-content:flex-end;gap:4px;padding:4px 8px;";

    const engBtn = createElement("button");
    engBtn.type = "button";
    engBtn.textContent = t("language.english");
    engBtn.style.cssText = "font-size:11px;padding:2px 8px;border:1px solid #ccc;border-radius:4px;background:var(--pi-bg, #fff);cursor:pointer;";

    const zhBtn = createElement("button");
    zhBtn.type = "button";
    zhBtn.textContent = "中文";
    zhBtn.style.cssText = "font-size:11px;padding:2px 8px;border:1px solid #ccc;border-radius:4px;background:var(--pi-bg, #fff);cursor:pointer;";

    const currentLang2 = getLanguage();
    if (currentLang2 === "zh-CN") {
      zhBtn.style.borderColor = "var(--color-accent, #3b82f6)";
      zhBtn.style.color = "var(--color-accent, #3b82f6)";
    } else {
      engBtn.style.borderColor = "var(--color-accent, #3b82f6)";
      engBtn.style.color = "var(--color-accent, #3b82f6)";
    }

    engBtn.addEventListener("click", () => {
      if (getLanguage() === "en") return;
      initLanguage("en");
      void (async () => {
        try {
          const storage = getAppStorage();
          await storage.settings.set("language", "en");
          location.reload();
        } catch { /* ignore */ }
      })();
    });

    zhBtn.addEventListener("click", () => {
      if (getLanguage() === "zh-CN") return;
      initLanguage("zh-CN");
      void (async () => {
        try {
          const storage = getAppStorage();
          await storage.settings.set("language", "zh-CN");
          location.reload();
        } catch { /* ignore */ }
      })();
    });

    langBar.append(engBtn, zhBtn);

    dialog.card.replaceChildren(
      langBar,
      logo,
      title,
      subtitle,
      intro,
      providerSectionTitle,
      providerList,
      customGatewayButton,
      proxyToggle,
      proxyPanel,
    );

    dialog.overlay.setAttribute("aria-labelledby", titleId);
    dialog.overlay.setAttribute("aria-describedby", subtitleId);

    customGatewayButton.addEventListener("click", () => {
      closeOverlay();

      void import("../commands/builtins/settings-overlay.js")
        .then(({ showSettingsDialog }) => {
          void showSettingsDialog({ section: "custom-gateways" });
        })
        .catch(() => {
          showToast(t("welcome.toast.cannot_open_settings"));
        });
    });

    proxyToggle.addEventListener("click", () => {
      const willOpen = proxyPanel.hidden;
      proxyPanel.hidden = !willOpen;
      proxyToggle.setAttribute("aria-expanded", willOpen ? "true" : "false");
      proxyToggle.textContent = willOpen
        ? "Hide proxy settings"
        : proxyToggleClosedLabel;

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
        proxyUrlEl.value = resolveConfiguredProxyUrl(url);
      } catch {
        proxyEnabledEl.checked = false;
        proxyUrlEl.value = DEFAULT_PROXY_URL;
      }
    };

    const saveProxyUi = async () => {
      try {
        const storage = getAppStorage();
        await storage.settings.set("proxy.enabled", proxyEnabledEl.checked);
        await storage.settings.set("proxy.url", proxyUrlEl.value.trim());
        showToast(t("welcome.toast.proxy_saved"));
      } catch {
        showToast(t("welcome.toast.proxy_failed"));
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

    for (const provider of VISIBLE_PROVIDERS) {
      const row = buildProviderRow(provider, {
        isActive: false,
        expandedRef,
        onConnected: (_row, _id, label) => {
          void (async () => {
            const updated = await providerKeys.list();
            setActiveProviders(new Set(updated));
            document.dispatchEvent(new CustomEvent("pi:providers-changed"));
            showToast(t("welcome.toast.connected", { label }), 3200);
            closeOverlay();
          })();
        },
        onDisconnected: (_row, _id, label) => {
          void (async () => {
            const updated = await providerKeys.list();
            setActiveProviders(new Set(updated));
            document.dispatchEvent(new CustomEvent("pi:providers-changed"));
            showToast(t("welcome.toast.disconnected", { label }));
          })();
        },
      });
      providerList.appendChild(row);
    }

    dialog.mount();
  });
}
