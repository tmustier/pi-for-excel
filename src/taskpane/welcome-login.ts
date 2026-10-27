/**
 * Welcome/login overlay shown when no providers are configured.
 */

import { t, getLanguage } from "../language/index.js";

import type { ModelRefreshOwner } from "../models/model-refresh-owner.js";
import { getAppStorage } from "../storage/local/app-storage.js";

import { closeOverlayById, createOverlayDialog } from "../ui/overlay-dialog.js";
import { WELCOME_LOGIN_OVERLAY_ID } from "../ui/overlay-ids.js";
import { showToast } from "../ui/toast.js";
import {
  DEFAULT_PROXY_IS_REMOTE,
  DEFAULT_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  probeProxyReachability,
  resolveConfiguredProxyUrl,
  validateOfficeProxyUrl,
} from "../auth/proxy-validation.js";
import { checkProxyOnce } from "./proxy-status.js";
import { readTaskpaneProxySettings } from "./settings.js";

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

export async function showWelcomeLogin(modelRefreshOwner: ModelRefreshOwner): Promise<void> {
  const { VISIBLE_PROVIDERS, buildProviderRow } = await import("../ui/provider-login.js");

  // Make OAuth flows usable even before the user can access /settings.
  try {
    const storage = getAppStorage();
    const proxySettings = await readTaskpaneProxySettings(storage.settings);
    const currentUrl = resolveConfiguredProxyUrl(proxySettings.url);

    if (proxySettings.url === null) {
      await storage.settings.set("proxy.url", currentUrl);
    }

    // Auto-enable if a local HTTPS proxy is actually reachable.
    if (!proxySettings.enabled) {
      const ok = await testLocalHttpsProxy(currentUrl);
      if (ok) {
        await storage.settings.set("proxy.enabled", true);
      }
    }
  } catch {
    // ignore — welcome overlay should still show
  }

  closeOverlayById(WELCOME_LOGIN_OVERLAY_ID);
  if (modelRefreshOwner.snapshot().availableProviders.length > 0) return;

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
    const unsubscribeModels = modelRefreshOwner.subscribe((snapshot) => {
      if (snapshot.availableProviders.length > 0) closeOverlay();
    });
    dialog.addCleanup(unsubscribeModels);

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
    langBar.setAttribute("role", "group");
    langBar.setAttribute("aria-label", t("welcome.language.aria"));

    const createLanguageButton = (lang: "en" | "zh-CN", label: string): HTMLButtonElement => {
      const button = createElement("button", "pi-welcome-lang-btn");
      button.type = "button";
      button.textContent = label;
      button.setAttribute("aria-pressed", getLanguage() === lang ? "true" : "false");
      button.addEventListener("click", () => {
        if (getLanguage() === lang) return;
        void (async () => {
          // Save first: switching in memory before a failed save would leave a
          // mixed-language pane until the next reload.
          try {
            await getAppStorage().settings.set("language", lang);
          } catch {
            showToast(t("welcome.language.saveFailed"));
            return;
          }
          initLanguage(lang);
          location.reload();
        })();
      });
      return button;
    };

    const engBtn = createLanguageButton("en", t("language.english"));
    const zhBtn = createLanguageButton("zh-CN", "中文");

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

      void import("../commands/builtins/settings-pages/index.js")
        .then(({ openSettings }) => {
          void openSettings("gateway");
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
        ? t("welcome.proxy.hide")
        : proxyToggleClosedLabel;

      if (willOpen) {
        proxyUrlEl.focus();
      }
    });

    const hydrateProxyUi = async () => {
      const storage = getAppStorage();
      const proxySettings = await readTaskpaneProxySettings(storage.settings);
      proxyEnabledEl.checked = proxySettings.enabled;
      proxyUrlEl.value = resolveConfiguredProxyUrl(proxySettings.url);
    };

    const saveProxyUi = async () => {
      const storage = getAppStorage();
      let proxyUrl: string;
      try {
        // Same validation as Settings → Proxy: an unusable URL is never stored.
        proxyUrl = validateOfficeProxyUrl(proxyUrlEl.value.trim() || DEFAULT_PROXY_URL);
      } catch (error) {
        showToast(t("welcome.toast.proxy_url_not_saved", { error: getErrorMessage(error) }), { variant: "error" });
        return;
      }

      try {
        await storage.settings.set("proxy.enabled", proxyEnabledEl.checked);
        await storage.settings.set("proxy.url", proxyUrlEl.value.trim());
        showToast(t("welcome.toast.proxy_saved"));
        void checkProxyOnce(storage.settings, { force: true });
      } catch {
        showToast(t("welcome.toast.proxy_failed"));
        return;
      }

      showToast(t("welcome.toast.proxy_saved"));
      if (needsRemoteProxyWarning({ enabled: proxyEnabledEl.checked, url: proxyUrl })) {
        showToast(t("init.securityWarning"), { variant: "error" });
      }
      void checkProxyOnce(storage.settings);
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
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
          showToast(t("welcome.toast.connected", { label }), 3200);
          closeOverlay();
        },
        onDisconnected: (_row, _id, label) => {
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
          showToast(t("welcome.toast.disconnected", { label }));
        },
      });
      providerList.appendChild(row);
    }

    dialog.mount();
  });
}
