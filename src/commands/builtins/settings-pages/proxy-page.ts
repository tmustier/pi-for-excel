/**
 * Proxy page — local proxy helper toggle, URL, and live status.
 */

import { getAppStorage } from "../../../storage/local/app-storage.js";

import {
  DEFAULT_PROXY_IS_REMOTE,
  DEFAULT_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  validateOfficeProxyUrl,
} from "../../../auth/proxy-validation.js";
import { getProxyState, type ProxyState } from "../../../taskpane/proxy-status.js";
import { t } from "../../../language/index.js";
import {
  createCallout,
  createConfigInput,
  createConfigRow,
  createToggleRow,
} from "../../../ui/extensions-hub-components.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import { showToast } from "../../../ui/toast.js";

function isProxyPagePayloadShape(value: DynamicValue): value is DynamicObject {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

interface SettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: DynamicValue): Promise<void>;
}

function resolveProxyCallout(args: {
  enabled: boolean;
  state: ProxyState;
  proxyUrl: string;
  validationError: string | null;
}): {
  tone: "info" | "warn" | "success";
  icon: string;
  message: string;
} {
  if (args.validationError) {
    return { tone: "warn", icon: "⚠", message: args.validationError };
  }

  if (!args.enabled) {
    return { tone: "info", icon: "ℹ", message: t("settings.proxy.disabled") };
  }

  if (args.state === "detected") {
    return {
      tone: "success",
      icon: "✓",
      message: t("settings.proxy.connected", { url: args.proxyUrl }),
    };
  }

  if (args.state === "not-detected") {
    return {
      tone: "warn",
      icon: "⚠",
      message: t("settings.section.proxy.not_reachable", { url: args.proxyUrl }),
    };
  }

  return {
    tone: "info",
    icon: "…",
    message: t("settings.section.proxy.checking", { url: args.proxyUrl }),
  };
}

export function createProxyPage(): SettingsShellPage {
  return {
    id: "proxy",
    parentId: "root",
    title: () => t("settings.section.proxy"),
    render: (ctx) => {
      const settingsStore: SettingsStore = getAppStorage().settings;

      const card = document.createElement("div");
      card.className = "pi-overlay-surface pi-settings-proxy-card";

      let enabled = false;
      let proxyUrl = DEFAULT_PROXY_URL;
      let proxyState: ProxyState = getProxyState();
      let validationError: string | null = null;
      let urlSaveTimer: ReturnType<typeof setTimeout> | null = null;

      const proxyToggle = createToggleRow({
        label: t("settings.section.proxy.label"),
        sublabel: DEFAULT_PROXY_IS_REMOTE
          ? t("settings.section.proxy.sublabel_remote")
          : t("settings.section.proxy.sublabel"),
        checked: enabled,
        onChange: (checked) => {
          void saveProxyEnabled(checked);
        },
      });
      proxyToggle.root.classList.add("pi-settings-proxy-toggle");

      const proxyUrlInput = createConfigInput({
        value: proxyUrl,
        placeholder: DEFAULT_PROXY_URL,
      });
      proxyUrlInput.classList.add("pi-settings-proxy-url");
      proxyUrlInput.spellcheck = false;

      const proxyUrlRow = createConfigRow(t("settings.section.proxy.url"), proxyUrlInput);
      proxyUrlRow.classList.add("pi-settings-proxy-url-row");

      const statusHost = document.createElement("div");
      statusHost.className = "pi-settings-proxy-status";

      const updateStatus = (): void => {
        const status = resolveProxyCallout({
          enabled,
          state: proxyState,
          proxyUrl,
          validationError,
        });

        statusHost.replaceChildren(createCallout(status.tone, status.icon, status.message, { compact: true }));
      };

      const saveProxyEnabled = async (nextEnabled: boolean): Promise<void> => {
        enabled = nextEnabled;
        updateStatus();

        try {
          await settingsStore.set("proxy.enabled", enabled);
        } catch {
          enabled = !nextEnabled;
          proxyToggle.input.checked = enabled;
          updateStatus();
          showToast(t("settings.toast.proxy_save_failed"));
          return;
        }

        showToast(enabled ? t("settings.toast.proxy_enabled") : t("settings.toast.proxy_disabled"));
      };

      const saveProxyUrl = async (): Promise<void> => {
        const raw = proxyUrlInput.value.trim();
        const candidate = raw.length > 0 ? raw : DEFAULT_PROXY_URL;

        let normalizedUrl: string;
        try {
          normalizedUrl = validateOfficeProxyUrl(candidate);
        } catch (error) {
          validationError = error instanceof Error ? error.message : t("settings.toast.proxy_url_invalid");
          updateStatus();
          showToast(t("settings.toast.proxy_url_not_saved", { error: validationError }));
          return;
        }

        validationError = null;

        if (normalizedUrl === proxyUrl) {
          proxyUrlInput.value = normalizedUrl;
          updateStatus();
          return;
        }

        try {
          await settingsStore.set("proxy.url", normalizedUrl);
        } catch {
          showToast(t("settings.toast.proxy_url_save_failed"));
          return;
        }

        proxyUrl = normalizedUrl;
        proxyUrlInput.value = normalizedUrl;
        updateStatus();
        showToast(t("settings.toast.proxy_url_saved"));
      };

      const scheduleProxyUrlSave = (): void => {
        if (urlSaveTimer !== null) {
          clearTimeout(urlSaveTimer);
        }

        urlSaveTimer = setTimeout(() => {
          urlSaveTimer = null;
          void saveProxyUrl();
        }, 140);
      };

      const flushPendingProxyUrlSave = (): void => {
        if (urlSaveTimer === null) {
          return;
        }

        clearTimeout(urlSaveTimer);
        urlSaveTimer = null;
        void saveProxyUrl();
      };

      proxyUrlInput.addEventListener("blur", scheduleProxyUrlSave);
      proxyUrlInput.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") {
          return;
        }

        event.preventDefault();
        scheduleProxyUrlSave();
        proxyUrlInput.blur();
      });

      const onProxyStateChanged = (event: Event): void => {
        if (!(event instanceof CustomEvent)) return;

        const detail: DynamicValue = event.detail;
        if (!isProxyPagePayloadShape(detail)) return;

        const state = detail.state;
        if (state === "detected" || state === "not-detected" || state === "unknown") {
          proxyState = state;
          updateStatus();
        }
      };

      document.addEventListener("pi:proxy-state-changed", onProxyStateChanged);
      ctx.addCleanup(() => {
        document.removeEventListener("pi:proxy-state-changed", onProxyStateChanged);
      });
      ctx.addCleanup(() => {
        flushPendingProxyUrlSave();
      });

      const helper = document.createElement("p");
      helper.className = "pi-overlay-hint pi-settings-proxy-helper";

      const recommendedUrl = document.createElement("code");
      recommendedUrl.textContent = DEFAULT_PROXY_URL;

      const guideLink = document.createElement("a");
      guideLink.href = PROXY_HELPER_DOCS_URL;
      guideLink.target = "_blank";
      guideLink.rel = "noopener noreferrer";
      guideLink.textContent = t("settings.section.proxy.guide");

      helper.append(
        t("settings.section.proxy.recommended"),
        " ",
        recommendedUrl,
        DEFAULT_PROXY_IS_REMOTE
          ? t("settings.section.proxy.org_proxy")
          : t("settings.section.proxy.keep_localhost"),
        guideLink,
        ".",
      );

      void (async () => {
        try {
          const storedEnabled = await settingsStore.get<boolean>("proxy.enabled");
          const storedUrl = await settingsStore.get<string>("proxy.url");

          enabled = storedEnabled === true;
          proxyUrl = typeof storedUrl === "string" && storedUrl.trim().length > 0
            ? storedUrl.trim()
            : DEFAULT_PROXY_URL;
        } catch {
          enabled = false;
          proxyUrl = DEFAULT_PROXY_URL;
        }

        proxyToggle.input.checked = enabled;
        proxyUrlInput.value = proxyUrl;
        updateStatus();
      })();

      card.append(proxyToggle.root, proxyUrlRow, statusHost, helper);
      ctx.body.appendChild(card);
    },
  };
}
