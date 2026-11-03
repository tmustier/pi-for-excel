import { createConcurrencyLimit } from "../../../utils/concurrency.js";
import { getAppStorage } from "../../../storage/local/app-storage.js";

import {
  DEFAULT_PROXY_IS_REMOTE,
  DEFAULT_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  validateOfficeProxyUrl,
} from "../../../auth/proxy-validation.js";
import {
  checkProxyOnce,
  getProxyState,
  isProxyBlockedByBrowser,
  type ProxyState,
} from "../../../taskpane/proxy-status.js";
import { needsRemoteProxyWarning, saveTaskpaneProxyUrl } from "../../../taskpane/settings.js";
import { t } from "../../../language/index.js";
import {
  createCallout,
  createConfigInput,
  createConfigRow,
  createToggleRow,
} from "../../../ui/extensions-hub-components.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import { showToast } from "../../../ui/toast.js";
import type { SettingsAccess } from "../../../storage/local/settings-store.js";

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

  if (args.state === "not-detected" && isProxyBlockedByBrowser()) {
    return {
      tone: "warn",
      icon: "⚠",
      message: t("settings.section.proxy.blocked_by_browser", { url: args.proxyUrl }),
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
    subtitle: () => t("settings.section.proxy.page_subtitle"),
    render: (ctx) => {
      const settingsStore: SettingsAccess = getAppStorage().settings;

      const card = document.createElement("div");
      card.className = "pi-overlay-surface pi-settings-proxy-card";

      const serializeSave = createConcurrencyLimit(1);
      let disposed = false;
      let enabledRevision = 0;
      let urlRevision = 0;
      let persistedEnabled = false;
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

      let recheckToken = 0;
      let recheckState: { kind: "idle" } | { kind: "checking"; token: number } = { kind: "idle" };
      const recheckProxy = async (): Promise<void> => {
        const token = ++recheckToken;
        recheckState = { kind: "checking", token };
        proxyState = "unknown";
        updateStatus();
        const state = await checkProxyOnce(settingsStore, { force: true });
        if (recheckState.kind !== "checking" || recheckState.token !== token || disposed) return;
        recheckState = { kind: "idle" };
        proxyState = state;
        updateStatus();
      };

      const warnIfRemote = (): void => {
        if (needsRemoteProxyWarning({ enabled, url: proxyUrl })) {
          showToast(t("init.securityWarning"), { variant: "error" });
        }
      };

      const updateStatus = (): void => {
        const status = resolveProxyCallout({
          enabled,
          state: proxyState,
          proxyUrl,
          validationError,
        });

        statusHost.replaceChildren(createCallout(status.tone, status.icon, status.message, { compact: true }));
      };

      const saveProxyEnabled = (nextEnabled: boolean): Promise<void> => {
        const revision = ++enabledRevision;
        enabled = nextEnabled;
        updateStatus();

        return serializeSave(async () => {
          await initialRead;
          try {
            await settingsStore.set("proxy.enabled", nextEnabled);
            persistedEnabled = nextEnabled;
          } catch {
            if (disposed || revision !== enabledRevision) return;
            enabled = persistedEnabled;
            proxyToggle.input.checked = enabled;
            updateStatus();
            showToast(t("settings.toast.proxy_save_failed"));
            return;
          }

          if (disposed || revision !== enabledRevision) return;
          showToast(enabled ? t("settings.toast.proxy_enabled") : t("settings.toast.proxy_disabled"));
          warnIfRemote();
          if (enabled) void recheckProxy();
          else void checkProxyOnce(settingsStore, { force: true });
        });
      };

      const saveProxyUrl = (): Promise<void> => {
        const revision = ++urlRevision;
        const raw = proxyUrlInput.value.trim();
        const candidate = raw.length > 0 ? raw : DEFAULT_PROXY_URL;
        return serializeSave(async () => {
          await initialRead;

          let normalizedUrl: string;
          try {
            normalizedUrl = validateOfficeProxyUrl(candidate);
          } catch (error) {
            if (disposed || revision !== urlRevision) return;
            validationError = error instanceof Error ? error.message : t("settings.toast.proxy_url_invalid");
            updateStatus();
            showToast(t("settings.toast.proxy_url_not_saved", { error: validationError }));
            return;
          }

          validationError = null;

          if (normalizedUrl === proxyUrl) {
            if (disposed || revision !== urlRevision) return;
            proxyUrlInput.value = normalizedUrl;
            updateStatus();
            return;
          }

          try {
            await saveTaskpaneProxyUrl(settingsStore, normalizedUrl);
          } catch {
            if (disposed || revision !== urlRevision) return;
            showToast(t("settings.toast.proxy_url_save_failed"));
            return;
          }

          proxyUrl = normalizedUrl;
          if (disposed || revision !== urlRevision) return;
          proxyUrlInput.value = normalizedUrl;
          updateStatus();
          showToast(t("settings.toast.proxy_url_saved"));
          warnIfRemote();
          void recheckProxy();
        });
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

      proxyUrlInput.addEventListener("input", () => { urlRevision += 1; recheckToken += 1; recheckState = { kind: "idle" }; });
      proxyUrlInput.addEventListener("blur", scheduleProxyUrlSave);
      proxyUrlInput.addEventListener("keydown", (event) => {
        if (event.key !== "Enter") {
          return;
        }

        event.preventDefault();
        scheduleProxyUrlSave();
        proxyUrlInput.blur();
      });

      const onProxyStateChanged = (): void => {
        if (disposed || recheckState.kind === "checking") return;
        void recheckProxy();
      };

      document.addEventListener("pi:proxy-state-changed", onProxyStateChanged);
      ctx.addCleanup(() => {
        document.removeEventListener("pi:proxy-state-changed", onProxyStateChanged);
      });
      ctx.addCleanup(() => {
        flushPendingProxyUrlSave();
        disposed = true;
        recheckToken += 1;
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

      const initialRead = (async () => {
        try {
          const storedEnabled = await settingsStore.get("proxy.enabled");
          const storedUrl = await settingsStore.get("proxy.url");

          persistedEnabled = storedEnabled === true;
          if (disposed) return;
          if (enabledRevision === 0) enabled = persistedEnabled;
          if (urlRevision === 0) proxyUrl = typeof storedUrl === "string" && storedUrl.trim().length > 0
            ? storedUrl.trim()
            : DEFAULT_PROXY_URL;
        } catch {
          if (disposed) return;
          if (enabledRevision === 0) enabled = false;
          if (urlRevision === 0) proxyUrl = DEFAULT_PROXY_URL;
        }

        if (enabledRevision === 0) proxyToggle.input.checked = enabled;
        if (urlRevision === 0) proxyUrlInput.value = proxyUrl;
        updateStatus();
      })();

      card.append(proxyToggle.root, proxyUrlRow, statusHost, helper);
      ctx.body.appendChild(card);
    },
  };
}
