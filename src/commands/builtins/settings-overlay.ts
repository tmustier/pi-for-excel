/**
 * Unified settings overlay.
 *
 * Sections:
 * - Providers
 * - Proxy
 * - Experimental
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  DEFAULT_LOCAL_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  validateOfficeProxyUrl,
} from "../../auth/proxy-validation.js";
import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  createOverlayInput,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { SETTINGS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { ALL_PROVIDERS, buildProviderRow } from "../../ui/provider-login.js";
import { showToast } from "../../ui/toast.js";
import {
  buildExperimentalFeatureContent,
  buildExperimentalFeatureFooter,
} from "./experimental-overlay.js";

export type SettingsOverlaySection = "providers" | "proxy" | "experimental";

export interface ShowSettingsDialogOptions {
  section?: SettingsOverlaySection;
}

interface SettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
}

let settingsDialogOpenInFlight: Promise<void> | null = null;
let pendingSectionFocus: SettingsOverlaySection | null = null;

function sectionSelector(section: SettingsOverlaySection): string {
  return `[data-settings-section=\"${section}\"]`;
}

function focusSettingsSection(overlay: HTMLElement, section: SettingsOverlaySection): void {
  const target = overlay.querySelector<HTMLElement>(sectionSelector(section));
  if (!target) {
    return;
  }

  target.scrollIntoView({ behavior: "smooth", block: "start" });
}

function createSectionShell(titleText: string, section: SettingsOverlaySection, hintText?: string): {
  section: HTMLElement;
  content: HTMLDivElement;
} {
  const sectionEl = document.createElement("section");
  sectionEl.className = "pi-overlay-section pi-settings-section";
  sectionEl.dataset.settingsSection = section;

  const title = createOverlaySectionTitle(titleText);
  sectionEl.appendChild(title);

  if (hintText) {
    const hint = document.createElement("p");
    hint.className = "pi-overlay-hint";
    hint.textContent = hintText;
    sectionEl.appendChild(hint);
  }

  const content = document.createElement("div");
  content.className = "pi-settings-section__content";
  sectionEl.appendChild(content);

  return { section: sectionEl, content };
}

async function buildProvidersSection(): Promise<HTMLElement> {
  const shell = createSectionShell(
    "Providers",
    "providers",
    "Connect providers to use their models.",
  );

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
    warning.textContent = "Saved provider state is temporarily unavailable. You can still connect providers.";
    shell.content.appendChild(warning);
  }

  const expandedRef: { current: HTMLElement | null } = { current: null };

  for (const provider of ALL_PROVIDERS) {
    const row = buildProviderRow(provider, {
      isActive: configuredSet.has(provider.id),
      expandedRef,
      onConnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} connected`);
      },
      onDisconnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} disconnected`);
      },
    });

    providerList.appendChild(row);
  }

  shell.content.appendChild(providerList);
  return shell.section;
}

function buildProxySection(settingsStore: SettingsStore): HTMLElement {
  const shell = createSectionShell(
    "Proxy",
    "proxy",
    "Use a local HTTPS proxy only when OAuth is blocked by CORS.",
  );

  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-settings-proxy-card";

  const controlsRow = document.createElement("div");
  controlsRow.className = "pi-settings-proxy-row";

  const enabledLabel = document.createElement("label");
  enabledLabel.className = "pi-settings-proxy-enabled";

  const enabledInput = document.createElement("input");
  enabledInput.type = "checkbox";

  const enabledText = document.createElement("span");
  enabledText.textContent = "Enable proxy";

  enabledLabel.append(enabledInput, enabledText);

  const urlInput = createOverlayInput({
    placeholder: DEFAULT_LOCAL_PROXY_URL,
    className: "pi-settings-proxy-url",
  });
  urlInput.type = "text";
  urlInput.spellcheck = false;

  const saveButton = createOverlayButton({
    text: "Save",
    className: "pi-overlay-btn--primary",
  });

  controlsRow.append(enabledLabel, urlInput, saveButton);

  const status = document.createElement("p");
  status.className = "pi-overlay-hint pi-settings-proxy-status";

  const helper = document.createElement("p");
  helper.className = "pi-overlay-hint";

  const guideLink = document.createElement("a");
  guideLink.href = PROXY_HELPER_DOCS_URL;
  guideLink.target = "_blank";
  guideLink.rel = "noopener noreferrer";
  guideLink.textContent = "Step-by-step guide";

  helper.append(
    "Recommended URL: ",
    (() => {
      const code = document.createElement("code");
      code.textContent = DEFAULT_LOCAL_PROXY_URL;
      return code;
    })(),
    ". Keep this on localhost. ",
    guideLink,
    ".",
  );

  const save = async (): Promise<void> => {
    const rawUrl = urlInput.value.trim();
    const candidateUrl = rawUrl.length > 0 ? rawUrl : DEFAULT_LOCAL_PROXY_URL;

    let normalizedUrl: string;
    try {
      normalizedUrl = validateOfficeProxyUrl(candidateUrl);
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : "Invalid proxy URL";
      status.textContent = message;
      status.classList.add("pi-overlay-text-warning");
      showToast(`Proxy not saved: ${message}`);
      return;
    }

    await settingsStore.set("proxy.enabled", enabledInput.checked);
    await settingsStore.set("proxy.url", normalizedUrl);

    urlInput.value = normalizedUrl;
    status.textContent = enabledInput.checked
      ? `Proxy enabled at ${normalizedUrl}`
      : `Proxy saved at ${normalizedUrl} (currently disabled)`;
    status.classList.remove("pi-overlay-text-warning");
    showToast("Proxy settings saved");
  };

  saveButton.addEventListener("click", () => {
    void save();
  });
  enabledInput.addEventListener("change", () => {
    void save();
  });
  urlInput.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }

    event.preventDefault();
    void save();
  });

  void (async () => {
    try {
      const enabled = await settingsStore.get<boolean>("proxy.enabled");
      const storedUrl = await settingsStore.get<string>("proxy.url");

      enabledInput.checked = enabled === true;
      urlInput.value = typeof storedUrl === "string" && storedUrl.trim().length > 0
        ? storedUrl.trim()
        : DEFAULT_LOCAL_PROXY_URL;

      status.textContent = enabledInput.checked
        ? `Proxy enabled at ${urlInput.value}`
        : "Proxy disabled";
    } catch {
      enabledInput.checked = false;
      urlInput.value = DEFAULT_LOCAL_PROXY_URL;
      status.textContent = "Proxy disabled";
    }
  })();

  card.append(controlsRow, status, helper);
  shell.content.appendChild(card);
  return shell.section;
}

function buildExperimentalSection(): HTMLElement {
  const shell = createSectionShell(
    "Experimental",
    "experimental",
    "Advanced and in-progress capabilities.",
  );

  const content = buildExperimentalFeatureContent();
  shell.content.appendChild(content);
  shell.content.appendChild(buildExperimentalFeatureFooter());

  return shell.section;
}

export async function showSettingsDialog(options: ShowSettingsDialogOptions = {}): Promise<void> {
  const existing = document.getElementById(SETTINGS_OVERLAY_ID);
  if (existing instanceof HTMLElement) {
    if (options.section) {
      focusSettingsSection(existing, options.section);
      return;
    }

    closeOverlayById(SETTINGS_OVERLAY_ID);
    return;
  }

  if (settingsDialogOpenInFlight) {
    if (options.section) {
      pendingSectionFocus = options.section;
    }

    await settingsDialogOpenInFlight;

    const mounted = document.getElementById(SETTINGS_OVERLAY_ID);
    if (mounted instanceof HTMLElement && options.section) {
      focusSettingsSection(mounted, options.section);
    }
    return;
  }

  pendingSectionFocus = options.section ?? pendingSectionFocus;

  settingsDialogOpenInFlight = (async () => {
    const appStorage = getAppStorage();

    const dialog = createOverlayDialog({
      overlayId: SETTINGS_OVERLAY_ID,
      cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-settings-dialog",
    });

    const { header } = createOverlayHeader({
      onClose: dialog.close,
      closeLabel: "Close settings",
      title: "Settings",
      subtitle: "Providers, proxy, and experimental features.",
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body pi-settings-body";

    const providersSection = await buildProvidersSection();
    const proxySection = buildProxySection(appStorage.settings);
    const experimentalSection = buildExperimentalSection();

    body.append(providersSection, proxySection, experimentalSection);
    dialog.card.append(header, body);
    dialog.mount();

    if (pendingSectionFocus) {
      const section = pendingSectionFocus;
      pendingSectionFocus = null;
      requestAnimationFrame(() => {
        const mounted = document.getElementById(SETTINGS_OVERLAY_ID);
        if (mounted instanceof HTMLElement) {
          focusSettingsSection(mounted, section);
        }
      });
    }
  })();

  try {
    await settingsDialogOpenInFlight;
  } finally {
    settingsDialogOpenInFlight = null;
  }
}
