/**
 * Unified settings overlay.
 *
 * Tabs:
 * - Providers (API keys, proxy)
 * - More (Advanced, Experimental)
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

type LegacyExtensionsSection = "connections" | "plugins" | "skills";
type SettingsPrimaryTab = "logins" | "extensions" | "more";

export type SettingsOverlaySection =
  | SettingsPrimaryTab
  | "providers"
  | "proxy"
  | "advanced"
  | "experimental"
  | LegacyExtensionsSection;

export interface ShowSettingsDialogOptions {
  section?: SettingsOverlaySection;
}

interface SettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
}

interface SettingsDialogDependencies {
  openRulesDialog?: () => Promise<void> | void;
  openRecoveryDialog?: () => Promise<void> | void;
  openShortcutsDialog?: () => void;
}

interface ResolvedSectionFocus {
  tab: SettingsPrimaryTab;
  anchor?: "proxy" | "providers" | "advanced" | "experimental";
}

const SETTINGS_TABS: ReadonlyArray<{ id: SettingsPrimaryTab; label: string }> = [
  { id: "logins", label: "Providers" },
  { id: "more", label: "More" },
];



let settingsDialogOpenInFlight: Promise<void> | null = null;
let pendingSectionFocus: SettingsOverlaySection | null = null;
let dependencies: SettingsDialogDependencies = {};

export function configureSettingsDialogDependencies(next: SettingsDialogDependencies): void {
  dependencies = { ...next };
}

function resolveSectionFocus(section: SettingsOverlaySection | undefined): ResolvedSectionFocus {
  switch (section) {
    case "providers":
      return { tab: "logins", anchor: "providers" };
    case "proxy":
      return { tab: "logins", anchor: "proxy" };
    case "advanced":
      return { tab: "more", anchor: "advanced" };
    case "experimental":
      return { tab: "more", anchor: "experimental" };
    case "more":
      return { tab: "more" };
    case "connections":
    case "plugins":
    case "skills":
    case "extensions":
    case "logins":
    default:
      return { tab: "logins" };
  }
}

function activateSettingsTab(overlay: HTMLElement, tab: SettingsPrimaryTab): void {
  const tabButtons = overlay.querySelectorAll<HTMLButtonElement>("[data-settings-tab]");
  for (const button of tabButtons) {
    const isActive = button.dataset.settingsTab === tab;
    button.classList.toggle("is-active", isActive);
    button.setAttribute("aria-selected", isActive ? "true" : "false");
  }

  const tabPanels = overlay.querySelectorAll<HTMLElement>("[data-settings-panel]");
  for (const panel of tabPanels) {
    panel.hidden = panel.dataset.settingsPanel !== tab;
  }
}

function applySectionFocus(overlay: HTMLElement, section: SettingsOverlaySection): void {
  const resolved = resolveSectionFocus(section);
  activateSettingsTab(overlay, resolved.tab);

  if (!resolved.anchor) {
    return;
  }

  const target = overlay.querySelector<HTMLElement>(`[data-settings-anchor="${resolved.anchor}"]`);
  if (!target) {
    return;
  }

  target.scrollIntoView({ behavior: "smooth", block: "start" });
}

function createSectionShell(titleText: string, anchor: string, hintText?: string): {
  section: HTMLElement;
  content: HTMLDivElement;
} {
  const sectionEl = document.createElement("section");
  sectionEl.className = "pi-overlay-section pi-settings-section";
  sectionEl.dataset.settingsAnchor = anchor;

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
    "Route API calls through a local proxy.",
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
  enabledText.textContent = "Route API calls through a local proxy";

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

function buildMoreSection(): HTMLElement {
  const panel = document.createElement("div");
  panel.className = "pi-settings-more";

  const advanced = createSectionShell(
    "Advanced",
    "advanced",
    "Power-user shortcuts for rules, backups, and keyboard shortcuts.",
  );

  const advancedActions = document.createElement("div");
  advancedActions.className = "pi-overlay-actions pi-settings-advanced-actions";

  const rulesButton = createOverlayButton({ text: "Rules & conventions…" });
  const backupsButton = createOverlayButton({ text: "Backups…" });
  const shortcutsButton = createOverlayButton({ text: "Keyboard shortcuts…" });

  rulesButton.disabled = !dependencies.openRulesDialog;
  backupsButton.disabled = !dependencies.openRecoveryDialog;
  shortcutsButton.disabled = !dependencies.openShortcutsDialog;

  rulesButton.addEventListener("click", () => {
    void dependencies.openRulesDialog?.();
  });
  backupsButton.addEventListener("click", () => {
    void dependencies.openRecoveryDialog?.();
  });
  shortcutsButton.addEventListener("click", () => {
    dependencies.openShortcutsDialog?.();
  });

  advancedActions.append(rulesButton, backupsButton, shortcutsButton);
  advanced.content.appendChild(advancedActions);

  const experimental = createSectionShell(
    "Experimental",
    "experimental",
    "Advanced and in-progress capabilities.",
  );
  experimental.content.appendChild(buildExperimentalFeatureContent());
  experimental.content.appendChild(buildExperimentalFeatureFooter());

  panel.append(advanced.section, experimental.section);
  return panel;
}

export async function showSettingsDialog(options: ShowSettingsDialogOptions = {}): Promise<void> {
  const existing = document.getElementById(SETTINGS_OVERLAY_ID);
  if (existing instanceof HTMLElement) {
    if (options.section) {
      applySectionFocus(existing, options.section);
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
      applySectionFocus(mounted, options.section);
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
      subtitle: "Providers, proxy, and preferences",
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body pi-settings-body";

    const tabs = document.createElement("div");
    tabs.className = "pi-overlay-tabs";
    tabs.setAttribute("role", "tablist");
    tabs.setAttribute("aria-label", "Settings tabs");

    const panels = document.createElement("div");
    panels.className = "pi-settings-panels";

    const loginsPanel = document.createElement("div");
    loginsPanel.className = "pi-settings-panel";
    loginsPanel.dataset.settingsPanel = "logins";
    loginsPanel.append(
      buildProxySection(appStorage.settings),
      await buildProvidersSection(),
    );

    const morePanel = document.createElement("div");
    morePanel.className = "pi-settings-panel";
    morePanel.dataset.settingsPanel = "more";
    morePanel.appendChild(buildMoreSection());

    panels.append(loginsPanel, morePanel);

    for (const tab of SETTINGS_TABS) {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "pi-overlay-tab";
      button.textContent = tab.label;
      button.dataset.settingsTab = tab.id;
      button.setAttribute("role", "tab");
      button.setAttribute("aria-selected", "false");
      button.addEventListener("click", () => {
        activateSettingsTab(dialog.overlay, tab.id);
      });
      tabs.appendChild(button);
    }

    body.append(tabs, panels);
    dialog.card.append(header, body);
    dialog.mount();

    const initialSection = pendingSectionFocus ?? "logins";
    pendingSectionFocus = null;
    requestAnimationFrame(() => {
      const mounted = document.getElementById(SETTINGS_OVERLAY_ID);
      if (mounted instanceof HTMLElement) {
        applySectionFocus(mounted, initialSection);
      }
    });
  })();

  try {
    await settingsDialogOpenInFlight;
  } finally {
    settingsDialogOpenInFlight = null;
  }
}
