/**
 * Unified settings overlay.
 *
 * Tabs:
 * - Logins (Proxy, Providers)
 * - Extensions (opens unified Extensions manager)
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
import type { AddonsSection } from "./addons-overlay.js";

type SettingsPrimaryTab = "logins" | "extensions" | "more";

export type SettingsOverlaySection =
  | SettingsPrimaryTab
  | "providers"
  | "proxy"
  | "advanced"
  | "experimental"
  | AddonsSection;

export interface ShowSettingsDialogOptions {
  section?: SettingsOverlaySection;
}

interface SettingsStore {
  get<T>(key: string): Promise<T | null>;
  set(key: string, value: unknown): Promise<void>;
}

interface SettingsDialogDependencies {
  openExtensionsHub?: (section?: AddonsSection) => void;
  openRulesDialog?: () => Promise<void> | void;
  openRecoveryDialog?: () => Promise<void> | void;
  openShortcutsDialog?: () => void;
}

interface ResolvedSectionFocus {
  tab: SettingsPrimaryTab;
  anchor?: "proxy" | "providers" | "advanced" | "experimental";
  extensionSection?: AddonsSection;
}

const SETTINGS_TABS: ReadonlyArray<{ id: SettingsPrimaryTab; label: string }> = [
  { id: "logins", label: "Logins" },
  { id: "extensions", label: "Extensions" },
  { id: "more", label: "More" },
];

const EXTENSIONS_LINKS: ReadonlyArray<{
  section: AddonsSection;
  label: string;
  description: string;
}> = [
  {
    section: "connections",
    label: "Connections",
    description: "Web search, MCP, and bridge setup",
  },
  {
    section: "plugins",
    label: "Plugins",
    description: "Installed plugins and enable/disable state",
  },
  {
    section: "skills",
    label: "Skills",
    description: "Bundled + external skill catalog",
  },
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
    case "connections":
    case "plugins":
    case "skills":
      return { tab: "extensions", extensionSection: section };
    case "advanced":
      return { tab: "more", anchor: "advanced" };
    case "experimental":
      return { tab: "more", anchor: "experimental" };
    case "extensions":
      return { tab: "extensions" };
    case "more":
      return { tab: "more" };
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

  if (resolved.extensionSection) {
    const preferred = overlay.querySelector<HTMLButtonElement>(
      `[data-settings-extension-link="${resolved.extensionSection}"]`,
    );
    if (preferred) {
      preferred.click();
    }
  }

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

function buildExtensionsSection(closeDialog: () => void): HTMLElement {
  const shell = createSectionShell(
    "Extensions",
    "extensions",
    "Connections, plugins, and skills live in one place.",
  );

  const tabs = document.createElement("div");
  tabs.className = "pi-overlay-tabs";
  tabs.setAttribute("role", "tablist");
  tabs.setAttribute("aria-label", "Extensions sections");

  const description = document.createElement("p");
  description.className = "pi-overlay-hint";

  let selectedSection: AddonsSection = "connections";

  const applySelection = (section: AddonsSection): void => {
    selectedSection = section;
    for (const button of tabs.querySelectorAll<HTMLButtonElement>("[data-settings-extension-link]")) {
      const isActive = button.dataset.settingsExtensionLink === selectedSection;
      button.classList.toggle("is-active", isActive);
      button.setAttribute("aria-selected", isActive ? "true" : "false");
    }

    const selected = EXTENSIONS_LINKS.find((item) => item.section === section);
    description.textContent = selected ? selected.description : "";
  };

  for (const item of EXTENSIONS_LINKS) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-overlay-tab";
    button.textContent = item.label;
    button.dataset.settingsExtensionLink = item.section;
    button.setAttribute("role", "tab");
    button.setAttribute("aria-selected", "false");
    button.addEventListener("click", () => {
      applySelection(item.section);
    });
    tabs.appendChild(button);
  }

  applySelection(selectedSection);

  const actionRow = document.createElement("div");
  actionRow.className = "pi-overlay-actions pi-settings-extensions-actions";

  const openButton = createOverlayButton({
    text: "Open Extensions manager…",
    className: "pi-overlay-btn--primary",
  });

  if (dependencies.openExtensionsHub) {
    openButton.addEventListener("click", () => {
      closeDialog();
      dependencies.openExtensionsHub?.(selectedSection);
    });
  } else {
    openButton.disabled = true;
  }

  const aliasHint = document.createElement("p");
  aliasHint.className = "pi-overlay-hint";
  aliasHint.textContent = "Slash commands: /extensions, /addons, /tools, /integrations, /plugins, /skills";

  shell.content.append(tabs, description, actionRow, aliasHint);
  actionRow.appendChild(openButton);

  if (!dependencies.openExtensionsHub) {
    const warning = document.createElement("p");
    warning.className = "pi-overlay-hint pi-overlay-text-warning";
    warning.textContent = "Extensions manager is unavailable in this context.";
    shell.content.appendChild(warning);
  }

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
      subtitle: "Logins, extensions, and advanced options.",
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

    const extensionsPanel = document.createElement("div");
    extensionsPanel.className = "pi-settings-panel";
    extensionsPanel.dataset.settingsPanel = "extensions";
    extensionsPanel.appendChild(buildExtensionsSection(dialog.close));

    const morePanel = document.createElement("div");
    morePanel.className = "pi-settings-panel";
    morePanel.dataset.settingsPanel = "more";
    morePanel.appendChild(buildMoreSection());

    panels.append(loginsPanel, extensionsPanel, morePanel);

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
