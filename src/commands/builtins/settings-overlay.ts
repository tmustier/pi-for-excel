/**
 * Unified settings overlay.
 *
 * Tabs:
 * - Providers (API keys, proxy)
 * - More (execution mode, advanced, experimental)
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import {
  DEFAULT_LOCAL_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  validateOfficeProxyUrl,
} from "../../auth/proxy-validation.js";
import { PI_EXECUTION_MODE_CHANGED_EVENT, type ExecutionMode } from "../../execution/mode.js";
import type { ModelSwitchBehavior } from "../../models/switch-behavior.js";
import { getProxyState, type ProxyState } from "../../taskpane/proxy-status.js";
import {
  createCallout,
  createConfigInput,
  createConfigRow,
  createToggleRow,
} from "../../ui/extensions-hub-components.js";
import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { SETTINGS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { ALL_PROVIDERS, buildProviderRow } from "../../ui/provider-login.js";
import { showToast } from "../../ui/toast.js";
import { isRecord } from "../../utils/type-guards.js";
import {
  buildExperimentalFeatureContent,
  buildExperimentalFeatureFooter,
} from "./experimental-overlay.js";
import { buildCustomGatewaySection } from "./custom-gateway-settings.js";

type LegacyExtensionsSection = "connections" | "plugins" | "skills";
type SettingsPrimaryTab = "logins" | "more";

type SettingsAnchor =
  | "proxy"
  | "providers"
  | "custom-gateways"
  | "execution-mode"
  | "advanced"
  | "experimental";

type SettingsCleanupRegistrar = (cleanup: () => void) => void;

export type SettingsOverlaySection =
  | SettingsPrimaryTab
  | "providers"
  | "custom-gateways"
  | "proxy"
  | "execution-mode"
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
  getExecutionMode?: () => ExecutionMode;
  setExecutionMode?: (mode: ExecutionMode) => Promise<void>;
  getModelSwitchBehavior?: () => ModelSwitchBehavior;
  setModelSwitchBehavior?: (behavior: ModelSwitchBehavior) => Promise<void>;
}

interface ResolvedSectionFocus {
  tab: SettingsPrimaryTab;
  anchor?: SettingsAnchor;
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
    case "custom-gateways":
      return { tab: "logins", anchor: "custom-gateways" };
    case "proxy":
      return { tab: "logins", anchor: "proxy" };
    case "execution-mode":
      return { tab: "more", anchor: "execution-mode" };
    case "advanced":
      return { tab: "more", anchor: "advanced" };
    case "experimental":
      return { tab: "more", anchor: "experimental" };
    case "more":
      return { tab: "more" };
    case "connections":
    case "plugins":
    case "skills":
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

function createSectionShell(titleText: string, anchor: SettingsAnchor, hintText?: string): {
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
    return {
      tone: "warn",
      icon: "⚠",
      message: args.validationError,
    };
  }

  if (!args.enabled) {
    return {
      tone: "info",
      icon: "ℹ",
      message: "Proxy disabled.",
    };
  }

  if (args.state === "detected") {
    return {
      tone: "success",
      icon: "✓",
      message: `Proxy connected at ${args.proxyUrl}`,
    };
  }

  if (args.state === "not-detected") {
    return {
      tone: "warn",
      icon: "⚠",
      message: `Proxy enabled at ${args.proxyUrl}, but it is not reachable right now.`,
    };
  }

  return {
    tone: "info",
    icon: "…",
    message: `Checking proxy at ${args.proxyUrl}…`,
  };
}

function buildProxySection(
  settingsStore: SettingsStore,
  registerCleanup?: SettingsCleanupRegistrar,
): HTMLElement {
  const shell = createSectionShell("Proxy", "proxy");

  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-settings-proxy-card";

  let enabled = false;
  let proxyUrl = DEFAULT_LOCAL_PROXY_URL;
  let proxyState: ProxyState = getProxyState();
  let validationError: string | null = null;
  let urlSaveTimer: ReturnType<typeof setTimeout> | null = null;

  const proxyToggle = createToggleRow({
    label: "Proxy",
    sublabel: "Route API calls through a local proxy",
    checked: enabled,
    onChange: (checked) => {
      void saveProxyEnabled(checked);
    },
  });
  proxyToggle.root.classList.add("pi-settings-proxy-toggle");

  const proxyUrlInput = createConfigInput({
    value: proxyUrl,
    placeholder: DEFAULT_LOCAL_PROXY_URL,
  });
  proxyUrlInput.classList.add("pi-settings-proxy-url");
  proxyUrlInput.spellcheck = false;

  const proxyUrlRow = createConfigRow("URL", proxyUrlInput);
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
      showToast("Failed to save proxy setting.");
      return;
    }

    showToast(enabled ? "Proxy enabled." : "Proxy disabled.");
  };

  const saveProxyUrl = async (): Promise<void> => {
    const raw = proxyUrlInput.value.trim();
    const candidate = raw.length > 0 ? raw : DEFAULT_LOCAL_PROXY_URL;

    let normalizedUrl: string;
    try {
      normalizedUrl = validateOfficeProxyUrl(candidate);
    } catch (error: unknown) {
      validationError = error instanceof Error ? error.message : "Invalid proxy URL.";
      updateStatus();
      showToast(`Proxy URL not saved: ${validationError}`);
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
      showToast("Failed to save proxy URL.");
      return;
    }

    proxyUrl = normalizedUrl;
    proxyUrlInput.value = normalizedUrl;
    updateStatus();
    showToast("Proxy URL saved.");
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

    const detail: unknown = event.detail;
    if (!isRecord(detail)) return;

    const state = detail.state;
    if (state === "detected" || state === "not-detected" || state === "unknown") {
      proxyState = state;
      updateStatus();
    }
  };

  document.addEventListener("pi:proxy-state-changed", onProxyStateChanged);
  registerCleanup?.(() => {
    document.removeEventListener("pi:proxy-state-changed", onProxyStateChanged);
  });
  registerCleanup?.(() => {
    flushPendingProxyUrlSave();
  });

  const helper = document.createElement("p");
  helper.className = "pi-overlay-hint pi-settings-proxy-helper";

  const recommendedUrl = document.createElement("code");
  recommendedUrl.textContent = DEFAULT_LOCAL_PROXY_URL;

  const guideLink = document.createElement("a");
  guideLink.href = PROXY_HELPER_DOCS_URL;
  guideLink.target = "_blank";
  guideLink.rel = "noopener noreferrer";
  guideLink.textContent = "Install and setup guide";

  helper.append(
    "Recommended URL: ",
    recommendedUrl,
    ". Keep this on localhost. ",
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
        : DEFAULT_LOCAL_PROXY_URL;
    } catch {
      enabled = false;
      proxyUrl = DEFAULT_LOCAL_PROXY_URL;
    }

    proxyToggle.input.checked = enabled;
    proxyUrlInput.value = proxyUrl;
    updateStatus();
  })();

  card.append(proxyToggle.root, proxyUrlRow, statusHost, helper);
  shell.content.appendChild(card);
  return shell.section;
}

function buildExecutionModeSection(registerCleanup?: SettingsCleanupRegistrar): HTMLElement {
  const shell = createSectionShell("Execution mode", "execution-mode");

  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-settings-execution-card";

  const autoModeToggle = createToggleRow({
    label: "Auto mode",
    sublabel: "Pi applies changes immediately without asking",
  });

  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint pi-settings-execution-hint";
  hint.textContent = "When off, Pi asks before each change (Confirm mode).";

  card.append(autoModeToggle.root, hint);
  shell.content.appendChild(card);

  const getExecutionMode = dependencies.getExecutionMode;
  const setExecutionMode = dependencies.setExecutionMode;

  if (!getExecutionMode || !setExecutionMode) {
    autoModeToggle.input.disabled = true;
    return shell.section;
  }

  let currentMode = getExecutionMode();
  autoModeToggle.input.checked = currentMode === "yolo";

  autoModeToggle.input.addEventListener("change", () => {
    const nextMode: ExecutionMode = autoModeToggle.input.checked ? "yolo" : "safe";
    if (nextMode === currentMode) {
      return;
    }

    autoModeToggle.input.disabled = true;

    void setExecutionMode(nextMode).then(
      () => {
        currentMode = nextMode;
        showToast(nextMode === "yolo" ? "Auto mode." : "Confirm mode.");
      },
      () => {
        autoModeToggle.input.checked = currentMode === "yolo";
        showToast("Couldn't update execution mode.");
      },
    ).finally(() => {
      autoModeToggle.input.disabled = false;
    });
  });

  const onExecutionModeChanged = (event: Event): void => {
    if (!(event instanceof CustomEvent)) {
      return;
    }

    const detail: unknown = event.detail;
    if (!isRecord(detail)) {
      return;
    }

    const mode = detail.mode;
    if (mode !== "yolo" && mode !== "safe") {
      return;
    }

    currentMode = mode;
    autoModeToggle.input.checked = mode === "yolo";
  };

  document.addEventListener(PI_EXECUTION_MODE_CHANGED_EVENT, onExecutionModeChanged);
  registerCleanup?.(() => {
    document.removeEventListener(PI_EXECUTION_MODE_CHANGED_EVENT, onExecutionModeChanged);
  });

  return shell.section;
}

function buildMoreSection(registerCleanup?: SettingsCleanupRegistrar): HTMLElement {
  const panel = document.createElement("div");
  panel.className = "pi-settings-more";

  const advanced = createSectionShell(
    "Advanced",
    "advanced",
    "Power-user shortcuts for rules, backups, and keyboard shortcuts.",
  );

  const modelSwitchCard = document.createElement("div");
  modelSwitchCard.className = "pi-overlay-surface pi-settings-model-switch-card";

  const modelSwitchToggle = createToggleRow({
    label: "Fork model switch into new tab",
    sublabel: "For non-empty chats, open a cloned tab instead of switching this tab in place.",
  });

  const modelSwitchHint = document.createElement("p");
  modelSwitchHint.className = "pi-overlay-hint pi-settings-model-switch-hint";
  modelSwitchHint.textContent = "Default: switch in place (pi-mono parity).";

  modelSwitchCard.append(modelSwitchToggle.root, modelSwitchHint);
  advanced.content.appendChild(modelSwitchCard);

  const getModelSwitchBehavior = dependencies.getModelSwitchBehavior;
  const setModelSwitchBehavior = dependencies.setModelSwitchBehavior;

  if (!getModelSwitchBehavior || !setModelSwitchBehavior) {
    modelSwitchToggle.input.disabled = true;
  } else {
    let currentBehavior = getModelSwitchBehavior();
    modelSwitchToggle.input.checked = currentBehavior === "fork";

    modelSwitchToggle.input.addEventListener("change", () => {
      const nextBehavior: ModelSwitchBehavior = modelSwitchToggle.input.checked ? "fork" : "inPlace";
      if (nextBehavior === currentBehavior) {
        return;
      }

      modelSwitchToggle.input.disabled = true;

      void setModelSwitchBehavior(nextBehavior).then(
        () => {
          currentBehavior = nextBehavior;
          showToast(
            nextBehavior === "fork"
              ? "Model switch will open a new tab for non-empty chats."
              : "Model switch will stay in the current tab.",
          );
        },
        () => {
          modelSwitchToggle.input.checked = currentBehavior === "fork";
          showToast("Couldn't update model switch setting.");
        },
      ).finally(() => {
        modelSwitchToggle.input.disabled = false;
      });
    });
  }

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

  panel.append(buildExecutionModeSection(registerCleanup), advanced.section, experimental.section);
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
      buildProxySection(appStorage.settings, dialog.addCleanup),
      await buildProvidersSection(),
      await buildCustomGatewaySection({
        onProvidersChanged: () => {
          document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        },
      }),
    );

    const morePanel = document.createElement("div");
    morePanel.className = "pi-settings-panel";
    morePanel.dataset.settingsPanel = "more";
    morePanel.appendChild(buildMoreSection(dialog.addCleanup));

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
