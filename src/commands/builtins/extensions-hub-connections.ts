/**
 * Extensions hub — Connections tab.
 *
 * External tools master toggle, web search config, MCP server management,
 * and bridge URLs.
 */

import { INTEGRATION_IDS } from "../../integrations/catalog.js";
import type { IntegrationSettingsStore } from "../../integrations/store.js";
import type { WebSearchConfigStore } from "../../tools/web-search-config.js";
import type { McpConfigStore, McpServerConfig } from "../../tools/mcp-config.js";
import {
  getExternalToolsEnabled,
  getSessionIntegrationIds,
  getWorkbookIntegrationIds,
  setExternalToolsEnabled,
  setIntegrationEnabledInScope,
} from "../../integrations/store.js";
import {
  checkApiKeyFormat,
  getApiKeyForProvider,
  isApiKeyRequired,
  loadWebSearchProviderConfig,
  maskSecret,
  saveWebSearchApiKey,
  saveWebSearchProvider,
  clearWebSearchApiKey,
  WEB_SEARCH_PROVIDER_INFO,
  type WebSearchProvider,
} from "../../tools/web-search-config.js";
import { validateWebSearchApiKey } from "../../tools/web-search.js";
import {
  createMcpServerConfig,
  loadMcpServers,
  saveMcpServers,
} from "../../tools/mcp-config.js";
import { getEnabledProxyBaseUrl } from "../../tools/external-fetch.js";
import { validateOfficeProxyUrl } from "../../auth/proxy-validation.js";
import { dispatchExperimentalToolConfigChanged } from "../../experiments/events.js";
import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
} from "../../tools/experimental-tool-gates.js";
import { probeMcpServer } from "./extensions-hub-mcp-probe.js";
import { showToast } from "../../ui/toast.js";
import {
  createToggleRow,
  createSectionHeader,
  createItemCard,
  createConfigRow,
  createConfigInput,
  createConfigValue,
  createAddForm,
  createAddFormRow,
  createAddFormInput,
  createEmptyInline,
  createActionsRow,
  type IconContent,
  createButton,
  createToggle,
} from "../../ui/extensions-hub-components.js";
import { lucide, Search, Terminal, Zap } from "../../ui/lucide-icons.js";
import { t } from "../../language/index.js";
import type { ExtensionsHubDependencies } from "./extensions-hub-overlay.js";
import { renderExtensionConnectionsSection } from "./extensions-hub-extension-connections.js";

type SettingsStore = IntegrationSettingsStore & WebSearchConfigStore & McpConfigStore & {
  delete?: (key: string) => Promise<void>;
};

// ── Helpers ─────────────────────────────────────────

function normalizeProvider(value: string): WebSearchProvider {
  if (value === "jina" || value === "firecrawl" || value === "serper" || value === "tavily" || value === "brave") return value;
  return "jina";
}

function getStatusBadge(ok: boolean, label: string): { text: string; tone: "ok" | "warn" | "muted" } {
  return ok ? { text: label, tone: "ok" } : { text: label, tone: "muted" };
}

function describeWebSearchAvailability(args: {
  sessionEnabled: boolean;
  workbookEnabled: boolean;
  workbookLabel: string;
  hasWorkbook: boolean;
}): string {
  const { sessionEnabled, workbookEnabled, workbookLabel, hasWorkbook } = args;

  if (sessionEnabled && workbookEnabled && hasWorkbook) {
    return `Session + workbook (${workbookLabel})`;
  }

  if (workbookEnabled && hasWorkbook) {
    return t("ext-hub-connections.scopeWorkbook", { label: workbookLabel });
  }

  if (sessionEnabled) {
    return hasWorkbook ? t("ext-hub-connections.scopeSessionOnly") : t("ext-hub-connections.scopeSession");
  }

  return hasWorkbook ? t("ext-hub-connections.scopeOff") : t("ext-hub-connections.scopeOffShort");
}

// Resolved lazily — t() must not run at module scope (language set at boot).
function bridgeSetupHint(): string {
  return t("ext-hub-connections.setupHint");
}

function selectElementText(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) return;

  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
}

function createBridgeSetupCommand(command: string): HTMLDivElement {
  const setup = document.createElement("div");
  setup.className = "pi-hub-bridge-setup";

  const commandRow = document.createElement("div");
  commandRow.className = "pi-hub-bridge-setup__command";

  const code = document.createElement("code");
  code.className = "pi-hub-bridge-setup__code";
  code.textContent = command;

  const copyButton = document.createElement("button");
  copyButton.type = "button";
  copyButton.className = "pi-hub-bridge-setup__copy";
  copyButton.textContent = "📋";
  copyButton.title = t("bridge-setup.copyCommandTitle");
  copyButton.addEventListener("click", () => {
    if (!navigator.clipboard?.writeText) {
      selectElementText(code);
      return;
    }

    void navigator.clipboard.writeText(command).then(
      () => {
        copyButton.textContent = "✓";
        setTimeout(() => {
          copyButton.textContent = "📋";
        }, 1400);
      },
      () => {
        selectElementText(code);
      },
    );
  });

  const hint = document.createElement("p");
  hint.className = "pi-hub-bridge-setup__hint";
  hint.textContent = bridgeSetupHint();

  commandRow.append(code, copyButton);
  setup.append(commandRow, hint);

  return setup;
}

// ── Main render ─────────────────────────────────────

export async function renderConnectionsTab(args: {
  container: HTMLElement;
  settings: SettingsStore;
  deps: ExtensionsHubDependencies;
  isBusy: () => boolean;
  runMutation: (action: () => Promise<void>, reason: "toggle" | "scope" | "external-toggle" | "config", msg?: string) => Promise<void>;
}): Promise<void> {
  const { container, settings, deps, isBusy, runMutation } = args;

  const sessionId = deps.getActiveSessionId();
  const workbookContext = await deps.resolveWorkbookContext();
  const workbookId = workbookContext.workbookId;

  // Load state
  const [
    externalEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    webSearchConfig,
    mcpServers,
    pythonUrlRaw,
    tmuxUrlRaw,
  ] = await Promise.all([
    getExternalToolsEnabled(settings),
    sessionId
      ? getSessionIntegrationIds(settings, sessionId, INTEGRATION_IDS, {
        applyDefaultsWhenUnconfigured: workbookId === null,
      })
      : Promise.resolve<string[]>([]),
    workbookId
      ? getWorkbookIntegrationIds(settings, workbookId, INTEGRATION_IDS)
      : Promise.resolve<string[]>([]),
    loadWebSearchProviderConfig(settings),
    loadMcpServers(settings),
    settings.get(PYTHON_BRIDGE_URL_SETTING_KEY),
    settings.get(TMUX_BRIDGE_URL_SETTING_KEY),
  ]);

  const pythonUrl = typeof pythonUrlRaw === "string" ? pythonUrlRaw.trim() : "";
  const tmuxUrl = typeof tmuxUrlRaw === "string" ? tmuxUrlRaw.trim() : "";
  const effectivePythonUrl = pythonUrl.length > 0 ? pythonUrl : DEFAULT_PYTHON_BRIDGE_URL;
  const effectiveTmuxUrl = tmuxUrl.length > 0 ? tmuxUrl : DEFAULT_TMUX_BRIDGE_URL;
  const selectedProvider = webSearchConfig.provider;
  const providerInfo = WEB_SEARCH_PROVIDER_INFO[selectedProvider];
  const apiKey = getApiKeyForProvider(webSearchConfig);
  const webSearchSessionEnabled = sessionIntegrationIds.includes("web_search");
  const webSearchWorkbookEnabled = workbookIntegrationIds.includes("web_search");
  const webSearchEnabled = webSearchSessionEnabled || webSearchWorkbookEnabled;

  container.replaceChildren();

  // ── Master toggle ─────────────────────────────
  const surface = document.createElement("div");
  surface.className = "pi-overlay-surface";

  const masterToggle = createToggleRow({
    label: t("extensions-hub-connections.externalTools"),
    sublabel: t("ext-hub-connections.allowExternal"),
    checked: externalEnabled,
    onChange: (checked) => {
      void runMutation(
        () => setExternalToolsEnabled(settings, checked),
        "external-toggle",
        `External tools ${checked ? "enabled" : "disabled"}`,
      );
    },
  });
  surface.appendChild(masterToggle.root);
  container.appendChild(surface);

  // ── Web search section ────────────────────────
  container.appendChild(createSectionHeader({ label: t("extensions-hub-connections.webSearch") }));

  const webBadgeText = !webSearchEnabled
    ? t("extensions-hub-connections.webSearchOff")
    : apiKey
      ? t("extensions-hub-connections.webSearchConnected")
      : (isApiKeyRequired(selectedProvider) ? t("extensions-hub-connections.noApiKey") : t("extensions-hub-connections.ready"));
  const webBadgeTone = !webSearchEnabled
    ? "muted"
    : (apiKey || !isApiKeyRequired(selectedProvider) ? "ok" : "warn");

  const webCard = createItemCard({
    icon: lucide(Search),
    iconColor: "green",
    name: providerInfo.title,
    description: providerInfo.shortDescription,
    expandable: true,
    badges: [{ text: webBadgeText, tone: webBadgeTone }],
  });

  // Provider picker
  const providerSelect = document.createElement("select");
  providerSelect.className = "pi-item-card__config-input pi-item-card__config-select";
  for (const [key, info] of Object.entries(WEB_SEARCH_PROVIDER_INFO)) {
    const option = document.createElement("option");
    option.value = key;
    option.textContent = info.title;
    if (key === selectedProvider) option.selected = true;
    providerSelect.appendChild(option);
  }
  providerSelect.addEventListener("change", () => {
    const provider = normalizeProvider(providerSelect.value);
    void runMutation(
      () => saveWebSearchProvider(settings, provider),
      "config",
      `Web search provider set to ${WEB_SEARCH_PROVIDER_INFO[provider].title}`,
    );
  });

  webCard.body.appendChild(createConfigRow(t("extensions-hub-connections.providerLabel"), providerSelect));

  // API key row
  const apiKeyInput = createConfigInput({
    placeholder: providerInfo.apiKeyLabel,
    type: "password",
    value: apiKey ? maskSecret(apiKey) : "",
  });

  const apiKeyRow = document.createElement("div");
  apiKeyRow.className = "pi-item-card__config-row";

  const apiKeyLabel = document.createElement("span");
  apiKeyLabel.className = "pi-item-card__config-label";
  apiKeyLabel.textContent = t("extensions-hub-connections.apiKeyLabel");

  const apiKeyControls = document.createElement("div");
  apiKeyControls.className = "pi-hub-inline-row";

  const validateBtn = createButton(t("extensions-hub-connections.validateButton"), {
    compact: true,
    onClick: () => {
      if (isBusy()) return;
      const key = apiKeyInput.value.trim();
      void (async () => {
        try {
          const config = await loadWebSearchProviderConfig(settings);
          const testKey = key.length > 0 ? key : (getApiKeyForProvider(config) ?? "");
          if (!testKey) { showToast(t("extensions-hub-connections.toast.noApiKeyToValidate")); return; }
          const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
          const result = await validateWebSearchApiKey({ provider: selectedProvider, apiKey: testKey, proxyBaseUrl });
          showToast(t(result.ok ? "extensions-hub-connections.toast.validationOk" : "extensions-hub-connections.toast.validationFailed", { message: result.message }));
        } catch (err) {
          showToast(t("extensions-hub-connections.toast.validationError", { error: err instanceof Error ? err.message : String(err) }));
        }
      })();
    },
  });

  const saveKeyBtn = createButton(t("extensions-hub-connections.saveButton"), {
    primary: true,
    compact: true,
    onClick: () => {
      const key = apiKeyInput.value.trim();
      if (!key) { showToast(t("extensions-hub-connections.toast.enterApiKey")); return; }
      const formatWarning = checkApiKeyFormat(selectedProvider, key);
      void runMutation(
        () => saveWebSearchApiKey(settings, selectedProvider, key),
        "config",
        formatWarning
          ? `⚠️ ${formatWarning} Key saved anyway — use Validate to test it.`
          : `Saved ${providerInfo.apiKeyLabel}`,
      );
    },
  });

  const clearKeyBtn = createButton(t("extensions-hub-connections.clearButton"), {
    compact: true,
    onClick: () => {
      void runMutation(
        () => clearWebSearchApiKey(settings, selectedProvider),
        "config",
        `Cleared ${providerInfo.apiKeyLabel}`,
      );
    },
  });

  apiKeyControls.append(apiKeyInput, validateBtn, saveKeyBtn, clearKeyBtn);
  apiKeyRow.append(apiKeyLabel, apiKeyControls);
  webCard.body.appendChild(apiKeyRow);

  const availability = createConfigValue(describeWebSearchAvailability({
    sessionEnabled: webSearchSessionEnabled,
    workbookEnabled: webSearchWorkbookEnabled,
    workbookLabel: workbookContext.workbookLabel,
    hasWorkbook: workbookId !== null,
  }));
  webCard.body.appendChild(createConfigRow(t("extensions-hub-connections.availability"), availability));

  const scopeDetails = document.createElement("details");
  scopeDetails.className = "pi-hub-advanced-disclosure pi-hub-scope-disclosure";
  if (!webSearchEnabled) {
    scopeDetails.open = true;
  }

  const scopeSummary = document.createElement("summary");
  scopeSummary.className = "pi-hub-advanced-summary";
  scopeSummary.textContent = t("extensions-hub-connections.scope-controls");

  const scopeBody = document.createElement("div");
  scopeBody.className = "pi-hub-advanced-body";

  const sessionToggleRow = createToggleRow({
    label: t("ext-hub-connections.enableSession"),
    checked: webSearchSessionEnabled,
    onChange: (checked) => {
      if (!sessionId) {
        showToast(t("extensions-hub-connections.toast.noActiveSession"));
        return;
      }
      void runMutation(async () => {
        await setIntegrationEnabledInScope({
          settings,
          scope: "session",
          identifier: sessionId,
          integrationId: "web_search",
          enabled: checked,
          knownIntegrationIds: INTEGRATION_IDS,
        });
      }, "scope", `Web search ${checked ? "enabled" : "disabled"} for this session`);
    },
  });
  sessionToggleRow.input.disabled = isBusy() || !sessionId;
  scopeBody.appendChild(sessionToggleRow.root);

  const workbookToggleRow = createToggleRow({
    label: workbookId
      ? t("ext-hub-connections.enableWorkbook", { label: workbookContext.workbookLabel })
      : t("ext-hub-connections.scopeUnavailable"),
    checked: webSearchWorkbookEnabled,
    onChange: (checked) => {
      if (!workbookId) {
        showToast(t("extensions-hub-connections.toast.workbookScopeUnavailable"));
        return;
      }
      void runMutation(async () => {
        await setIntegrationEnabledInScope({
          settings,
          scope: "workbook",
          identifier: workbookId,
          integrationId: "web_search",
          enabled: checked,
          knownIntegrationIds: INTEGRATION_IDS,
        });
      }, "scope", `Web search ${checked ? "enabled" : "disabled"} for this workbook`);
    },
  });
  workbookToggleRow.input.disabled = isBusy() || !workbookId;
  scopeBody.appendChild(workbookToggleRow.root);

  scopeDetails.append(scopeSummary, scopeBody);
  webCard.body.appendChild(scopeDetails);

  container.appendChild(webCard.root);

  // ── Extension connections section ─────────────
  await renderExtensionConnectionsSection({
    container,
    connectionManager: deps.connectionManager,
    extensionManager: deps.extensionManager,
  });

  // ── MCP servers section ───────────────────────
  const mcpAddForm = createAddForm();
  const mcpAddVisible = { value: false };

  const mcpHeader = createSectionHeader({
    label: t("extensions-hub-connections.mcpSection"),
    actionLabel: t("extensions-hub-connections.addServer"),
    onAction: () => {
      mcpAddVisible.value = !mcpAddVisible.value;
      mcpAddForm.hidden = !mcpAddVisible.value;
    },
  });
  container.appendChild(mcpHeader);

  const mcpList = document.createElement("div");
  mcpList.className = "pi-hub-stack";

  if (mcpServers.length === 0) {
    mcpList.appendChild(createEmptyInline(lucide(Zap), t("ext-hub-connections.noMcpServers")));
  } else {
    for (const server of mcpServers) {
      mcpList.appendChild(renderMcpServerCard(server, settings, isBusy, runMutation));
    }
  }
  container.appendChild(mcpList);

  // MCP add form (hidden by default)
  const nameInput = createAddFormInput(t("ext-hub-connections.serverNamePlaceholder"));
  const urlInput = createAddFormInput(t("ext-hub-connections.serverUrlPlaceholder"));
  const tokenInput = createAddFormInput(t("ext-hub-connections.bearerTokenPlaceholder"));
  tokenInput.type = "password";

  const addRow = createAddFormRow();
  addRow.append(nameInput, urlInput);

  const tokenRow = createAddFormRow();
  tokenRow.append(tokenInput, createButton(t("ext-hub-connections.addButton"), {
    primary: true,
    compact: true,
    onClick: () => {
      void runMutation(async () => {
        const servers = await loadMcpServers(settings);
        const next = createMcpServerConfig({
          name: nameInput.value,
          url: urlInput.value,
          token: tokenInput.value,
          enabled: true,
        });
        await saveMcpServers(settings, [...servers, next]);
        nameInput.value = "";
        urlInput.value = "";
        tokenInput.value = "";
      }, "config", "Added MCP server");
    },
  }));

  mcpAddForm.append(addRow, tokenRow);
  mcpAddForm.hidden = true;
  container.appendChild(mcpAddForm);

  // ── Bridges section ───────────────────────────
  const showPython = true;
  const showTmux = true;

  if (showPython || showTmux) {
    container.appendChild(createSectionHeader({ label: t("extensions-hub-connections.bridgesSection") }));

    const bridgeList = document.createElement("div");
    bridgeList.className = "pi-hub-stack";

    if (showPython) {
      bridgeList.appendChild(renderBridgeCard({
        icon: lucide(Terminal),
        name: t("ext-hub-connections.pythonName"),
        description: t("ext-hub-connections.pythonDesc"),
        settingKey: PYTHON_BRIDGE_URL_SETTING_KEY,
        setupCommand: "npx pi-for-excel-python-bridge",
        defaultUrl: DEFAULT_PYTHON_BRIDGE_URL,
        placeholder: DEFAULT_PYTHON_BRIDGE_URL,
        currentUrl: effectivePythonUrl,
        hasCustomUrl: pythonUrl.length > 0,
        settings,
        runMutation,
      }));
    }

    if (showTmux) {
      bridgeList.appendChild(renderBridgeCard({
        icon: lucide(Terminal),
        name: t("ext-hub-connections.tmuxName"),
        description: t("ext-hub-connections.tmuxDesc"),
        settingKey: TMUX_BRIDGE_URL_SETTING_KEY,
        setupCommand: "npx pi-for-excel-tmux-bridge",
        defaultUrl: DEFAULT_TMUX_BRIDGE_URL,
        placeholder: DEFAULT_TMUX_BRIDGE_URL,
        currentUrl: effectiveTmuxUrl,
        hasCustomUrl: tmuxUrl.length > 0,
        settings,
        runMutation,
      }));
    }

    container.appendChild(bridgeList);
  }
}

// ── MCP server card ─────────────────────────────────

function renderMcpServerCard(
  server: McpServerConfig,
  settings: SettingsStore,
  isBusy: () => boolean,
  runMutation: (action: () => Promise<void>, reason: "toggle" | "scope" | "external-toggle" | "config", msg?: string) => Promise<void>,
): HTMLElement {
  const toolLabel = server.enabled ? t("ext-hub-connections.badgeEnabled") : t("ext-hub-connections.badgeDisabled");
  const card = createItemCard({
    icon: lucide(Zap),
    iconColor: "blue",
    name: server.name,
    meta: server.url,
    expandable: true,
    badges: [getStatusBadge(server.enabled, toolLabel)],
  });

  // URL
  card.body.appendChild(createConfigRow(t("extensions-hub-connections.url"), createConfigValue(server.url)));

  // Token
  const tokenValue = server.token ? maskSecret(server.token) : t("ext-hub-connections.badgeNoToken");
  card.body.appendChild(createConfigRow(t("extensions-hub-connections.token"), createConfigValue(tokenValue)));

  // Enabled toggle
  const enabledRow = document.createElement("div");
  enabledRow.className = "pi-item-card__config-row";
  const enabledLabel = document.createElement("span");
  enabledLabel.className = "pi-item-card__config-label";
  enabledLabel.textContent = t("extensions-hub-connections.enabled");
  const enabledToggle = createToggle({
    checked: server.enabled,
    onChange: (checked) => {
      void runMutation(async () => {
        const servers = await loadMcpServers(settings);
        const updated = servers.map((s) =>
          s.id === server.id ? { ...s, enabled: checked } : s,
        );
        await saveMcpServers(settings, updated);
      }, "config", `${server.name}: ${checked ? "enabled" : "disabled"}`);
    },
  });
  enabledRow.append(enabledLabel, enabledToggle.root);
  card.body.appendChild(enabledRow);

  // Actions
  const testBtn = createButton(t("ext-hub-connections.testButton"), {
    compact: true,
    onClick: () => {
      if (isBusy()) return;
      void (async () => {
        try {
          const result = await probeMcpServer(server, settings);
          const transport = result.proxied
            ? t("extensions-hub-connections.transport.proxy")
            : t("extensions-hub-connections.transport.direct");
          showToast(t("extensions-hub-connections.toast.serverReachable", { name: server.name, count: result.toolCount, plural: result.toolCount === 1 ? "" : "s", transport }));
        } catch (err) {
          showToast(t("extensions-hub-connections.toast.serverError", { name: server.name, error: err instanceof Error ? err.message : String(err) }));
        }
      })();
    },
  });

  const removeBtn = createButton(t("ext-hub-connections.removeButton"), {
    danger: true,
    compact: true,
    onClick: () => {
      void runMutation(async () => {
        const servers = await loadMcpServers(settings);
        await saveMcpServers(settings, servers.filter((s) => s.id !== server.id));
      }, "config", `Removed MCP server: ${server.name}`);
    },
  });

  card.body.appendChild(createActionsRow(testBtn, removeBtn));

  return card.root;
}

// ── Bridge card ─────────────────────────────────────

function renderBridgeCard(args: {
  icon: IconContent;
  name: string;
  description: string;
  settingKey: string;
  setupCommand: string;
  defaultUrl: string;
  placeholder: string;
  currentUrl: string;
  hasCustomUrl: boolean;
  settings: SettingsStore;
  runMutation: (action: () => Promise<void>, reason: "toggle" | "scope" | "external-toggle" | "config", msg?: string) => Promise<void>;
}): HTMLElement {
  const card = createItemCard({
    icon: args.icon,
    iconColor: "amber",
    name: args.name,
    description: args.description,
    expandable: true,
    expanded: !args.hasCustomUrl,
    badges: [args.hasCustomUrl
      ? { text: t("ext-hub-connections.configured"), tone: "ok" as const }
      : { text: t("ext-hub-connections.defaultUrl"), tone: "muted" as const },
    ],
  });

  const setupLabel = document.createElement("p");
  setupLabel.className = "pi-hub-bridge-setup__label";
  setupLabel.textContent = t("extensions-hub-connections.quick-setup");
  card.body.append(setupLabel, createBridgeSetupCommand(args.setupCommand));

  const urlInput = createConfigInput({
    value: args.currentUrl,
    placeholder: args.placeholder,
  });
  card.body.appendChild(createConfigRow(t("ext-hub-connections.bridgeUrl"), urlInput));

  const saveBridgeUrl = (clear: boolean): void => {
    const candidateUrl = clear ? "" : urlInput.value.trim();
    let normalizedCandidateUrl = "";

    if (candidateUrl.length > 0) {
      try {
        normalizedCandidateUrl = validateOfficeProxyUrl(candidateUrl);
      } catch (err) {
        showToast(t("ext-hub-connections.toast.invalidUrl", { error: err instanceof Error ? err.message : String(err) }));
        return;
      }
    }

    const useDefaultUrl = normalizedCandidateUrl.length === 0 || normalizedCandidateUrl === args.defaultUrl;

    void args.runMutation(async () => {
      if (useDefaultUrl) {
        if (typeof args.settings.delete === "function") {
          await args.settings.delete(args.settingKey);
        } else {
          await args.settings.set(args.settingKey, "");
        }
      } else {
        await args.settings.set(args.settingKey, normalizedCandidateUrl);
      }
      dispatchExperimentalToolConfigChanged({ configKey: args.settingKey });
    }, "config", useDefaultUrl ? `${args.name} URL set to default` : `${args.name} URL saved`);
  };

  const saveBtn = createButton(t("ext-hub-connections.saveButton"), { compact: true, onClick: () => saveBridgeUrl(false) });
  const clearBtn = createButton(t("ext-hub-connections.clearButton"), { compact: true, onClick: () => saveBridgeUrl(true) });
  card.body.appendChild(createActionsRow(saveBtn, clearBtn));

  return card.root;
}
