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
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
} from "../../tools/experimental-tool-gates.js";
import { isExperimentalFeatureEnabled } from "../../experiments/flags.js";
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
  createButton,
  createToggle,
} from "../../ui/extensions-hub-components.js";
import type { ExtensionsHubDependencies } from "./extensions-hub-overlay.js";

type SettingsStore = IntegrationSettingsStore & WebSearchConfigStore & McpConfigStore & {
  delete?: (key: string) => Promise<void>;
};

// ── Helpers ─────────────────────────────────────────

function normalizeProvider(value: string): WebSearchProvider {
  if (value === "jina" || value === "serper" || value === "tavily" || value === "brave") return value;
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
    return `Workbook (${workbookLabel})`;
  }

  if (sessionEnabled) {
    return hasWorkbook ? "Session only" : "Session";
  }

  return hasWorkbook ? "Off in all scopes" : "Off";
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
    label: "External tools",
    sublabel: "Allow Pi to search the web and call external services",
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
  container.appendChild(createSectionHeader({ label: "Web search" }));

  const webBadgeText = !webSearchEnabled
    ? "Off"
    : apiKey
      ? "Connected"
      : (isApiKeyRequired(selectedProvider) ? "No API key" : "Ready");
  const webBadgeTone = !webSearchEnabled
    ? "muted"
    : (apiKey || !isApiKeyRequired(selectedProvider) ? "ok" : "warn");

  const webCard = createItemCard({
    icon: "🔍",
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

  webCard.body.appendChild(createConfigRow("Provider", providerSelect));

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
  apiKeyLabel.textContent = "API key";

  const apiKeyControls = document.createElement("div");
  apiKeyControls.className = "pi-hub-inline-row";

  const validateBtn = createButton("Validate", {
    compact: true,
    onClick: () => {
      if (isBusy()) return;
      const key = apiKeyInput.value.trim();
      void (async () => {
        try {
          const config = await loadWebSearchProviderConfig(settings);
          const testKey = key.length > 0 ? key : (getApiKeyForProvider(config) ?? "");
          if (!testKey) { showToast("No API key to validate."); return; }
          const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
          const result = await validateWebSearchApiKey({ provider: selectedProvider, apiKey: testKey, proxyBaseUrl });
          showToast(result.ok ? `✓ ${result.message}` : `✗ ${result.message}`);
        } catch (err: unknown) {
          showToast(`✗ ${err instanceof Error ? err.message : String(err)}`);
        }
      })();
    },
  });

  const saveKeyBtn = createButton("Save", {
    primary: true,
    compact: true,
    onClick: () => {
      const key = apiKeyInput.value.trim();
      if (!key) { showToast("Enter an API key first."); return; }
      void runMutation(
        () => saveWebSearchApiKey(settings, selectedProvider, key),
        "config",
        `Saved ${providerInfo.apiKeyLabel}`,
      );
    },
  });

  const clearKeyBtn = createButton("Clear", {
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
  webCard.body.appendChild(createConfigRow("Availability", availability));

  const scopeDetails = document.createElement("details");
  scopeDetails.className = "pi-hub-advanced-disclosure pi-hub-scope-disclosure";
  if (!webSearchEnabled) {
    scopeDetails.open = true;
  }

  const scopeSummary = document.createElement("summary");
  scopeSummary.className = "pi-hub-advanced-summary";
  scopeSummary.textContent = "Scope controls";

  const scopeBody = document.createElement("div");
  scopeBody.className = "pi-hub-advanced-body";

  const sessionToggleRow = createToggleRow({
    label: "Enable for this session",
    checked: webSearchSessionEnabled,
    onChange: (checked) => {
      if (!sessionId) {
        showToast("No active session");
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
      ? `Enable for workbook (${workbookContext.workbookLabel})`
      : "Workbook scope unavailable",
    checked: webSearchWorkbookEnabled,
    onChange: (checked) => {
      if (!workbookId) {
        showToast("Workbook scope unavailable");
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

  // ── MCP servers section ───────────────────────
  const mcpAddForm = createAddForm();
  const mcpAddVisible = { value: false };

  const mcpHeader = createSectionHeader({
    label: "MCP servers",
    actionLabel: "+ Add server",
    onAction: () => {
      mcpAddVisible.value = !mcpAddVisible.value;
      mcpAddForm.hidden = !mcpAddVisible.value;
    },
  });
  container.appendChild(mcpHeader);

  const mcpList = document.createElement("div");
  mcpList.className = "pi-hub-stack";

  if (mcpServers.length === 0) {
    mcpList.appendChild(createEmptyInline("⚡", "No MCP servers configured.\nAdd one to connect external tools."));
  } else {
    for (const server of mcpServers) {
      mcpList.appendChild(renderMcpServerCard(server, settings, isBusy, runMutation));
    }
  }
  container.appendChild(mcpList);

  // MCP add form (hidden by default)
  const nameInput = createAddFormInput("Server name");
  const urlInput = createAddFormInput("https://server-url/rpc");
  const tokenInput = createAddFormInput("Bearer token (optional)");
  tokenInput.type = "password";

  const addRow = createAddFormRow();
  addRow.append(nameInput, urlInput);

  const tokenRow = createAddFormRow();
  tokenRow.append(tokenInput, createButton("Add", {
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
  // Python bridge is always available; tmux bridge requires the experimental flag
  const showPython = true;
  const showTmux = isExperimentalFeatureEnabled("tmux_bridge") || tmuxUrl.length > 0;

  if (showPython || showTmux) {
    container.appendChild(createSectionHeader({ label: "Bridges" }));

    const bridgeList = document.createElement("div");
    bridgeList.className = "pi-hub-stack";

    if (showPython) {
      bridgeList.appendChild(renderBridgeCard({
        icon: "🐍",
        name: "Python bridge",
        description: "Execute Python code in a local environment",
        settingKey: PYTHON_BRIDGE_URL_SETTING_KEY,
        placeholder: "http://localhost:3340",
        currentUrl: pythonUrl,
        settings,
        isBusy,
        runMutation,
      }));
    }

    if (showTmux) {
      bridgeList.appendChild(renderBridgeCard({
        icon: "🖥",
        name: "tmux bridge",
        description: "Remote shell sessions via tmux",
        settingKey: TMUX_BRIDGE_URL_SETTING_KEY,
        placeholder: "http://localhost:3341",
        currentUrl: tmuxUrl,
        settings,
        isBusy,
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
  const toolLabel = server.enabled ? "Enabled" : "Disabled";
  const card = createItemCard({
    icon: "⚡",
    iconColor: "blue",
    name: server.name,
    meta: server.url,
    expandable: true,
    badges: [getStatusBadge(server.enabled, toolLabel)],
  });

  // URL
  card.body.appendChild(createConfigRow("URL", createConfigValue(server.url)));

  // Token
  const tokenValue = server.token ? maskSecret(server.token) : "(none)";
  card.body.appendChild(createConfigRow("Token", createConfigValue(tokenValue)));

  // Enabled toggle
  const enabledRow = document.createElement("div");
  enabledRow.className = "pi-item-card__config-row";
  const enabledLabel = document.createElement("span");
  enabledLabel.className = "pi-item-card__config-label";
  enabledLabel.textContent = "Enabled";
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
  const testBtn = createButton("Test", {
    compact: true,
    onClick: () => {
      if (isBusy()) return;
      void (async () => {
        try {
          const result = await probeMcpServer(server, settings);
          const transport = result.proxied ? "proxy" : "direct";
          showToast(`${server.name}: reachable (${result.toolCount} tool${result.toolCount === 1 ? "" : "s"}, ${transport})`);
        } catch (err: unknown) {
          showToast(`${server.name}: ${err instanceof Error ? err.message : String(err)}`);
        }
      })();
    },
  });

  const removeBtn = createButton("Remove", {
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
  icon: string;
  name: string;
  description: string;
  settingKey: string;
  placeholder: string;
  currentUrl: string;
  settings: SettingsStore;
  isBusy: () => boolean;
  runMutation: (action: () => Promise<void>, reason: "toggle" | "scope" | "external-toggle" | "config", msg?: string) => Promise<void>;
}): HTMLElement {
  const connected = args.currentUrl.length > 0;
  const card = createItemCard({
    icon: args.icon,
    iconColor: "amber",
    name: args.name,
    description: args.description,
    expandable: true,
    badges: [connected
      ? { text: "Configured", tone: "ok" as const }
      : { text: "Not connected", tone: "muted" as const },
    ],
  });

  const urlInput = createConfigInput({
    value: args.currentUrl,
    placeholder: args.placeholder,
  });
  card.body.appendChild(createConfigRow("Bridge URL", urlInput));

  const saveBridgeUrl = (clear: boolean): void => {
    const raw = clear ? "" : urlInput.value.trim();

    if (raw.length > 0) {
      try {
        validateOfficeProxyUrl(raw);
      } catch (err: unknown) {
        showToast(`Invalid URL: ${err instanceof Error ? err.message : String(err)}`);
        return;
      }
    }

    void args.runMutation(async () => {
      if (raw.length === 0) {
        if (typeof args.settings.delete === "function") {
          await args.settings.delete(args.settingKey);
        } else {
          await args.settings.set(args.settingKey, "");
        }
      } else {
        await args.settings.set(args.settingKey, raw);
      }
      dispatchExperimentalToolConfigChanged({ configKey: args.settingKey });
    }, "config", raw.length > 0 ? `${args.name} URL saved` : `${args.name} URL cleared`);
  };

  const saveBtn = createButton("Save", { compact: true, onClick: () => saveBridgeUrl(false) });
  const clearBtn = createButton("Clear", { compact: true, onClick: () => saveBridgeUrl(true) });
  card.body.appendChild(createActionsRow(saveBtn, clearBtn));

  return card.root;
}
