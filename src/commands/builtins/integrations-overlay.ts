/**
 * Integrations manager overlay.
 */

import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";

import { APP_NAME, APP_VERSION } from "../../app/metadata.js";
import {
  INTEGRATION_IDS,
  listIntegrationDefinitions,
  type IntegrationDefinition,
} from "../../integrations/catalog.js";
import { dispatchIntegrationsChanged } from "../../integrations/events.js";
import {
  INTEGRATIONS_LABEL,
  INTEGRATIONS_LABEL_LOWER,
} from "../../integrations/naming.js";
import {
  getExternalToolsEnabled,
  getSessionIntegrationIds,
  getWorkbookIntegrationIds,
  resolveConfiguredIntegrationIds,
  setExternalToolsEnabled,
  setIntegrationEnabledInScope,
  type IntegrationSettingsStore,
} from "../../integrations/store.js";
import { getEnabledProxyBaseUrl, resolveOutboundRequestUrl } from "../../tools/external-fetch.js";
import {
  clearWebSearchApiKey,
  getApiKeyForProvider,
  loadWebSearchProviderConfig,
  maskSecret,
  saveWebSearchApiKey,
  saveWebSearchProvider,
  WEB_SEARCH_PROVIDERS,
  WEB_SEARCH_PROVIDER_INFO,
  type WebSearchConfigStore,
  type WebSearchProvider,
  type WebSearchProviderConfig,
} from "../../tools/web-search-config.js";
import { validateWebSearchApiKey } from "../../tools/web-search.js";
import {
  createMcpServerConfig,
  loadMcpServers,
  saveMcpServers,
  type McpConfigStore,
  type McpServerConfig,
} from "../../tools/mcp-config.js";
import {
  closeOverlayById,
  createOverlayBadge,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  createOverlayInput,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { INTEGRATIONS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import {
  getHttpErrorReason,
  runWithTimeoutAbort,
} from "../../utils/network.js";
import { isRecord } from "../../utils/type-guards.js";

const MCP_PROBE_TIMEOUT_MS = 8_000;

interface WorkbookContextSnapshot {
  workbookId: string | null;
  workbookLabel: string;
}

export interface IntegrationsDialogDependencies {
  getActiveSessionId: () => string | null;
  resolveWorkbookContext: () => Promise<WorkbookContextSnapshot>;
  onChanged?: () => Promise<void> | void;
}

interface IntegrationsSnapshot {
  sessionId: string;
  workbookId: string | null;
  workbookLabel: string;
  externalToolsEnabled: boolean;
  sessionIntegrationIds: string[];
  workbookIntegrationIds: string[];
  activeIntegrationIds: string[];
  webSearchConfig: WebSearchProviderConfig;
  mcpServers: McpServerConfig[];
}

function getErrorMessage(error: unknown): string {
  if (error instanceof Error && error.message.trim().length > 0) {
    return error.message;
  }
  return String(error);
}

function normalizeWebSearchProvider(value: string): WebSearchProvider {
  if (value === "serper" || value === "tavily" || value === "brave") {
    return value;
  }
  return "serper";
}

function isEnabledInList(integrationIds: readonly string[], integrationId: string): boolean {
  return integrationIds.includes(integrationId);
}

function getSettingsStore(): Promise<
  IntegrationSettingsStore & WebSearchConfigStore & McpConfigStore
> {
  return Promise.resolve(getAppStorage().settings);
}

async function buildSnapshot(
  dependencies: IntegrationsDialogDependencies,
): Promise<IntegrationsSnapshot> {
  const settings = await getSettingsStore();
  const sessionId = dependencies.getActiveSessionId();
  if (!sessionId) {
    throw new Error("No active session.");
  }

  const workbookContext = await dependencies.resolveWorkbookContext();

  const [
    externalToolsEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    activeIntegrationIds,
    webSearchConfig,
    mcpServers,
  ] = await Promise.all([
    getExternalToolsEnabled(settings),
    getSessionIntegrationIds(settings, sessionId, INTEGRATION_IDS),
    workbookContext.workbookId
      ? getWorkbookIntegrationIds(settings, workbookContext.workbookId, INTEGRATION_IDS)
      : Promise.resolve([]),
    resolveConfiguredIntegrationIds({
      settings,
      sessionId,
      workbookId: workbookContext.workbookId,
      knownIntegrationIds: INTEGRATION_IDS,
    }),
    loadWebSearchProviderConfig(settings),
    loadMcpServers(settings),
  ]);

  return {
    sessionId,
    workbookId: workbookContext.workbookId,
    workbookLabel: workbookContext.workbookLabel,
    externalToolsEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    activeIntegrationIds,
    webSearchConfig,
    mcpServers,
  };
}

function parseToolCountFromListResponse(value: unknown): number {
  if (!isRecord(value)) return 0;
  if (!isRecord(value.result)) return 0;
  const tools = value.result.tools;
  return Array.isArray(tools) ? tools.length : 0;
}

async function postJsonRpc(args: {
  server: McpServerConfig;
  method: string;
  params?: unknown;
  settings: IntegrationSettingsStore;
  expectResponse?: boolean;
}): Promise<{ response: unknown; proxied: boolean; proxyBaseUrl?: string } | null> {
  const { server, method, params, settings, expectResponse = true } = args;

  const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
  const resolved = resolveOutboundRequestUrl({
    targetUrl: server.url,
    proxyBaseUrl,
  });

  const body: Record<string, unknown> = {
    jsonrpc: "2.0",
    method,
  };

  if (params !== undefined) {
    body.params = params;
  }

  if (expectResponse) {
    body.id = crypto.randomUUID();
  }

  const headers: Record<string, string> = {
    "Content-Type": "application/json",
    Accept: "application/json",
  };

  if (server.token) {
    headers.Authorization = `Bearer ${server.token}`;
  }

  return runWithTimeoutAbort({
    signal: undefined,
    timeoutMs: MCP_PROBE_TIMEOUT_MS,
    timeoutErrorMessage: `MCP request timed out after ${MCP_PROBE_TIMEOUT_MS}ms.`,
    run: async (requestSignal) => {
      const response = await fetch(resolved.requestUrl, {
        method: "POST",
        headers,
        body: JSON.stringify(body),
        signal: requestSignal,
      });

      if (!response.ok) {
        const text = await response.text();
        const reason = getHttpErrorReason(response.status, text);
        throw new Error(`MCP request failed (${response.status}): ${reason}`);
      }

      if (!expectResponse) {
        return {
          response: null,
          proxied: resolved.proxied,
          proxyBaseUrl: resolved.proxyBaseUrl,
        };
      }

      const text = await response.text();
      const payload: unknown = text.trim().length > 0 ? JSON.parse(text) : null;

      return {
        response: payload,
        proxied: resolved.proxied,
        proxyBaseUrl: resolved.proxyBaseUrl,
      };
    },
  });
}

async function probeMcpServer(
  server: McpServerConfig,
  settings: IntegrationSettingsStore,
): Promise<{ toolCount: number; proxied: boolean; proxyBaseUrl?: string }> {
  await postJsonRpc({
    server,
    method: "initialize",
    params: {
      protocolVersion: "2025-03-26",
      capabilities: {},
      clientInfo: {
        name: APP_NAME,
        version: APP_VERSION,
      },
    },
    settings,
  });

  await postJsonRpc({
    server,
    method: "notifications/initialized",
    settings,
    expectResponse: false,
  });

  const list = await postJsonRpc({
    server,
    method: "tools/list",
    params: {},
    settings,
  });

  if (!list) {
    throw new Error("MCP tools/list returned no response.");
  }

  return {
    toolCount: parseToolCountFromListResponse(list.response),
    proxied: list.proxied,
    proxyBaseUrl: list.proxyBaseUrl,
  };
}

function createIntegrationCard(args: {
  integration: IntegrationDefinition;
  snapshot: IntegrationsSnapshot;
  onToggleSession: (integrationId: string, next: boolean) => Promise<void>;
  onToggleWorkbook: (integrationId: string, next: boolean) => Promise<void>;
}): HTMLElement {
  const { integration, snapshot } = args;

  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-integrations-card";

  const top = document.createElement("div");
  top.className = "pi-integrations-card__top";

  const textWrap = document.createElement("div");
  textWrap.className = "pi-integrations-card__text-wrap";

  const title = document.createElement("strong");
  title.textContent = integration.title;
  title.className = "pi-integrations-card__title";

  const description = document.createElement("span");
  description.textContent = integration.description;
  description.className = "pi-integrations-card__description";

  textWrap.append(title, description);

  const badges = document.createElement("div");
  badges.className = "pi-overlay-badges";

  if (isEnabledInList(snapshot.activeIntegrationIds, integration.id) && snapshot.externalToolsEnabled) {
    badges.appendChild(createOverlayBadge("active", "ok"));
  } else if (isEnabledInList(snapshot.activeIntegrationIds, integration.id) && !snapshot.externalToolsEnabled) {
    badges.appendChild(createOverlayBadge("configured (blocked)", "warn"));
  } else {
    badges.appendChild(createOverlayBadge("inactive", "muted"));
  }

  top.append(textWrap, badges);

  const warning = document.createElement("div");
  warning.className = "pi-integrations-card__warning";
  warning.textContent = integration.warning ?? "";
  warning.hidden = !integration.warning;

  const toggles = document.createElement("div");
  toggles.className = "pi-integrations-card__toggles";

  const sessionLabel = document.createElement("label");
  sessionLabel.className = "pi-integrations-card__toggle-label";

  const sessionToggle = document.createElement("input");
  sessionToggle.type = "checkbox";
  sessionToggle.checked = isEnabledInList(snapshot.sessionIntegrationIds, integration.id);
  sessionToggle.addEventListener("change", () => {
    void args.onToggleSession(integration.id, sessionToggle.checked);
  });
  sessionLabel.append(sessionToggle, document.createTextNode("Enable for this session"));

  const workbookLabel = document.createElement("label");
  workbookLabel.className = "pi-integrations-card__toggle-label";

  const workbookToggle = document.createElement("input");
  workbookToggle.type = "checkbox";
  workbookToggle.checked = isEnabledInList(snapshot.workbookIntegrationIds, integration.id);
  workbookToggle.disabled = snapshot.workbookId === null;
  workbookToggle.addEventListener("change", () => {
    void args.onToggleWorkbook(integration.id, workbookToggle.checked);
  });

  const workbookText = snapshot.workbookId
    ? `Enable for workbook (${snapshot.workbookLabel})`
    : "Workbook scope unavailable";

  workbookLabel.append(workbookToggle, document.createTextNode(workbookText));

  toggles.append(sessionLabel, workbookLabel);

  card.append(top, warning, toggles);
  return card;
}

export function showIntegrationsDialog(dependencies: IntegrationsDialogDependencies): void {
  if (closeOverlayById(INTEGRATIONS_OVERLAY_ID)) {
    return;
  }

  const dialog = createOverlayDialog({
    overlayId: INTEGRATIONS_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--l pi-integrations-dialog",
  });

  const closeOverlay = dialog.close;

  const { header } = createOverlayHeader({
    onClose: closeOverlay,
    closeLabel: "Close integrations",
    title: INTEGRATIONS_LABEL,
    subtitle: "Add capabilities like web search and external integrations. External tools are off by default.",
  });

  const body = document.createElement("div");
  body.className = "pi-overlay-body";

  const externalSection = document.createElement("section");
  externalSection.className = "pi-overlay-section";
  externalSection.appendChild(createOverlaySectionTitle("External tools gate"));

  const externalCard = document.createElement("div");
  externalCard.className = "pi-overlay-surface";

  const externalToggleLabel = document.createElement("label");
  externalToggleLabel.className = "pi-integrations-toggle-label";

  const externalToggle = document.createElement("input");
  externalToggle.type = "checkbox";

  const externalToggleText = document.createElement("span");
  externalToggleText.textContent = "Allow external tools (web search / MCP)";

  externalToggleLabel.append(externalToggle, externalToggleText);

  const activeSummary = document.createElement("div");
  activeSummary.className = "pi-integrations-active-summary";

  externalCard.append(externalToggleLabel, activeSummary);
  externalSection.appendChild(externalCard);

  const integrationsSection = document.createElement("section");
  integrationsSection.className = "pi-overlay-section";
  integrationsSection.appendChild(createOverlaySectionTitle(`${INTEGRATIONS_LABEL} bundles`));

  const integrationsList = document.createElement("div");
  integrationsList.className = "pi-overlay-list";
  integrationsSection.appendChild(integrationsList);

  const webSearchSection = document.createElement("section");
  webSearchSection.className = "pi-overlay-section";
  webSearchSection.appendChild(createOverlaySectionTitle("Web search config"));

  const webSearchCard = document.createElement("div");
  webSearchCard.className = "pi-overlay-surface";

  const webSearchStatus = document.createElement("div");
  webSearchStatus.className = "pi-integrations-web-search-status";

  const webSearchProviderRow = document.createElement("div");
  webSearchProviderRow.className = "pi-integrations-web-search-provider-row";

  const webSearchProviderSelect = document.createElement("select");
  webSearchProviderSelect.className = "pi-overlay-input";

  for (const providerId of WEB_SEARCH_PROVIDERS) {
    const option = document.createElement("option");
    option.value = providerId;
    option.textContent = WEB_SEARCH_PROVIDER_INFO[providerId].title;
    webSearchProviderSelect.appendChild(option);
  }

  const webSearchProviderSignupLink = document.createElement("a");
  webSearchProviderSignupLink.className = "pi-overlay-link";
  webSearchProviderSignupLink.target = "_blank";
  webSearchProviderSignupLink.rel = "noopener noreferrer";

  webSearchProviderRow.append(webSearchProviderSelect, webSearchProviderSignupLink);

  const webSearchInputRow = document.createElement("div");
  webSearchInputRow.className = "pi-integrations-web-search-row";

  const webSearchApiKeyInput = createOverlayInput({ placeholder: "API key", type: "password" });
  const webSearchSaveButton = createOverlayButton({ text: "Save key" });
  const webSearchValidateButton = createOverlayButton({ text: "Validate" });
  const webSearchClearButton = createOverlayButton({ text: "Clear" });

  webSearchInputRow.append(
    webSearchApiKeyInput,
    webSearchSaveButton,
    webSearchValidateButton,
    webSearchClearButton,
  );

  const webSearchHint = document.createElement("p");
  webSearchHint.className = "pi-overlay-hint";

  const webSearchValidationStatus = document.createElement("p");
  webSearchValidationStatus.className = "pi-overlay-hint";

  webSearchCard.append(
    webSearchStatus,
    webSearchProviderRow,
    webSearchInputRow,
    webSearchHint,
    webSearchValidationStatus,
  );
  webSearchSection.appendChild(webSearchCard);

  const mcpSection = document.createElement("section");
  mcpSection.className = "pi-overlay-section";
  mcpSection.appendChild(createOverlaySectionTitle("MCP servers"));

  const mcpList = document.createElement("div");
  mcpList.className = "pi-overlay-list";

  const mcpAddCard = document.createElement("div");
  mcpAddCard.className = "pi-overlay-surface";

  const mcpAddTitle = document.createElement("div");
  mcpAddTitle.textContent = "Add server";
  mcpAddTitle.className = "pi-integrations-mcp-add-title";

  const mcpAddRow = document.createElement("div");
  mcpAddRow.className = "pi-integrations-mcp-add-row";

  const mcpNameInput = createOverlayInput({ placeholder: "Name" });
  const mcpUrlInput = createOverlayInput({ placeholder: "https://example.com/mcp" });
  const mcpTokenInput = createOverlayInput({ placeholder: "Bearer token (optional)", type: "password" });

  const mcpEnabledLabel = document.createElement("label");
  mcpEnabledLabel.className = "pi-integrations-toggle-label";
  const mcpEnabledInput = document.createElement("input");
  mcpEnabledInput.type = "checkbox";
  mcpEnabledInput.checked = true;
  mcpEnabledLabel.append(mcpEnabledInput, document.createTextNode("Enabled"));

  const mcpAddButton = createOverlayButton({ text: "Add" });

  mcpAddRow.append(mcpNameInput, mcpUrlInput, mcpTokenInput, mcpEnabledLabel, mcpAddButton);

  const mcpHint = document.createElement("p");
  mcpHint.className = "pi-overlay-hint";
  mcpHint.textContent = "Server URL, optional bearer token, and one-click connection test.";

  mcpAddCard.append(mcpAddTitle, mcpAddRow, mcpHint);

  mcpSection.append(mcpList, mcpAddCard);

  body.append(externalSection, integrationsSection, webSearchSection, mcpSection);
  dialog.card.append(header, body);

  let busy = false;
  let snapshot: IntegrationsSnapshot | null = null;

  const setBusy = (next: boolean): void => {
    busy = next;
    externalToggle.disabled = next;
    webSearchProviderSelect.disabled = next;
    webSearchApiKeyInput.disabled = next;
    webSearchSaveButton.disabled = next;
    webSearchValidateButton.disabled = next;
    webSearchClearButton.disabled = next;
    mcpNameInput.disabled = next;
    mcpUrlInput.disabled = next;
    mcpTokenInput.disabled = next;
    mcpEnabledInput.disabled = next;
    mcpAddButton.disabled = next;
  };

  const afterMutation = async (reason: "toggle" | "scope" | "external-toggle" | "config"): Promise<void> => {
    dispatchIntegrationsChanged({ reason });
    if (dependencies.onChanged) {
      await dependencies.onChanged();
    }
  };

  const runAction = async (
    action: () => Promise<void>,
    reason: "toggle" | "scope" | "external-toggle" | "config",
    successMessage?: string,
  ): Promise<void> => {
    if (busy) return;
    setBusy(true);

    try {
      await action();
      await afterMutation(reason);
      await refresh();
      if (successMessage) {
        showToast(successMessage);
      }
    } catch (error: unknown) {
      showToast(`Integrations: ${getErrorMessage(error)}`);
    } finally {
      setBusy(false);
    }
  };

  const renderMcpServerRow = (server: McpServerConfig): HTMLElement => {
    const row = document.createElement("div");
    row.className = "pi-overlay-surface pi-integrations-mcp-row";

    const top = document.createElement("div");
    top.className = "pi-integrations-mcp-row__top";

    const info = document.createElement("div");
    info.className = "pi-integrations-mcp-row__info";

    const name = document.createElement("strong");
    name.textContent = server.name;
    name.className = "pi-integrations-mcp-row__name";

    const url = document.createElement("code");
    url.textContent = server.url;
    url.className = "pi-integrations-mcp-row__url";

    info.append(name, url);

    const badges = document.createElement("div");
    badges.className = "pi-overlay-badges";
    badges.appendChild(createOverlayBadge(server.enabled ? "enabled" : "disabled", server.enabled ? "ok" : "muted"));
    if (server.token) {
      badges.appendChild(createOverlayBadge("token set", "muted"));
    }

    top.append(info, badges);

    const actions = document.createElement("div");
    actions.className = "pi-overlay-actions pi-overlay-actions--wrap";

    const testButton = createOverlayButton({ text: "Test" });
    const removeButton = createOverlayButton({ text: "Remove" });

    testButton.addEventListener("click", () => {
      void runAction(async () => {
        const settings = await getSettingsStore();
        const result = await probeMcpServer(server, settings);
        const transport = result.proxied ? `proxy (${result.proxyBaseUrl ?? "configured"})` : "direct";
        showToast(`MCP ${server.name}: reachable (${result.toolCount} tool${result.toolCount === 1 ? "" : "s"}, ${transport})`);
      }, "config");
    });

    removeButton.addEventListener("click", () => {
      void runAction(async () => {
        const settings = await getSettingsStore();
        const servers = await loadMcpServers(settings);
        const next = servers.filter((entry) => entry.id !== server.id);
        await saveMcpServers(settings, next);
      }, "config", `Removed MCP server: ${server.name}`);
    });

    actions.append(testButton, removeButton);
    row.append(top, actions);
    return row;
  };

  const render = (): void => {
    if (!snapshot) return;

    const currentSnapshot = snapshot;
    externalToggle.checked = currentSnapshot.externalToolsEnabled;

    const activeNames = currentSnapshot.activeIntegrationIds
      .map((integrationId) => listIntegrationDefinitions().find((integration) => integration.id === integrationId)?.title ?? integrationId)
      .join(", ");

    activeSummary.textContent = currentSnapshot.externalToolsEnabled
      ? (currentSnapshot.activeIntegrationIds.length > 0
        ? `Active now: ${activeNames}`
        : `No active ${INTEGRATIONS_LABEL_LOWER} in this session/workbook.`)
      : `External tools are disabled globally. ${INTEGRATIONS_LABEL} remain configured but inactive.`;

    integrationsList.replaceChildren();
    for (const integration of listIntegrationDefinitions()) {
      integrationsList.appendChild(createIntegrationCard({
        integration,
        snapshot: currentSnapshot,
        onToggleSession: async (integrationId, next) => {
          await runAction(async () => {
            const settings = await getSettingsStore();
            await setIntegrationEnabledInScope({
              settings,
              scope: "session",
              identifier: currentSnapshot.sessionId,
              integrationId,
              enabled: next,
              knownIntegrationIds: INTEGRATION_IDS,
            });
          }, "scope", `${integration.title}: ${next ? "enabled" : "disabled"} for this session`);
        },
        onToggleWorkbook: async (integrationId, next) => {
          const workbookId = currentSnapshot.workbookId;
          if (!workbookId) return;

          await runAction(async () => {
            const settings = await getSettingsStore();
            await setIntegrationEnabledInScope({
              settings,
              scope: "workbook",
              identifier: workbookId,
              integrationId,
              enabled: next,
              knownIntegrationIds: INTEGRATION_IDS,
            });
          }, "scope", `${integration.title}: ${next ? "enabled" : "disabled"} for workbook`);
        },
      }));
    }

    const selectedProvider = currentSnapshot.webSearchConfig.provider;
    const selectedProviderInfo = WEB_SEARCH_PROVIDER_INFO[selectedProvider];
    const selectedProviderKey = getApiKeyForProvider(currentSnapshot.webSearchConfig, selectedProvider);

    webSearchProviderSelect.value = selectedProvider;
    webSearchProviderSignupLink.href = selectedProviderInfo.signupUrl;
    webSearchProviderSignupLink.textContent = `Get key (${selectedProviderInfo.title})`;
    webSearchApiKeyInput.placeholder = selectedProviderInfo.apiKeyLabel;

    if (selectedProviderKey) {
      webSearchStatus.textContent = `${selectedProviderInfo.apiKeyLabel}: ${maskSecret(selectedProviderKey)} (length ${selectedProviderKey.length})`;
    } else {
      webSearchStatus.textContent = `${selectedProviderInfo.apiKeyLabel} not set.`;
    }

    webSearchHint.textContent = `${selectedProviderInfo.shortDescription} ${selectedProviderInfo.apiKeyHelp} Used by web_search and fetch_page.`;
    webSearchValidationStatus.textContent = "";

    mcpList.replaceChildren();
    if (currentSnapshot.mcpServers.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = "No MCP servers configured.";
      empty.className = "pi-overlay-empty";
      mcpList.appendChild(empty);
    } else {
      for (const server of currentSnapshot.mcpServers) {
        mcpList.appendChild(renderMcpServerRow(server));
      }
    }
  };

  const refresh = async (): Promise<void> => {
    snapshot = await buildSnapshot(dependencies);
    render();
  };

  const getSelectedWebSearchProvider = (): WebSearchProvider => {
    return normalizeWebSearchProvider(webSearchProviderSelect.value);
  };

  externalToggle.addEventListener("change", () => {
    const next = externalToggle.checked;
    void runAction(async () => {
      const settings = await getSettingsStore();
      await setExternalToolsEnabled(settings, next);
    }, "external-toggle", `External tools: ${next ? "enabled" : "disabled"}`);
  });

  webSearchProviderSelect.addEventListener("change", () => {
    const provider = getSelectedWebSearchProvider();
    void runAction(async () => {
      const settings = await getSettingsStore();
      await saveWebSearchProvider(settings, provider);
      webSearchValidationStatus.textContent = "";
    }, "config", `Web search provider set to ${WEB_SEARCH_PROVIDER_INFO[provider].title}.`);
  });

  webSearchSaveButton.addEventListener("click", () => {
    void runAction(async () => {
      const key = webSearchApiKeyInput.value.trim();
      const provider = getSelectedWebSearchProvider();
      if (key.length === 0) {
        throw new Error(`Provide a ${WEB_SEARCH_PROVIDER_INFO[provider].apiKeyLabel}.`);
      }

      const settings = await getSettingsStore();
      await saveWebSearchApiKey(settings, provider, key);
      webSearchApiKeyInput.value = "";
      webSearchValidationStatus.textContent = "";
    }, "config", `Saved ${WEB_SEARCH_PROVIDER_INFO[getSelectedWebSearchProvider()].apiKeyLabel}.`);
  });

  webSearchClearButton.addEventListener("click", () => {
    void runAction(async () => {
      const provider = getSelectedWebSearchProvider();
      const settings = await getSettingsStore();
      await clearWebSearchApiKey(settings, provider);
      webSearchApiKeyInput.value = "";
      webSearchValidationStatus.textContent = "";
    }, "config", `Cleared ${WEB_SEARCH_PROVIDER_INFO[getSelectedWebSearchProvider()].apiKeyLabel}.`);
  });

  webSearchValidateButton.addEventListener("click", () => {
    if (busy) return;

    const provider = getSelectedWebSearchProvider();
    const enteredKey = webSearchApiKeyInput.value.trim();

    void (async () => {
      setBusy(true);
      try {
        const settings = await getSettingsStore();
        const providerConfig = await loadWebSearchProviderConfig(settings);
        const apiKey = enteredKey.length > 0
          ? enteredKey
          : (getApiKeyForProvider(providerConfig, provider) ?? "");

        if (apiKey.length === 0) {
          throw new Error(`No ${WEB_SEARCH_PROVIDER_INFO[provider].apiKeyLabel} available to validate.`);
        }

        const proxyBaseUrl = await getEnabledProxyBaseUrl(settings);
        const validation = await validateWebSearchApiKey({
          provider,
          apiKey,
          proxyBaseUrl,
        });

        webSearchValidationStatus.textContent = validation.ok
          ? `✓ ${validation.message}`
          : `✗ ${validation.message}`;
      } catch (error: unknown) {
        webSearchValidationStatus.textContent = `✗ ${getErrorMessage(error)}`;
      } finally {
        setBusy(false);
      }
    })();
  });

  mcpAddButton.addEventListener("click", () => {
    void runAction(async () => {
      const settings = await getSettingsStore();
      const servers = await loadMcpServers(settings);
      const nextServer = createMcpServerConfig({
        name: mcpNameInput.value,
        url: mcpUrlInput.value,
        token: mcpTokenInput.value,
        enabled: mcpEnabledInput.checked,
      });

      await saveMcpServers(settings, [...servers, nextServer]);
      mcpNameInput.value = "";
      mcpUrlInput.value = "";
      mcpTokenInput.value = "";
      mcpEnabledInput.checked = true;
    }, "config", "Added MCP server.");
  });

  dialog.mount();
  setBusy(true);
  void refresh()
    .catch((error: unknown) => {
      showToast(`Integrations: ${getErrorMessage(error)}`);
      closeOverlay();
    })
    .finally(() => {
      setBusy(false);
    });
}
