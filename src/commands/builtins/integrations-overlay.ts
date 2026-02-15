/**
 * Integrations manager overlay.
 */

import {
  INTEGRATION_IDS,
  listIntegrationDefinitions,
} from "../../integrations/catalog.js";
import { dispatchIntegrationsChanged } from "../../integrations/events.js";
import {
  INTEGRATIONS_LABEL,
  INTEGRATIONS_LABEL_LOWER,
  INTEGRATIONS_MANAGER_LABEL,
} from "../../integrations/naming.js";
import {
  setExternalToolsEnabled,
  setIntegrationEnabledInScope,
} from "../../integrations/store.js";
import { getEnabledProxyBaseUrl } from "../../tools/external-fetch.js";
import {
  clearWebSearchApiKey,
  getApiKeyForProvider,
  loadWebSearchProviderConfig,
  maskSecret,
  saveWebSearchApiKey,
  saveWebSearchProvider,
  WEB_SEARCH_PROVIDER_INFO,
  type WebSearchProvider,
} from "../../tools/web-search-config.js";
import { validateWebSearchApiKey } from "../../tools/web-search.js";
import {
  createMcpServerConfig,
  loadMcpServers,
  saveMcpServers,
  type McpServerConfig,
} from "../../tools/mcp-config.js";
import {
  closeOverlayById,
  createOverlayBadge,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "../../ui/overlay-dialog.js";
import { INTEGRATIONS_OVERLAY_ID } from "../../ui/overlay-ids.js";
import { showToast } from "../../ui/toast.js";
import { createIntegrationCard } from "./integrations-overlay-card.js";
import { createIntegrationsDialogElements } from "./integrations-overlay-elements.js";
import { probeMcpServer } from "./integrations-overlay-mcp-probe.js";
import {
  buildSnapshot,
  getSettingsStore,
  normalizeWebSearchProvider,
} from "./integrations-overlay-state.js";
import {
  getIntegrationsErrorMessage,
  type IntegrationMutationReason,
  type IntegrationsDialogDependencies,
  type IntegrationsSnapshot,
} from "./integrations-overlay-types.js";

function buildActiveIntegrationsSummary(snapshot: IntegrationsSnapshot): string {
  const activeNames = snapshot.activeIntegrationIds
    .map((integrationId) => listIntegrationDefinitions().find((integration) => integration.id === integrationId)?.title ?? integrationId)
    .join(", ");

  if (!snapshot.externalToolsEnabled) {
    return `External tools are disabled globally. ${INTEGRATIONS_LABEL} remain configured but inactive.`;
  }

  if (snapshot.activeIntegrationIds.length > 0) {
    return `Active now: ${activeNames}`;
  }

  return `No active ${INTEGRATIONS_LABEL_LOWER} in this session/workbook.`;
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
    closeLabel: "Close tools and MCP",
    title: INTEGRATIONS_MANAGER_LABEL,
    subtitle: "Manage web search, page fetch, and MCP servers. External tools are off by default.",
  });

  const elements = createIntegrationsDialogElements();
  dialog.card.append(header, elements.body);

  let busy = false;
  let snapshot: IntegrationsSnapshot | null = null;

  const setBusy = (next: boolean): void => {
    busy = next;
    elements.externalToggle.disabled = next;
    elements.webSearchProviderSelect.disabled = next;
    elements.webSearchApiKeyInput.disabled = next;
    elements.webSearchSaveButton.disabled = next;
    elements.webSearchValidateButton.disabled = next;
    elements.webSearchClearButton.disabled = next;
    elements.mcpNameInput.disabled = next;
    elements.mcpUrlInput.disabled = next;
    elements.mcpTokenInput.disabled = next;
    elements.mcpEnabledInput.disabled = next;
    elements.mcpAddButton.disabled = next;
  };

  const afterMutation = async (reason: IntegrationMutationReason): Promise<void> => {
    dispatchIntegrationsChanged({ reason });
    if (dependencies.onChanged) {
      await dependencies.onChanged();
    }
  };

  const refresh = async (): Promise<void> => {
    snapshot = await buildSnapshot(dependencies);
    render();
  };

  const runAction = async (
    action: () => Promise<void>,
    reason: IntegrationMutationReason,
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
      showToast(`Integrations: ${getIntegrationsErrorMessage(error)}`);
    } finally {
      setBusy(false);
    }
  };

  const getSelectedWebSearchProvider = (): WebSearchProvider => {
    return normalizeWebSearchProvider(elements.webSearchProviderSelect.value);
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
    const removeButton = createOverlayButton({
      text: "Remove",
      className: "pi-overlay-btn--danger",
    });

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
    elements.externalToggle.checked = currentSnapshot.externalToolsEnabled;
    elements.activeSummary.textContent = buildActiveIntegrationsSummary(currentSnapshot);

    elements.integrationsList.replaceChildren();
    for (const integration of listIntegrationDefinitions()) {
      elements.integrationsList.appendChild(createIntegrationCard({
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

    elements.webSearchProviderSelect.value = selectedProvider;
    elements.webSearchProviderSignupLink.href = selectedProviderInfo.signupUrl;
    elements.webSearchProviderSignupLink.textContent = `Get key (${selectedProviderInfo.title})`;
    elements.webSearchApiKeyInput.placeholder = selectedProviderInfo.apiKeyLabel;

    if (selectedProviderKey) {
      elements.webSearchStatus.textContent = `${selectedProviderInfo.apiKeyLabel}: ${maskSecret(selectedProviderKey)} (length ${selectedProviderKey.length})`;
    } else {
      elements.webSearchStatus.textContent = `${selectedProviderInfo.apiKeyLabel} not set.`;
    }

    elements.webSearchHint.textContent = `${selectedProviderInfo.shortDescription} ${selectedProviderInfo.apiKeyHelp} Used by web_search and fetch_page.`;
    elements.webSearchValidationStatus.textContent = "";

    elements.mcpList.replaceChildren();
    if (currentSnapshot.mcpServers.length === 0) {
      const empty = document.createElement("div");
      empty.textContent = "No MCP servers configured.";
      empty.className = "pi-overlay-empty";
      elements.mcpList.appendChild(empty);
    } else {
      for (const server of currentSnapshot.mcpServers) {
        elements.mcpList.appendChild(renderMcpServerRow(server));
      }
    }
  };

  elements.externalToggle.addEventListener("change", () => {
    const next = elements.externalToggle.checked;
    void runAction(async () => {
      const settings = await getSettingsStore();
      await setExternalToolsEnabled(settings, next);
    }, "external-toggle", `External tools: ${next ? "enabled" : "disabled"}`);
  });

  elements.webSearchProviderSelect.addEventListener("change", () => {
    const provider = getSelectedWebSearchProvider();
    void runAction(async () => {
      const settings = await getSettingsStore();
      await saveWebSearchProvider(settings, provider);
      elements.webSearchValidationStatus.textContent = "";
    }, "config", `Web search provider set to ${WEB_SEARCH_PROVIDER_INFO[provider].title}.`);
  });

  elements.webSearchSaveButton.addEventListener("click", () => {
    void runAction(async () => {
      const key = elements.webSearchApiKeyInput.value.trim();
      const provider = getSelectedWebSearchProvider();
      if (key.length === 0) {
        throw new Error(`Provide a ${WEB_SEARCH_PROVIDER_INFO[provider].apiKeyLabel}.`);
      }

      const settings = await getSettingsStore();
      await saveWebSearchApiKey(settings, provider, key);
      elements.webSearchApiKeyInput.value = "";
      elements.webSearchValidationStatus.textContent = "";
    }, "config", `Saved ${WEB_SEARCH_PROVIDER_INFO[getSelectedWebSearchProvider()].apiKeyLabel}.`);
  });

  elements.webSearchClearButton.addEventListener("click", () => {
    void runAction(async () => {
      const provider = getSelectedWebSearchProvider();
      const settings = await getSettingsStore();
      await clearWebSearchApiKey(settings, provider);
      elements.webSearchApiKeyInput.value = "";
      elements.webSearchValidationStatus.textContent = "";
    }, "config", `Cleared ${WEB_SEARCH_PROVIDER_INFO[getSelectedWebSearchProvider()].apiKeyLabel}.`);
  });

  elements.webSearchValidateButton.addEventListener("click", () => {
    if (busy) return;

    const provider = getSelectedWebSearchProvider();
    const enteredKey = elements.webSearchApiKeyInput.value.trim();

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

        elements.webSearchValidationStatus.textContent = validation.ok
          ? `✓ ${validation.message}`
          : `✗ ${validation.message}`;
      } catch (error: unknown) {
        elements.webSearchValidationStatus.textContent = `✗ ${getIntegrationsErrorMessage(error)}`;
      } finally {
        setBusy(false);
      }
    })();
  });

  elements.mcpAddButton.addEventListener("click", () => {
    void runAction(async () => {
      const settings = await getSettingsStore();
      const servers = await loadMcpServers(settings);
      const nextServer = createMcpServerConfig({
        name: elements.mcpNameInput.value,
        url: elements.mcpUrlInput.value,
        token: elements.mcpTokenInput.value,
        enabled: elements.mcpEnabledInput.checked,
      });

      await saveMcpServers(settings, [...servers, nextServer]);
      elements.mcpNameInput.value = "";
      elements.mcpUrlInput.value = "";
      elements.mcpTokenInput.value = "";
      elements.mcpEnabledInput.checked = true;
    }, "config", "Added MCP server.");
  });

  dialog.mount();
  setBusy(true);
  void refresh()
    .catch((error: unknown) => {
      showToast(`Integrations: ${getIntegrationsErrorMessage(error)}`);
      closeOverlay();
    })
    .finally(() => {
      setBusy(false);
    });
}

export type { IntegrationsDialogDependencies } from "./integrations-overlay-types.js";
