import { validateOfficeProxyUrl } from "../../auth/proxy-validation.js";
import { dispatchExperimentalToolConfigChanged } from "../../experiments/events.js";
import {
  INTEGRATION_IDS,
  listIntegrationDefinitions,
} from "../../integrations/catalog.js";
import {
  getExternalToolsEnabled,
  getSessionIntegrationIds,
  getWorkbookIntegrationIds,
  resolveConfiguredIntegrationIds,
  setExternalToolsEnabled,
  setIntegrationEnabledInScope,
} from "../../integrations/store.js";
import {
  PYTHON_BRIDGE_URL_SETTING_KEY,
} from "../../tools/experimental-tool-gates.js";
import {
  getApiKeyForProvider,
  loadWebSearchProviderConfig,
  maskSecret,
  WEB_SEARCH_PROVIDER_INFO,
} from "../../tools/web-search-config.js";
import { loadMcpServers } from "../../tools/mcp-config.js";
import {
  createOverlayBadge,
  createOverlayButton,
  createOverlayInput,
  createOverlaySectionTitle,
} from "../../ui/overlay-dialog.js";
import { showToast } from "../../ui/toast.js";
import {
  type AddonsDialogActions,
  type AddonsSettingsStore,
  type ConnectionsSnapshot,
} from "./addons-overlay-types.js";

function normalizeOptionalString(value: unknown): string {
  if (typeof value !== "string") {
    return "";
  }

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : "";
}

export async function buildConnectionsSnapshot(
  settings: AddonsSettingsStore,
  actions: AddonsDialogActions,
): Promise<ConnectionsSnapshot> {
  const sessionId = actions.getActiveSessionId();
  const workbookContext = await actions.resolveWorkbookContext();

  const [
    externalToolsEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    activeIntegrationIds,
    webSearchConfig,
    mcpServers,
    pythonBridgeUrlRaw,
  ] = await Promise.all([
    getExternalToolsEnabled(settings),
    sessionId
      ? getSessionIntegrationIds(settings, sessionId, INTEGRATION_IDS, {
        applyDefaultsWhenUnconfigured: workbookContext.workbookId === null,
      })
      : Promise.resolve([]),
    workbookContext.workbookId
      ? getWorkbookIntegrationIds(settings, workbookContext.workbookId, INTEGRATION_IDS)
      : Promise.resolve([]),
    sessionId
      ? resolveConfiguredIntegrationIds({
        settings,
        sessionId,
        workbookId: workbookContext.workbookId,
        knownIntegrationIds: INTEGRATION_IDS,
      })
      : Promise.resolve([]),
    loadWebSearchProviderConfig(settings),
    loadMcpServers(settings),
    settings.get(PYTHON_BRIDGE_URL_SETTING_KEY),
  ]);

  const webSearchInfo = WEB_SEARCH_PROVIDER_INFO[webSearchConfig.provider];
  const webSearchApiKey = getApiKeyForProvider(webSearchConfig);
  const webSearchStatusText = webSearchApiKey
    ? `Provider: ${webSearchInfo.title} · key set (${maskSecret(webSearchApiKey)})`
    : `Provider: ${webSearchInfo.title} · key not set`;

  const enabledMcpServers = mcpServers.filter((server) => server.enabled);
  const mcpStatusText = mcpServers.length === 0
    ? "No MCP servers configured"
    : `${enabledMcpServers.length}/${mcpServers.length} MCP server${mcpServers.length === 1 ? "" : "s"} enabled`;

  return {
    sessionId,
    workbookId: workbookContext.workbookId,
    workbookLabel: workbookContext.workbookLabel,
    externalToolsEnabled,
    sessionIntegrationIds,
    workbookIntegrationIds,
    activeIntegrationIds,
    webSearchStatusText,
    mcpStatusText,
    pythonBridgeUrl: normalizeOptionalString(pythonBridgeUrlRaw),
  };
}

function buildIntegrationCard(args: {
  snapshot: ConnectionsSnapshot;
  integrationId: string;
  title: string;
  description: string;
  warning?: string;
  onToggleSession: (next: boolean) => void;
  onToggleWorkbook: (next: boolean) => void;
  busy: boolean;
}): HTMLElement {
  const card = document.createElement("div");
  card.className = "pi-overlay-surface pi-addons-connection";

  const top = document.createElement("div");
  top.className = "pi-addons-connection__top";

  const textWrap = document.createElement("div");
  textWrap.className = "pi-addons-connection__text-wrap";

  const titleEl = document.createElement("strong");
  titleEl.className = "pi-addons-connection__title";
  titleEl.textContent = args.title;

  const descriptionEl = document.createElement("span");
  descriptionEl.className = "pi-addons-connection__description";
  descriptionEl.textContent = args.description;

  textWrap.append(titleEl, descriptionEl);

  const badges = document.createElement("div");
  badges.className = "pi-overlay-badges";

  const active = args.snapshot.activeIntegrationIds.includes(args.integrationId);
  const blockedByGlobalGate = active && !args.snapshot.externalToolsEnabled;

  if (blockedByGlobalGate) {
    badges.appendChild(createOverlayBadge("configured (blocked)", "warn"));
  } else {
    badges.appendChild(createOverlayBadge(active ? "active" : "inactive", active ? "ok" : "muted"));
  }

  top.append(textWrap, badges);

  const toggles = document.createElement("div");
  toggles.className = "pi-addons-connection__toggles";

  const sessionLabel = document.createElement("label");
  sessionLabel.className = "pi-addons-connection__toggle";

  const sessionToggle = document.createElement("input");
  sessionToggle.type = "checkbox";
  sessionToggle.disabled = args.busy || args.snapshot.sessionId === null;
  sessionToggle.checked = args.snapshot.sessionIntegrationIds.includes(args.integrationId);
  sessionToggle.addEventListener("change", () => {
    args.onToggleSession(sessionToggle.checked);
  });

  sessionLabel.append(sessionToggle, document.createTextNode("Enable for this session"));

  const workbookLabel = document.createElement("label");
  workbookLabel.className = "pi-addons-connection__toggle";

  const workbookToggle = document.createElement("input");
  workbookToggle.type = "checkbox";
  workbookToggle.disabled = args.busy || args.snapshot.workbookId === null;
  workbookToggle.checked = args.snapshot.workbookIntegrationIds.includes(args.integrationId);
  workbookToggle.addEventListener("change", () => {
    args.onToggleWorkbook(workbookToggle.checked);
  });

  const workbookText = args.snapshot.workbookId
    ? `Enable for workbook (${args.snapshot.workbookLabel})`
    : "Workbook scope unavailable";
  workbookLabel.append(workbookToggle, document.createTextNode(workbookText));

  toggles.append(sessionLabel, workbookLabel);

  card.append(top, toggles);

  if (args.warning) {
    const warning = document.createElement("p");
    warning.className = "pi-overlay-hint pi-overlay-text-warning";
    warning.textContent = args.warning;
    card.appendChild(warning);
  }

  return card;
}

export function renderConnectionsSection(args: {
  container: HTMLElement;
  snapshot: ConnectionsSnapshot;
  settings: AddonsSettingsStore;
  actions: AddonsDialogActions;
  busy: boolean;
  onRefresh: () => void;
  onMutate: (
    mutation: () => Promise<void>,
    reason: "toggle" | "scope" | "external-toggle" | "config",
    successMessage?: string,
  ) => Promise<void>;
}): void {
  args.container.replaceChildren();

  const section = document.createElement("section");
  section.className = "pi-overlay-section pi-addons-section";
  section.dataset.addonsSection = "connections";
  section.appendChild(createOverlaySectionTitle("Connections"));

  const hint = document.createElement("p");
  hint.className = "pi-overlay-hint";
  hint.textContent = "External services Pi can call (web search, MCP, Python bridge).";
  section.appendChild(hint);

  const statusCard = document.createElement("div");
  statusCard.className = "pi-overlay-surface pi-addons-connection-status";

  const globalToggleLabel = document.createElement("label");
  globalToggleLabel.className = "pi-addons-connection__toggle";

  const globalToggle = document.createElement("input");
  globalToggle.type = "checkbox";
  globalToggle.checked = args.snapshot.externalToolsEnabled;
  globalToggle.disabled = args.busy;
  globalToggle.addEventListener("change", () => {
    void args.onMutate(
      () => setExternalToolsEnabled(args.settings, globalToggle.checked),
      "external-toggle",
      `External tools ${globalToggle.checked ? "enabled" : "disabled"}`,
    );
  });

  globalToggleLabel.append(globalToggle, document.createTextNode("Enable external tools globally"));

  const summary = document.createElement("p");
  summary.className = "pi-overlay-hint";
  if (args.snapshot.activeIntegrationIds.length === 0) {
    summary.textContent = "No integrations enabled for the current session/workbook.";
  } else if (!args.snapshot.externalToolsEnabled) {
    summary.textContent = `Configured integrations are blocked while external tools are disabled: ${args.snapshot.activeIntegrationIds.join(", ")}`;
  } else {
    summary.textContent = `Active integrations: ${args.snapshot.activeIntegrationIds.join(", ")}`;
  }

  const webSearchStatus = document.createElement("p");
  webSearchStatus.className = "pi-overlay-hint";
  webSearchStatus.textContent = `Web search · ${args.snapshot.webSearchStatusText}`;

  const mcpStatus = document.createElement("p");
  mcpStatus.className = "pi-overlay-hint";
  mcpStatus.textContent = `MCP · ${args.snapshot.mcpStatusText}`;

  statusCard.append(globalToggleLabel, summary, webSearchStatus, mcpStatus);

  const list = document.createElement("div");
  list.className = "pi-overlay-list";

  for (const integration of listIntegrationDefinitions()) {
    const card = buildIntegrationCard({
      snapshot: args.snapshot,
      integrationId: integration.id,
      title: integration.title,
      description: integration.description,
      warning: integration.warning,
      busy: args.busy,
      onToggleSession: (next: boolean) => {
        if (!args.snapshot.sessionId) {
          showToast("No active session");
          return;
        }

        void args.onMutate(async () => {
          await setIntegrationEnabledInScope({
            settings: args.settings,
            scope: "session",
            identifier: args.snapshot.sessionId ?? "",
            integrationId: integration.id,
            enabled: next,
            knownIntegrationIds: INTEGRATION_IDS,
          });
        }, "scope", `${integration.title} ${next ? "enabled" : "disabled"} for this session`);
      },
      onToggleWorkbook: (next: boolean) => {
        if (!args.snapshot.workbookId) {
          showToast("Workbook scope unavailable");
          return;
        }

        void args.onMutate(async () => {
          await setIntegrationEnabledInScope({
            settings: args.settings,
            scope: "workbook",
            identifier: args.snapshot.workbookId ?? "",
            integrationId: integration.id,
            enabled: next,
            knownIntegrationIds: INTEGRATION_IDS,
          });
        }, "scope", `${integration.title} ${next ? "enabled" : "disabled"} for this workbook`);
      },
    });

    list.appendChild(card);
  }

  const pythonCard = document.createElement("div");
  pythonCard.className = "pi-overlay-surface pi-addons-python-bridge";

  const pythonTitle = document.createElement("strong");
  pythonTitle.className = "pi-addons-python-bridge__title";
  pythonTitle.textContent = "Python bridge URL";

  const pythonRow = document.createElement("div");
  pythonRow.className = "pi-addons-python-bridge__row";

  const pythonInput = createOverlayInput({
    placeholder: "https://localhost:3340",
    className: "pi-addons-python-bridge__input",
  });
  pythonInput.type = "text";
  pythonInput.value = args.snapshot.pythonBridgeUrl;
  pythonInput.disabled = args.busy;

  const saveButton = createOverlayButton({
    text: "Save",
    className: "pi-overlay-btn--primary",
  });
  saveButton.disabled = args.busy;

  const clearButton = createOverlayButton({ text: "Clear" });
  clearButton.disabled = args.busy;

  const saveBridgeUrl = async (clear: boolean): Promise<void> => {
    const raw = clear ? "" : pythonInput.value.trim();

    if (raw.length === 0) {
      await args.onMutate(async () => {
        if (typeof args.settings.delete === "function") {
          await args.settings.delete(PYTHON_BRIDGE_URL_SETTING_KEY);
        } else {
          await args.settings.set(PYTHON_BRIDGE_URL_SETTING_KEY, "");
        }
        dispatchExperimentalToolConfigChanged({ configKey: PYTHON_BRIDGE_URL_SETTING_KEY });
      }, "config", "Python bridge URL cleared");
      return;
    }

    let normalizedUrl: string;
    try {
      normalizedUrl = validateOfficeProxyUrl(raw);
    } catch (error: unknown) {
      const message = error instanceof Error ? error.message : "Invalid URL";
      showToast(`Python bridge URL not saved: ${message}`);
      args.onRefresh();
      return;
    }

    await args.onMutate(async () => {
      await args.settings.set(PYTHON_BRIDGE_URL_SETTING_KEY, normalizedUrl);
      dispatchExperimentalToolConfigChanged({ configKey: PYTHON_BRIDGE_URL_SETTING_KEY });
    }, "config", "Python bridge URL saved");
  };

  saveButton.addEventListener("click", () => {
    void saveBridgeUrl(false);
  });
  clearButton.addEventListener("click", () => {
    void saveBridgeUrl(true);
  });
  pythonInput.addEventListener("keydown", (event) => {
    if (event.key !== "Enter") {
      return;
    }

    event.preventDefault();
    void saveBridgeUrl(false);
  });

  const pythonHint = document.createElement("p");
  pythonHint.className = "pi-overlay-hint";
  pythonHint.textContent = "Use /experimental python-bridge-token <token> to set auth token.";

  pythonRow.append(pythonInput, saveButton, clearButton);
  pythonCard.append(pythonTitle, pythonRow, pythonHint);

  const actionsRow = document.createElement("div");
  actionsRow.className = "pi-overlay-actions";

  const openDetailed = createOverlayButton({ text: "Open detailed Tools & MCP manager…" });
  openDetailed.addEventListener("click", () => {
    args.actions.openIntegrationsManager();
  });

  actionsRow.appendChild(openDetailed);

  section.append(statusCard, list, pythonCard, actionsRow);
  args.container.appendChild(section);
}
