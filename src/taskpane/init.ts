function isTaskpaneInitPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Taskpane initialization.
 *
 * Contains the bulk of app wiring (storage, auth restore, runtime manager,
 * sidebar mount, slash commands, keyboard shortcuts).
 */

import { html, render } from "lit";
import { BACKGROUND_CONTEXT as durableContext } from "@earendil-works/chord/context";
import type { Api, Model } from "@earendil-works/pi-ai";
import { ConversationPresenter } from "../durable/conversation-view.js";
import { createDurableTaskpane } from "../durable/taskpane.js";
import type { DurableCapabilities } from "../durable/runtime.js";
import type { WorkbookMutationGate } from "../tools/mutation/gate.js";
import { SessionInfoDoc } from "../durable/sessions.js";
import type { SessionData } from "../storage/local/types.js";

import { createOfficeStreamFn } from "../auth/stream-proxy.js";
import {
  DEFAULT_PROXY_URL,
  PROXY_HELPER_DOCS_URL,
  isLoopbackProxyUrl,
  validateOfficeProxyUrl,
} from "../auth/proxy-validation.js";
import { restoreCredentials } from "../auth/restore.js";
import { invalidateBlueprint } from "../context/blueprint.js";
import { ChangeTracker } from "../context/change-tracker.js";
import {
  PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT,
  PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT,
} from "../experiments/events.js";
import { normalizePromptContent } from "../messages/image-input-limits.js";
import { effectiveToolOutputLimits } from "../context/window-budgets.js";
import { readAutoCompactionEnabled } from "../compaction/settings.js";
import { getFilesWorkspace } from "../files/workspace.js";
import { ConnectionManager } from "../connections/manager.js";
import { createAllTools } from "../tools/index.js";
import { collectRequiredConnectionIds } from "../tools/connection-requirements.js";
import {
  applyToolOutputTruncation,
  saveTruncatedToolOutputToWorkspace,
} from "../tools/output-truncation.js";
import { withConnectionPreflight } from "../tools/with-connection-preflight.js";
import {
  migrateLegacyWebSearchApiKeysToConnectionStore,
} from "../tools/web-search-config.js";
import { migrateLegacyMcpTokensToConnectionStore } from "../tools/mcp-config.js";
import {
  applyExperimentalToolGates,
  buildOfficeJsExecuteApprovalMessage,
  buildPythonBridgeApprovalMessage,
} from "../tools/experimental-tool-gates.js";
import {
  buildMutationApprovalMessage,
  withWorkbookCoordinator,
} from "../tools/with-workbook-coordinator.js";
import { registerBuiltins } from "../commands/builtins.js";
import { ExtensionRuntimeManager } from "../extensions/runtime-manager.js";
import type { ResumeDialogTarget } from "../commands/builtins/resume-target.js";
import {
  configureSettingsPages,
  openSettings,
  showResumeDialog,
  type ExtensionsHubTab,
  type RecoveryCheckpointSummary,
} from "../commands/builtins/overlays.js";
import { wireCommandMenu } from "../commands/command-menu.js";
import { executeSlashCommand } from "../commands/slash-command-execution.js";
import {
  getUserRules,
  getWorkbookRules,
  hasAnyRules,
} from "../rules/store.js";
import { createExecutionModeController } from "../execution/controller.js";
import {
  PI_EXECUTION_MODE_CHANGED_EVENT,
  type ExecutionMode,
} from "../execution/mode.js";
import {
  getStoredModelSwitchBehavior,
  setStoredModelSwitchBehavior,
  type ModelSwitchBehavior,
} from "../models/switch-behavior.js";
import { getResolvedConventions } from "../conventions/store.js";
import {
  buildIntegrationPromptEntries,
  createToolsForIntegrations,
  getIntegrationToolNames,
  INTEGRATION_IDS,
} from "../integrations/catalog.js";
import { PI_INTEGRATIONS_CHANGED_EVENT } from "../integrations/events.js";
import { getExternalToolsEnabled, resolveConfiguredIntegrationIds } from "../integrations/store.js";
import { buildSystemPrompt } from "../prompt/system-prompt.js";
import { probeLocalServices, type LocalServiceEntry } from "../tools/bridge-health.js";
import {
  buildAgentSkillPromptEntries,
  listAgentSkills,
  mergeAgentSkillDefinitions,
} from "../skills/catalog.js";
import {
  filterAgentSkillsByEnabledState,
  loadDisabledSkillNamesFromSettings,
} from "../skills/activation-store.js";
import { loadDiscoverableAgentSkillsFromWorkspace } from "../skills/external-store.js";
import { PI_SKILLS_CHANGED_EVENT } from "../skills/events.js";
import { createSkillReadCache } from "../skills/read-cache.js";
import { initLanguage, t } from "../language/index.js";
import { initAppStorage } from "../storage/init-app-storage.js";
import { renderError } from "../ui/loading.js";
import { showFilesWorkspaceDialog } from "../ui/files-dialog.js";

import {
  PI_REQUEST_INPUT_FOCUS_EVENT,
  moveCursorToEnd,
} from "../ui/input-focus.js";
import { requestConfirmationDialog } from "../ui/confirm-dialog.js";
import { TOOL_APPROVAL_OVERLAY_ID } from "../ui/overlay-ids.js";
import { showActionToast, showToast } from "../ui/toast.js";
import { PiSidebar } from "../ui/pi-sidebar.js";
import { createProxyBanner } from "../ui/proxy-banner.js";
import { setActiveProviders } from "../models/active-providers.js";
import { BrowserModelRuntime } from "../models/browser-model-runtime.js";
import {
  ModelRefreshOwner,
  type ModelRefreshRuntime,
} from "../models/model-refresh-owner.js";
import { promptForProviderConnection } from "../ui/api-key-dialog.js";
import { openModelSelectorDialog } from "../ui/model-selector-dialog.js";
import { getCurrentSpreadsheetHost } from "../host/index.js";
import { createWorkbookCoordinator } from "../workbook/coordinator.js";
import { formatWorkbookLabel, type WorkbookContext } from "../workbook/context.js";
import {
  getManualFullWorkbookBackupStore,
  ManualFullWorkbookBackupStore,
  type ManualFullWorkbookBackup,
} from "../workbook/manual-full-backup.js";
import {
  getWorkbookRecoveryLog,
  WorkbookRecoveryLog,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
import {
  getWorkbookChangeAuditLog,
  WorkbookChangeAuditLog,
} from "../audit/workbook-change-audit.js";
import { readRetentionLimit, writeRetentionLimit } from "../workbook/recovery/log-store.js";
import {
  WorkbookSaveBoundaryMonitor,
  startWorkbookSaveBoundaryPolling,
} from "../workbook/save-boundary-monitor.js";

import { createContextInjector } from "./context-injection.js";
import { getThinkingLevels, installKeyboardShortcuts } from "./keyboard-shortcuts.js";
import { createQueueDisplay } from "./queue-display.js";
import { maybeStartBackgroundVerificationBridge } from "./background-verification-bridge.js";
import { RecentlyClosedStack, type RecentlyClosedItem } from "./recently-closed.js";
import {
  loadWorkbookTabLayout,
  saveWorkbookTabLayout,
  type WorkbookTabLayout,
} from "./tab-layout.js";
import { createTabLayoutPersistence } from "./tab-layout-persistence.js";
import { injectStatusBar } from "./status-bar.js";
import {
  parseStatusContextWarningSeverity,
  STATUS_CONTEXT_DESC_ATTR,
  getStatusContextPopoverFallbackDescription,
  STATUS_CONTEXT_TOKENS_ATTR,
  STATUS_CONTEXT_WARNING_ATTR,
  STATUS_CONTEXT_WARNING_SEVERITY_ATTR,
} from "./status-context.js";
import { getProxyState, startProxyPolling } from "./proxy-status.js";
import {
  closeStatusPopover,
  toggleContextPopover,
  toggleThinkingPopover,
} from "./status-popovers.js";
import { showWelcomeLogin } from "./welcome-login.js";
import {
  ensureDefaultProxyUrl,
  readTaskpaneLanguage,
  readTaskpaneProxySettings,
} from "./settings.js";
import {
  SessionRuntimeManager,
  type CreateRuntimeOptions,
  type SessionRuntime,
} from "./session-runtime-manager.js";
import {
  awaitWithTimeout,
  createRuntimeToolFingerprint,
  isLikelyCorsErrorMessage,
  shouldApplyRuntimeToolUpdate,
  isRuntimeAgentTool,
  normalizeRuntimeTools,
} from "./runtime-utils.js";
import { bindRuntimeSidebar } from "./runtime-sidebar-binding.js";
import { doesOverlayClaimEscape } from "../utils/escape-guard.js";

function showErrorBanner(errorRoot: HTMLElement, message: string): void {
  render(renderError(message), errorRoot);
}

function clearErrorBanner(errorRoot: HTMLElement): void {
  render(html``, errorRoot);
}

export async function initTaskpane(opts: {
  appEl: HTMLElement;
  errorRoot: HTMLElement;
  autoResizeImages: boolean;
}): Promise<void> {
  const { appEl, errorRoot, autoResizeImages } = opts;

  const changeTracker = new ChangeTracker();
  const spreadsheetHost = getCurrentSpreadsheetHost();

  // 1. Storage
  const { providerKeys, settings, customProviders, modelCatalogs } = initAppStorage();

  // Initialize language from storage
  initLanguage(await readTaskpaneLanguage(settings));

  // Seed a predictable proxy default for OAuth flows.
  await ensureDefaultProxyUrl(settings, spreadsheetHost.kind);

  // Migrate legacy web-search API keys to the connection store schema.
  try {
    await migrateLegacyWebSearchApiKeysToConnectionStore(settings);
  } catch (error) {
    console.warn("[pi] Failed to migrate legacy web-search API keys:", error);
  }

  // Migrate legacy MCP bearer tokens to the connection store schema.
  try {
    await migrateLegacyMcpTokensToConnectionStore(settings);
  } catch (error) {
    console.warn("[pi] Failed to migrate legacy MCP bearer tokens:", error);
  }

  // 1b. Auto-compaction (Pi defaults to enabled)
  const autoCompactEnabled = await readAutoCompactionEnabled(settings);

  // 1c. Security warning: remote proxies can see your prompts + credentials.
  const initialProxySettings = await readTaskpaneProxySettings(settings);
  if (
    initialProxySettings.enabled &&
    initialProxySettings.url !== null &&
    initialProxySettings.url.length > 0 &&
    !isLoopbackProxyUrl(initialProxySettings.url)
  ) {
    showToast(t("init.securityWarning"));
  }

  const getConfiguredProxyUrl = async (): Promise<string | undefined> => {
    const proxySettings = await readTaskpaneProxySettings(settings);
    if (!proxySettings.enabled) return undefined;

    const candidateUrl = proxySettings.url && proxySettings.url.length > 0
      ? proxySettings.url
      : DEFAULT_PROXY_URL;
    try {
      return validateOfficeProxyUrl(candidateUrl);
    } catch {
      return undefined;
    }
  };

  const modelRuntime = new BrowserModelRuntime({
    providerKeys,
    modelCatalogs,
    getProxyUrl: getConfiguredProxyUrl,
  });

  // 2. Resolve available providers through one owner. Static, custom,
  // dynamically discovered and extension providers share its snapshot.
  let listModelRefreshRuntimes: () => readonly ModelRefreshRuntime[] = () => [];
  const modelRefreshOwner = new ModelRefreshOwner({
    modelRuntime,
    loadCustomProviders: () => customProviders.getAll(),
    getRuntimes: () => listModelRefreshRuntimes(),
    warn: (message, error) => console.warn(message, error),
  });
  modelRefreshOwner.subscribe((snapshot) => {
    setActiveProviders(new Set(snapshot.availableProviders));
    if (snapshot.providerErrors.length > 0) {
      console.warn("[models] Some provider catalogues could not be refreshed:", snapshot.providerErrors);
    }
    document.dispatchEvent(new Event("pi:models-changed"));
  });

  document.addEventListener("pi:providers-changed", () => {
    void modelRefreshOwner.refreshConfiguredProviders(true)
      .catch((error: unknown) => {
        console.warn("[auth] Provider refresh after settings change failed:", error);
      });
  });

  // 2b. Cached catalogues are published before first paint. Remote discovery
  // starts afterward; a late credential restore owns one follow-up refresh.
  const credentialRestorePromise = restoreCredentials(providerKeys, settings);
  try {
    await awaitWithTimeout(
      "Provider lookup",
      9500,
      modelRefreshOwner.startup(credentialRestorePromise, 6000),
    );
  } catch (error) {
    console.warn("[auth] Provider startup failed:", error);
  }

  if (modelRefreshOwner.snapshot().availableProviders.length === 0) {
    void showWelcomeLogin(modelRefreshOwner).catch((error: unknown) => {
      console.warn("[auth] Failed to open welcome login:", error);
    });
  }

  // 3. Change tracker
  changeTracker.start().catch(() => {});

  // 4. Shared runtime dependencies
  // Workbook structure context is injected separately by transformContext.

  const officeStreamFn = createOfficeStreamFn(
    getConfiguredProxyUrl,
    modelRuntime.models,
    (providerId) => modelRuntime.shouldProxyProvider(providerId),
  );
  const streamFn: ReturnType<typeof createOfficeStreamFn> = async (model, transcript, options) => {
    let auth = await modelRuntime.models.getAuth(model);
    if (!auth) {
      if (modelRuntime.isExtensionProvider(model.provider)) {
        throw new Error(t("runtime.error.noApiKey", { provider: model.provider }));
      }
      const success = await promptForProviderConnection(model.provider);
      await modelRefreshOwner.refreshConfiguredProviders(false);
      if (!success) throw new Error(t("runtime.error.noApiKey", { provider: model.provider }));
      auth = await modelRuntime.models.getAuth(model);
    }
    return officeStreamFn(model, transcript, { ...options, ...(auth ? { apiKey: auth.auth.apiKey } : {}) });
  };

  const workbookCoordinator = createWorkbookCoordinator();

  // 5. Create and mount PiSidebar
  const sidebar = new PiSidebar();
  sidebar.emptyHints = [
    {
      label: t("hint.explain.label"),
      prompt: t("hint.explain.prompt"),
    },
    {
      label: t("hint.quality.label"),
      prompt: t("hint.quality.prompt"),
    },
    {
      label: t("hint.financial.label"),
      prompt: t("hint.financial.prompt"),
    },
    {
      label: t("hint.format.label"),
      prompt: t("hint.format.prompt"),
    },
  ];

  appEl.replaceChildren();
  appEl.appendChild(sidebar);

  let rulesActive = false;

  const setRulesActive = (next: boolean) => {
    if (rulesActive === next) return;
    rulesActive = next;
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  };

  const executionModeController = await createExecutionModeController({
    settings,
    showToast,
  });

  const getExecutionMode = (): ExecutionMode => executionModeController.getMode();

  const setExecutionMode = async (mode: ExecutionMode): Promise<void> => {
    await executionModeController.setMode(mode);
  };

  const toggleExecutionModeFromUi = async (): Promise<void> => {
    await executionModeController.toggleFromUi();
  };

  const loadModelSwitchBehavior = async (): Promise<ModelSwitchBehavior> => {
    try {
      return await getStoredModelSwitchBehavior(settings);
    } catch {
      return "inPlace";
    }
  };

  let modelSwitchBehavior: ModelSwitchBehavior = await loadModelSwitchBehavior();

  const getModelSwitchBehavior = (): ModelSwitchBehavior => modelSwitchBehavior;

  const setModelSwitchBehavior = async (nextBehavior: ModelSwitchBehavior): Promise<void> => {
    modelSwitchBehavior = await setStoredModelSwitchBehavior(settings, nextBehavior);
  };

  const resolveWorkbookContext = async (): Promise<WorkbookContext> => {
    try {
      return await spreadsheetHost.getWorkbookContext();
    } catch {
      return {
        workbookId: null,
        workbookName: null,
        source: "unknown",
      };
    }
  };

  const resolveWorkbookId = async (): Promise<string | null> => {
    const workbookContext = await resolveWorkbookContext();
    return workbookContext.workbookId;
  };

  const resolveRuntimeIntegrationIds = async (args: {
    sessionId: string;
    workbookId: string | null;
  }): Promise<string[]> => {
    const configuredIntegrationIds = await resolveConfiguredIntegrationIds({
      settings,
      sessionId: args.sessionId,
      workbookId: args.workbookId,
      knownIntegrationIds: INTEGRATION_IDS,
    });

    const externalToolsEnabled = await getExternalToolsEnabled(settings);
    return externalToolsEnabled ? configuredIntegrationIds : [];
  };

  const resolveAvailableSkills = async () => {
    const bundledSkills = listAgentSkills();

    let mergedSkills = bundledSkills;

    try {
      const discoverableSkills = await loadDiscoverableAgentSkillsFromWorkspace(getFilesWorkspace());
      mergedSkills = mergeAgentSkillDefinitions(bundledSkills, discoverableSkills);
    } catch (error) {
      console.warn("[skills] Failed to load discoverable workspace skills:", error);
    }

    try {
      const disabledSkillNames = await loadDisabledSkillNamesFromSettings(settings);
      const enabledSkills = filterAgentSkillsByEnabledState({
        skills: mergedSkills,
        disabledSkillNames,
      });
      return buildAgentSkillPromptEntries(enabledSkills);
    } catch (error) {
      console.warn("[skills] Failed to load skill activation state:", error);
      return buildAgentSkillPromptEntries(mergedSkills);
    }
  };

  // Probe local bridge services once at init. Snapshot is stable for the session.
  let localServicesSnapshot: LocalServiceEntry[] = [];
  let localServicesReady: Promise<void> = Promise.resolve();

  const buildRuntimeSystemPrompt = async (args: {
    workbookId: string | null;
    activeIntegrationIds: readonly string[];
    requiredConnectionIds: readonly string[];
  }): Promise<string> => {
    await localServicesReady;

    const availableSkills = await resolveAvailableSkills();
    const activeConnections = await connectionManager.listPromptEntries(args.requiredConnectionIds);

    try {
      const userRules = await getUserRules(settings);
      const workbookRules = await getWorkbookRules(settings, args.workbookId);
      setRulesActive(hasAnyRules({ userRules, workbookRules }));
      const conventions = await getResolvedConventions(settings);
      const activeIntegrations = buildIntegrationPromptEntries(args.activeIntegrationIds);
      return buildSystemPrompt({
        userInstructions: userRules,
        workbookInstructions: workbookRules,
        activeIntegrations,
        activeConnections,
        localServices: localServicesSnapshot,
        availableSkills,
        executionMode: getExecutionMode(),
        conventions,
      });
    } catch {
      setRulesActive(false);
      return buildSystemPrompt({
        activeConnections,
        localServices: localServicesSnapshot,
        availableSkills,
        executionMode: getExecutionMode(),
      });
    }
  };

  const activeRuntimeOwner: { current: SessionRuntimeManager | null } = { current: null };
  const getActiveRuntime = (): SessionRuntime | null => activeRuntimeOwner.current?.getActiveRuntime() ?? null;
  const durable = await createDurableTaskpane({
    models: modelRuntime.models,
    streamFn,
    host: spreadsheetHost,
    coordinator: workbookCoordinator,
    defaultModel: modelRefreshOwner.snapshot().defaultModel,
    autoCompaction: autoCompactEnabled,
    allowFaultInjection: import.meta.env.DEV,
    getActiveRuntime,
    onStorageWarning: (message: string) => showToast(message),
  });
  const sessions = durable.sessions;
  const runtimeManager = new SessionRuntimeManager({
    sessions,
    createRuntime: buildSessionRuntime,
    warnCapabilityRefresh: (error) => {
      console.warn("[pi] Failed to refresh runtime capabilities:", error);
    },
  });
  activeRuntimeOwner.current = runtimeManager;
  const recentlyClosed = new RecentlyClosedStack(10);
  const hiddenRuntimes: SessionRuntime[] = [];
  let capabilitiesReady = false;
  let workbookSealed = false;
  const requireWorkbookAssociation = async (): Promise<void> => {
    if (workbookSealed) throw new Error("Workbook changed. Reload Pi before continuing.");
    try {
      const association = await durable.resolveAssociation();
      if (association.scopeId === durable.association.scopeId) return;
    } catch (error) {
      workbookSealed = true;
      capabilitiesReady = false;
      // This guard also runs inside tasks: never join the calling task here.
      void durable.runtime.harness.close(durableContext);
      showErrorBanner(errorRoot, "Workbook changed. Reload Pi before continuing.");
      throw error;
    }
    workbookSealed = true;
    capabilitiesReady = false;
    void durable.runtime.harness.close(durableContext);
    showErrorBanner(errorRoot, "Workbook changed. Reload Pi before continuing.");
    throw new Error("Workbook changed. Reload Pi before continuing.");
  };

  const getActiveConversation = () => getActiveRuntime()?.view ?? null;
  const getActiveQueueDisplay = () => getActiveRuntime()?.queueDisplay ?? null;
  const getActiveLockState = () => getActiveRuntime()?.lockState ?? "idle";

  listModelRefreshRuntimes = () => runtimeManager.listRuntimes().map((runtime) => ({
    runtimeId: runtime.runtimeId,
    model: runtime.view.current.model,
    isBusy: runtime.view.current.busy || runtime.view.current.inbox.length > 0,
    applyModel: (model, thinkingLevel) => {
      void runtime.view.conversation.configure({ model: { provider: model.provider, modelId: model.id },
        ...(thinkingLevel !== undefined ? { thinkingLevel } : {}) }, durableContext).then(() => runtime.view.refresh());
      document.dispatchEvent(new CustomEvent("pi:status-update"));
      if (runtime.runtimeId === getActiveRuntime()?.runtimeId) {
        requestAnimationFrame(() => sidebar.requestUpdate());
      }
    },
  }));

  const workbookRecoveryLog = getWorkbookRecoveryLog(new WorkbookRecoveryLog({ settings }));
  getWorkbookChangeAuditLog(new WorkbookChangeAuditLog({ settings }));
  const manualFullBackupStore = getManualFullWorkbookBackupStore(
    new ManualFullWorkbookBackupStore(),
  );

  const toManualFullBackupSummary = (backup: ManualFullWorkbookBackup) => ({
    id: backup.id,
    createdAt: backup.createdAt,
    sizeBytes: backup.sizeBytes,
  });

  const saveBoundaryMonitor = new WorkbookSaveBoundaryMonitor({
    clearBackupsForWorkbook: (workbookId) => workbookRecoveryLog.clearForWorkbook(workbookId),
  });
  const stopSaveBoundaryPolling = startWorkbookSaveBoundaryPolling({
    monitor: saveBoundaryMonitor,
  });

  window.addEventListener("beforeunload", () => {
    stopSaveBoundaryPolling();
  }, { once: true });

  const restoreCheckpointById = async (snapshotId: string): Promise<void> => {
    const activeRuntime = getActiveRuntime();
    const sessionId = activeRuntime ? String(activeRuntime.view.conversation.id) : crypto.randomUUID();
    const coordinatorWorkbookId = durable.association.scopeId;

    const restored = await workbookCoordinator.runWrite(
      {
        workbookId: coordinatorWorkbookId,
        sessionId,
        opId: crypto.randomUUID(),
        toolName: "workbook_history",
      },
      () => workbookRecoveryLog.restore(snapshotId),
    );

    const address = restored.result.address;
    showToast(t("init.revertedAddress", { address }));
  };

  const toRecoveryCheckpointSummary = (
    snapshot: WorkbookRecoverySnapshot,
  ): RecoveryCheckpointSummary => ({
    id: snapshot.id,
    at: snapshot.at,
    toolName: snapshot.toolName,
    address: snapshot.address,
    changedCount: snapshot.changedCount,
    ...(snapshot.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: snapshot.restoredFromSnapshotId } : {}),
  });

  const formatSessionTitle = (title: string): string => {
    const trimmed = title.trim();
    return trimmed.length > 0 ? trimmed : "Untitled";
  };

  const snapshotRuntimeTabLayout = (): WorkbookTabLayout => {
    const runtimes = runtimeManager.listRuntimes();
    const activeRuntime = runtimeManager.getActiveRuntime();

    return {
      sessionIds: runtimes.map((runtime) => String(runtime.view.conversation.id)),
      activeSessionId: activeRuntime ? String(activeRuntime.view.conversation.id) : null,
    };
  };

  const tabLayoutPersistence = createTabLayoutPersistence({
    resolveWorkbookId,
    saveLayout: async (workbookId, layout) => {
      await saveWorkbookTabLayout(settings, workbookId, layout);
    },
    warn: (message, error) => {
      console.warn(message, error);
    },
  });

  let previousActiveRuntimeId: string | null = null;
  let suppressNextInputAutofocus = false;

  const maybePersistTabLayout = (): void => {
    tabLayoutPersistence.persist(snapshotRuntimeTabLayout());
  };

  const focusChatInput = (): void => {
    if (doesOverlayClaimEscape(document.activeElement)) {
      return;
    }

    const input = sidebar.getInput();
    if (!input) {
      return;
    }

    input.focus();

    const textarea = sidebar.getTextarea();
    if (textarea) {
      moveCursorToEnd(textarea);
    }
  };

  const focusChatInputSoon = (): void => {
    requestAnimationFrame(() => {
      focusChatInput();
    });
  };

  document.addEventListener(PI_REQUEST_INPUT_FOCUS_EVENT, () => {
    focusChatInputSoon();
  });

  bindRuntimeSidebar({ runtimeManager, sidebar });

  runtimeManager.subscribe((snapshot) => {
    if (snapshot.activeRuntimeId !== previousActiveRuntimeId) {
      previousActiveRuntimeId = snapshot.activeRuntimeId;
      document.dispatchEvent(new CustomEvent("pi:active-runtime-changed"));
      if (snapshot.activeRuntimeId && !suppressNextInputAutofocus) {
        focusChatInputSoon();
      }
      suppressNextInputAutofocus = false;
    }

    maybePersistTabLayout();
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  });

  const refreshCapabilitiesForAllRuntimes = async (): Promise<void> => {
    await runtimeManager.refreshCapabilities();
    for (const runtime of hiddenRuntimes) await runtime.refreshCapabilities();
  };

  const reservedToolNames = new Set([
    ...createAllTools({ hostKind: spreadsheetHost.kind }).map((tool) => tool.name),
    ...getIntegrationToolNames(),
  ]);
  const connectionManager = new ConnectionManager({ settings });
  const extensionManager = new ExtensionRuntimeManager({
    settings,
    connectionManager,
    modelRuntime,
    getActiveConversation,
    streamFn,
    refreshRuntimeTools: refreshCapabilitiesForAllRuntimes,
    refreshRuntimeModels: () => modelRefreshOwner.refresh(true),
    reservedToolNames,
    afterInjectAgentContext: async () => {
      const activeRuntime = getActiveRuntime();
      if (!activeRuntime) {
        return;
      }

      await activeRuntime.view.refresh();
      sidebar.requestUpdate();
      document.dispatchEvent(new CustomEvent("pi:status-update"));
    },
  });

  connectionManager.subscribe(() => {
    void refreshCapabilitiesForAllRuntimes();
    void modelRefreshOwner.refresh(true)
      .catch((error: unknown) => {
        console.warn("[models] Failed to refresh after connection change:", error);
      });
  });

  const refreshWorkbookState = async () => {
    try {
      await requireWorkbookAssociation();
    } catch (error) {
      showErrorBanner(errorRoot, error instanceof Error ? error.message : t("init.unknownError"));
      return;
    }
    await resolveWorkbookContext();
    sidebar.requestUpdate();

    await refreshCapabilitiesForAllRuntimes();
    maybePersistTabLayout();
  };

  void refreshWorkbookState();

  window.addEventListener("focus", () => {
    void refreshWorkbookState();
  });

  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") {
      void refreshWorkbookState();
    }
  });

  document.addEventListener("pi:rules-updated", () => {
    void refreshWorkbookState();
  });

  document.addEventListener(PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_INTEGRATIONS_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_SKILLS_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_EXECUTION_MODE_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  // Start bridge health probe now that refreshCapabilitiesForAllRuntimes is available.
  // The first runtime prompt waits for this promise, preventing a first-turn
  // gap where Local Services has not been populated yet.
  localServicesReady = probeLocalServices().then(
    (result) => {
      localServicesSnapshot = result;
      void refreshCapabilitiesForAllRuntimes();
    },
    (error: unknown) => { console.warn("[pi] Local services probe failed:", error); },
  );

  const normalizeApprovalMessage = (title: string, message: string): string => {
    const lines = message.split("\n");
    const firstLine = lines[0]?.trim() ?? "";

    if (firstLine !== title.trim()) {
      return message;
    }

    let startIndex = 1;
    while (startIndex < lines.length && lines[startIndex]?.trim().length === 0) {
      startIndex += 1;
    }

    const body = lines.slice(startIndex).join("\n").trim();
    return body.length > 0 ? body : message;
  };

  const requestRuntimeToolApproval = (args: {
    title: string;
    message: string;
    confirmLabel?: string;
  }): Promise<boolean> => {
    return requestConfirmationDialog({
      overlayId: TOOL_APPROVAL_OVERLAY_ID,
      title: args.title,
      message: normalizeApprovalMessage(args.title, args.message),
      ...(args.confirmLabel !== undefined ? { confirmLabel: args.confirmLabel } : {}),
      cancelLabel: t("confirm.cancel"),
      restoreFocusOnClose: true,
    });
  };

  async function buildSessionRuntime(
    optsForRuntime: CreateRuntimeOptions,
  ): Promise<SessionRuntime> {
    const runtimeId = crypto.randomUUID();
    const initialModel = getActiveRuntime()?.view.current.model ?? modelRefreshOwner.snapshot().defaultModel;
    const latestId = optsForRuntime.autoRestoreLatest ? (await sessions.getAllMetadata())[0]?.id : undefined;
    const requestedId = optsForRuntime.conversationId ?? latestId;
    const conversation = requestedId ? await sessions.resolve(requestedId)
      : await sessions.create(initialModel, initialModel.reasoning ? "high" : "off");
    if (!conversation) throw new Error(`Conversation not found: ${requestedId}`);
    const runtimeSessionId = String(conversation.id);
    const runtimeSkillReadCache = createSkillReadCache();
    const view = await ConversationPresenter.open(durable.runtime.harness, conversation, modelRuntime.models, initialModel);

    const buildRuntimeCapabilities = async (sessionId: string): Promise<DurableCapabilities & {
      extensionToolRevision: number;
    }> => {
      const workbookId = await resolveWorkbookId();
      const activeIntegrationIds = await resolveRuntimeIntegrationIds({
        sessionId,
        workbookId,
      });

      const coreTools = createAllTools({
        hostKind: spreadsheetHost.kind,
        getExtensionManager: () => extensionManager,
        getSessionId: () => runtimeSessionId,
        skillReadCache: runtimeSkillReadCache,
      }).filter(isRuntimeAgentTool);

      const gatedCoreTools = await applyExperimentalToolGates(coreTools, {
        getExecutionMode: () => Promise.resolve(getExecutionMode()),
        requestOfficeJsExecuteApproval: (request) => {
          const apiName = request.apiName ?? "Office.js";
          return requestRuntimeToolApproval({
            title: apiName === "WPS JSAPI"
              ? t("init.confirm.wpsJsTitle")
              : t("init.confirm.officeJsTitle"),
            message: buildOfficeJsExecuteApprovalMessage(request),
            confirmLabel: t("init.confirm.allowOnce"),
          });
        },
        requestPythonBridgeApproval: (request) => {
          return requestRuntimeToolApproval({
            title: t("init.confirm.pythonTitle"),
            message: buildPythonBridgeApprovalMessage(
              request.toolName,
              request.bridgeUrl,
              request.params,
            ),
            confirmLabel: t("init.confirm.allowBridge"),
          });
        },
      });

      // Snapshot extension tools + revision together so schema-stable handler
      // reloads are compared against the same point-in-time payload.
      const extensionTools = extensionManager.getRegisteredTools();
      const extensionToolRevision = extensionManager.getToolRevision();

      const runtimeTools = normalizeRuntimeTools([
        ...gatedCoreTools,
        ...createToolsForIntegrations(activeIntegrationIds),
        ...extensionTools,
      ]);

      const requiredConnectionIds = collectRequiredConnectionIds(runtimeTools);

      const bindMutationGate = (mutationGate: WorkbookMutationGate, mutationId: string): ReturnType<typeof withWorkbookCoordinator> => {
      const coordinatedTools = withWorkbookCoordinator(
        runtimeTools,
        workbookCoordinator,
        {
          getWorkbookId: () => Promise.resolve(durable.association.scopeId),
          getSessionId: () => runtimeSessionId,
        },
        {
          onWriteCommitted: (event) => {
            if (event.impact !== "structure") return;
            invalidateBlueprint(event.workbookId);
            invalidateBlueprint(workbookId);
          },
        },
        {
          mutationGate,
          getMutationId: () => mutationId,
          getExecutionMode: () => Promise.resolve(getExecutionMode()),
          requestMutationApproval: (request) => {
            return requestRuntimeToolApproval({
              title: t("init.confirm.safeModeMutationTitle"),
              message: buildMutationApprovalMessage(request),
              confirmLabel: t("init.confirm.allowOnce"),
            });
          },
        },
      );

      const preflightTools = withConnectionPreflight(coordinatedTools, {
        connectionManager,
      });

      const truncatedTools = applyToolOutputTruncation(preflightTools, {
        // Scale caps with the active model's context window (#566).
        limits: () => effectiveToolOutputLimits(view.current.model.contextWindow),
        saveTruncatedOutput: saveTruncatedToolOutputToWorkspace,
      });
      return truncatedTools.map((tool) => ({ ...tool, execute: async (...args: Parameters<typeof tool.execute>) => {
        await requireWorkbookAssociation();
        return tool.execute(...args);
      } }));
      };
      const tools = bindMutationGate(durable.runtime.mutations, "");

      const systemPrompt = await buildRuntimeSystemPrompt({
        workbookId,
        activeIntegrationIds,
        requiredConnectionIds,
      });

      return {
        tools,
        bindMutationGate,
        systemPrompt,
        extensionToolRevision,
      };
    };

    const initialCapabilities = await buildRuntimeCapabilities(runtimeSessionId);
    const injectContext = createContextInjector(changeTracker);
    const transformContext: typeof injectContext = async (messages, signal) => {
      await requireWorkbookAssociation();
      return injectContext(messages, signal);
    };
    await durable.runtime.install(conversation, initialCapabilities, transformContext);
    let currentRuntimeToolsFingerprint = createRuntimeToolFingerprint(initialCapabilities.tools);
    let currentExtensionToolRevision = initialCapabilities.extensionToolRevision;
    let currentSystemPrompt = initialCapabilities.systemPrompt;

    const refreshRuntimeCapabilities = async () => {
      const next = await buildRuntimeCapabilities(runtimeSessionId);
      const nextToolsFingerprint = createRuntimeToolFingerprint(next.tools);
      const nextExtensionToolRevision = next.extensionToolRevision;

      if (
        shouldApplyRuntimeToolUpdate({
          previousFingerprint: currentRuntimeToolsFingerprint,
          nextFingerprint: nextToolsFingerprint,
          previousExtensionToolRevision: currentExtensionToolRevision,
          nextExtensionToolRevision,
        }) || next.systemPrompt !== currentSystemPrompt
      ) {
        await durable.runtime.install(conversation, next, transformContext);
        runtime.tools = next.tools;
        currentRuntimeToolsFingerprint = nextToolsFingerprint;
        currentExtensionToolRevision = nextExtensionToolRevision;
        currentSystemPrompt = next.systemPrompt;
      }
    };
    const queueDisplay = createQueueDisplay({ view });
    const unsubscribeErrorTracking = view.onEvent((event) => {
      if (event.type === "run_end") window.setTimeout(() => modelRefreshOwner.reconcileRuntimes(), 0);
      if (getActiveRuntime()?.runtimeId !== runtimeId) return;
      if (event.type === "message_start" && event.message.role === "user") clearErrorBanner(errorRoot);
      if (event.type === "task_failed" && !/abort|cancel/i.test(event.message)) {
        showErrorBanner(errorRoot, t("init.llmError", { error: event.message }));
      }
      if (event.type === "message_end") {
        for (const message of event.entry.model ?? []) {
          if (message.role !== "assistant" || !message.errorMessage) continue;
          const error = message.errorMessage;
          if (!/abort|cancel/i.test(error)) showErrorBanner(errorRoot,
            isLikelyCorsErrorMessage(error)
              ? `Network error (likely CORS). If you're using OAuth, enable /settings → Proxy with ${DEFAULT_PROXY_URL} and retry. Guide: ${PROXY_HELPER_DOCS_URL}`
              : t("init.llmError", { error }));
        }
      }
    });
    const metadata = await sessions.metadata(runtimeSessionId);
    const info = await durable.runtime.harness.snapshot(SessionInfoDoc, conversation.id, durableContext);
    const runtime: SessionRuntime = {
      runtimeId,
      view,
      tools: initialCapabilities.tools,
      queueDisplay,
      title: metadata?.title ?? "",
      hasExplicitTitle: info?.explicitTitle ?? false,
      lockState: "idle",
      refreshCapabilities: refreshRuntimeCapabilities,
      dispose: () => {
        runtimeSkillReadCache.clearAll();
        unsubscribeErrorTracking();
        queueDisplay.dispose();
        void view.dispose();
      },
    };
    return runtime;
  }

  const createRuntimeFromUi = async (): Promise<SessionRuntime | null> => {
    try {
      return await runtimeManager.createRuntime({ activate: true, autoRestoreLatest: false });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Unknown error";
      console.warn("[pi] Failed to create a new runtime:", error);
      showToast(t("init.couldNotOpenNewTab", { message }));
      return null;
    }
  };

  let tabLayoutReadSucceeded = false;

  const restorePersistedTabLayout = async (): Promise<SessionRuntime | null> => {
    const workbookId = await resolveWorkbookId();
    let savedLayout: WorkbookTabLayout | null;
    try {
      savedLayout = await loadWorkbookTabLayout(settings, workbookId);
      tabLayoutReadSucceeded = true;
    } catch (error) {
      console.warn("[pi] Failed to read persisted tab layout:", error);
      return null;
    }
    if (!savedLayout) return null;

    const runtimesBySessionId = new Map<string, SessionRuntime>();
    let firstRuntime: SessionRuntime | null = null;

    for (const sessionId of savedLayout.sessionIds) {
      const conversation = await sessions.resolve(sessionId);
      if (!conversation) continue;
      const runtime = await runtimeManager.createRuntime({
        activate: false,
        autoRestoreLatest: false,
        conversationId: String(conversation.id),
      });

      if (!firstRuntime) {
        firstRuntime = runtime;
      }

      if (!runtimesBySessionId.has(sessionId)) {
        runtimesBySessionId.set(sessionId, runtime);
      }
    }

    if (!firstRuntime) return null;

    const preferredActiveRuntime = savedLayout.activeSessionId
      ? runtimesBySessionId.get(savedLayout.activeSessionId) ?? null
      : null;

    const nextActiveRuntime = preferredActiveRuntime ?? firstRuntime;
    runtimeManager.switchRuntime(nextActiveRuntime.runtimeId);
    return nextActiveRuntime;
  };

  const syncRuntimeAfterSessionLoad = (runtime: SessionRuntime): void => {
    void runtime.view.refresh();
    sidebar.requestUpdate();
    document.dispatchEvent(new CustomEvent("pi:model-changed"));
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  };

  const replaceActiveRuntimeSession = async (sessionData: SessionData): Promise<void> => {
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast(t("init.noActiveSession"));
      return;
    }

    const busy = activeRuntime.view.current.busy || activeRuntime.view.current.inbox.length > 0;
    if (busy) {
      showToast(t("init.currentTabBusy"));
      return;
    }

    const replacement = await openSessionInNewTab(sessionData);
    runtimeManager.closeRuntime(activeRuntime.runtimeId);
    syncRuntimeAfterSessionLoad(replacement);
  };

  const openSessionInNewTab = async (sessionData: SessionData): Promise<SessionRuntime> => {
    const runtime = await runtimeManager.createRuntime({
      activate: true,
      autoRestoreLatest: false,
      conversationId: sessionData.id,
    });
    syncRuntimeAfterSessionLoad(runtime);
    return runtime;
  };

  const openResumePicker = async (defaultTarget: ResumeDialogTarget = "new_tab"): Promise<void> => {
    await showResumeDialog({
      sessions,
      defaultTarget,
      onOpenInNewTab: async (sessionData: SessionData) => {
        await openSessionInNewTab(sessionData);
      },
      onReplaceCurrent: async (sessionData: SessionData) => {
        await replaceActiveRuntimeSession(sessionData);
      },
      getRecentlyClosedItems: () => recentlyClosed.snapshot(),
      onReopenRecentlyClosed: async (item) => {
        return reopenRecentlyClosedById(item.id);
      },
    });
  };

  const openRulesEditor = async (): Promise<void> => {
    await openSettings("rules");
  };

  type ReopenRecentlyClosedResult = "reopened" | "missing" | "failed";

  const reopenRecentlyClosedItem = async (item: RecentlyClosedItem): Promise<ReopenRecentlyClosedResult> => {
    try {
      const sessionData = await sessions.loadSession(item.sessionId);
      if (!sessionData) {
        showToast(t("init.couldNotReopenSession"));
        return "missing";
      }

      await openSessionInNewTab(sessionData);
      showToast(t("init.reopened", { title: formatSessionTitle(item.title) }));
      return "reopened";
    } catch {
      showToast(t("init.couldNotReopenSession"));
      return "failed";
    }
  };

  const reopenRecentlyClosedById = async (recentlyClosedId: string): Promise<boolean> => {
    const item = recentlyClosed.removeById(recentlyClosedId);
    if (!item) {
      showToast(t("init.sessionNotInRecentlyClosed"));
      return false;
    }

    const reopenResult = await reopenRecentlyClosedItem(item);
    if (reopenResult === "failed") {
      recentlyClosed.push(item);
    }

    return reopenResult === "reopened";
  };

  const reopenLastClosed = async (): Promise<void> => {
    const item = recentlyClosed.popMostRecent();
    if (!item) {
      showToast(t("init.noRecentlyClosedTab"));
      return;
    }

    const reopenResult = await reopenRecentlyClosedItem(item);
    if (reopenResult === "failed") {
      recentlyClosed.push(item);
    }
  };

  const revertLatestCheckpoint = async (): Promise<void> => {
    const latest = await workbookRecoveryLog.listForCurrentWorkbook(1);
    const checkpoint = latest[0];

    if (!checkpoint) {
      showToast(t("init.noBackupsYet"));
      return;
    }

    await restoreCheckpointById(checkpoint.id);
  };

  const createManualFullBackup = async (): Promise<{
    id: string;
    createdAt: number;
    sizeBytes: number;
  }> => {
    const backup = await manualFullBackupStore.create();
    await getFilesWorkspace().downloadFile(backup.path);
    return toManualFullBackupSummary(backup);
  };

  const listManualFullBackups = async (limit = 5): Promise<Array<{
    id: string;
    createdAt: number;
    sizeBytes: number;
  }>> => {
    const safeLimit = Math.max(1, Math.min(20, Math.floor(limit)));
    const backups = await manualFullBackupStore.listForCurrentWorkbook(safeLimit);
    return backups.map((backup) => toManualFullBackupSummary(backup));
  };

  const restoreManualFullBackup = async (
    backupId?: string,
  ): Promise<{
    id: string;
    createdAt: number;
    sizeBytes: number;
  } | null> => {
    const resolved = backupId?.trim()
      ? await manualFullBackupStore.downloadByIdForCurrentWorkbook(backupId)
      : await manualFullBackupStore.downloadLatestForCurrentWorkbook();

    return resolved ? toManualFullBackupSummary(resolved) : null;
  };

  const clearManualFullBackups = async (): Promise<number> => {
    return manualFullBackupStore.clearForCurrentWorkbook();
  };

  const closeRuntimeWithRecovery = async (
    runtimeId: string,
    optsForClose?: { showUndoToast?: boolean },
  ): Promise<boolean> => {
    if (runtimeManager.listRuntimes().length <= 1) {
      showToast(t("init.cantCloseLastTab"));
      return false;
    }

    const runtime = runtimeManager.getRuntime(runtimeId);
    if (!runtime) return false;

    if (runtime.lockState === "holding_lock") {
      showToast(t("init.waitForChangesBeforeClose"));
      return false;
    }

    const closeTitle = runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.title
      ?? formatSessionTitle(runtime.title);

    const closedItem: RecentlyClosedItem = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`,
      sessionId: String(runtime.view.conversation.id),
      title: closeTitle,
      closedAt: new Date().toISOString(),
      workbookId: await resolveWorkbookId(),
    };

    runtimeManager.closeRuntime(runtimeId);
    recentlyClosed.push(closedItem);

    const showUndoToast = optsForClose?.showUndoToast !== false;
    if (showUndoToast) {
      showActionToast({
        message: t("init.closedTab", { title: closedItem.title }),
        actionLabel: t("init.undo"),
        duration: 9000,
        onAction: () => {
          void reopenRecentlyClosedById(closedItem.id);
        },
      });
    }

    return true;
  };

  const renameRuntimeTab = async (runtimeId: string): Promise<void> => {
    const runtime = runtimeManager.getRuntime(runtimeId);
    if (!runtime) {
      showToast(t("init.sessionNotFound"));
      return;
    }

    if (runtime.view.current.busy || runtime.view.current.inbox.length > 0) {
      showToast(t("init.waitBeforeRenaming"));
      return;
    }

    const currentTitle = runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.title
      ?? formatSessionTitle(runtime.title);
    const defaultTitle = runtime.hasExplicitTitle
      ? runtime.title.trim()
      : currentTitle;

    const nextTitleRaw = window.prompt("Rename tab", defaultTitle);
    if (nextTitleRaw === null) {
      return;
    }

    const nextTitle = nextTitleRaw.trim();
    await runtimeManager.renameRuntime(runtimeId, nextTitle);
    sidebar.requestUpdate();
    document.dispatchEvent(new CustomEvent("pi:status-update"));

    if (nextTitle.length === 0) {
      showToast(t("init.tabNameReset"));
      return;
    }

    showToast(t("init.renamedTo", { title: nextTitle }));
  };

  const resolveRuntimeTabTitle = (runtimeId: string, runtime: SessionRuntime): string => {
    return runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.title
      ?? formatSessionTitle(runtime.title);
  };

  type RuntimeModel = Model<Api>;

  const cloneRuntimeToNewTab = async (args: {
    sourceRuntime: SessionRuntime;
    targetModel: RuntimeModel;
    targetTitle: string;
  }): Promise<SessionRuntime> => {
    const fork = await sessions.fork(String(args.sourceRuntime.view.conversation.id), args.targetModel, args.targetTitle);
    const clonedRuntime = await runtimeManager.createRuntime({
      activate: true,
      autoRestoreLatest: false,
      conversationId: String(fork.id),
    });
    sidebar.requestUpdate();
    document.dispatchEvent(new CustomEvent("pi:model-changed"));
    document.dispatchEvent(new CustomEvent("pi:status-update"));

    return clonedRuntime;
  };

  const duplicateRuntimeTab = async (runtimeId: string): Promise<void> => {
    const sourceRuntime = runtimeManager.getRuntime(runtimeId);
    if (!sourceRuntime) {
      showToast(t("init.sessionNotFound"));
      return;
    }

    if (sourceRuntime.view.current.busy || sourceRuntime.view.current.inbox.length > 0) {
      showToast(t("init.waitBeforeDuplicating"));
      return;
    }

    const sourceTitle = resolveRuntimeTabTitle(runtimeId, sourceRuntime);
    const duplicateTitle = `${sourceTitle} copy`;

    await cloneRuntimeToNewTab({
      sourceRuntime,
      targetModel: sourceRuntime.view.current.model,
      targetTitle: duplicateTitle,
    });

    showToast(t("init.duplicated", { title: sourceTitle }));
  };

  const closeOtherRuntimes = async (runtimeId: string): Promise<void> => {
    const tabsToClose = runtimeManager.snapshotTabs()
      .filter((tab) => tab.runtimeId !== runtimeId)
      .map((tab) => tab.runtimeId);

    if (tabsToClose.length === 0) {
      showToast(t("init.noOtherTabs"));
      return;
    }

    let closedCount = 0;
    for (const tabId of tabsToClose) {
      const closed = await closeRuntimeWithRecovery(tabId, { showUndoToast: false });
      if (closed) {
        closedCount += 1;
      }
    }

    runtimeManager.switchRuntime(runtimeId);

    if (closedCount === 0) {
      showToast(t("init.noTabsWereClosed"));
      return;
    }

    showToast(t("init.closedOtherTabs", { count: closedCount }));
  };

  const moveRuntimeTab = (runtimeId: string, direction: -1 | 1): void => {
    const moved = runtimeManager.moveRuntime(runtimeId, direction);
    if (!moved) return;

    const directionLabel = direction < 0 ? "left" : "right";
    showToast(t("init.movedTab", { direction: directionLabel }), 1200);
  };

  workbookCoordinator.subscribe((event) => {
    if (event.operationType !== "write") return;

    const runtime = runtimeManager.findRuntimeBySessionId(event.context.sessionId);
    if (!runtime) return;

    if (event.type === "queued") {
      runtimeManager.setRuntimeLockState(runtime.runtimeId, "waiting_for_lock");
      return;
    }

    if (event.type === "started") {
      runtimeManager.setRuntimeLockState(runtime.runtimeId, "holding_lock");
      return;
    }

    if (event.type === "completed" || event.type === "failed") {
      runtimeManager.setRuntimeLockState(runtime.runtimeId, "idle");
    }
  });

  const openExtensionsHub = (tab?: ExtensionsHubTab): Promise<void> => {
    return openSettings(tab ?? "connections");
  };

  const openRecoveryDialog = async (): Promise<void> => {
    await openSettings("backups");
  };

  configureSettingsPages({
    getExecutionMode,
    setExecutionMode,
    getModelSwitchBehavior,
    setModelSwitchBehavior,
    models: modelRuntime.models,
    onRulesSaved: async () => {
      await refreshWorkbookState();
    },
    extensionsHub: {
      getActiveSessionId: () => {
        const runtime = getActiveRuntime();
        return runtime ? String(runtime.view.conversation.id) : null;
      },
      resolveWorkbookContext: async () => {
        const workbookContext = await resolveWorkbookContext();
        return {
          workbookId: workbookContext.workbookId,
          workbookLabel: formatWorkbookLabel(workbookContext),
        };
      },
      extensionManager,
      connectionManager,
      onChanged: refreshCapabilitiesForAllRuntimes,
    },
    backups: {
      loadCheckpoints: async () => {
        const checkpoints = await workbookRecoveryLog.listForCurrentWorkbook(40);
        return checkpoints.map((checkpoint) => toRecoveryCheckpointSummary(checkpoint));
      },
      onRestore: async (snapshotId: string) => {
        await restoreCheckpointById(snapshotId);
      },
      onDelete: async (snapshotId: string) => {
        return workbookRecoveryLog.delete(snapshotId);
      },
      onClear: async () => {
        return workbookRecoveryLog.clearForCurrentWorkbook();
      },
      onCreateManualFullBackup: async () => {
        return createManualFullBackup();
      },
      getRetentionConfig: async () => {
        const maxSnapshots = await readRetentionLimit();
        return { maxSnapshots };
      },
      setRetentionConfig: async (config) => {
        await writeRetentionLimit(config.maxSnapshots);
      },
    },
  });

  const applyModelSelection = async (runtimeId: string, nextModel: RuntimeModel): Promise<void> => {
    const result = await runtimeManager.selectModel({
      runtimeId,
      nextModel,
      behavior: getModelSwitchBehavior(),
    });

    if (result.outcome === "missing") {
      showToast(t("init.sessionNotFound"));
      return;
    }
    if (result.outcome === "busy") {
      showToast(t("init.waitBeforeChangingModels"));
      return;
    }
    if (result.outcome === "unchanged") return;

    if (result.outcome === "forked") {
      sidebar.requestUpdate();
      showToast(t("init.openedInNewTab", { title: result.title }));
    }
    document.dispatchEvent(new CustomEvent("pi:model-changed"));
    document.dispatchEvent(new CustomEvent("pi:status-update"));
    requestAnimationFrame(() => sidebar.requestUpdate());
  };

  const openModelSelector = (): void => {
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast(t("init.noActiveSession"));
      return;
    }

    const targetRuntimeId = activeRuntime.runtimeId;
    const currentModel = activeRuntime.view.current.model;

    void (async () => {
      try {
        await modelRefreshOwner.refreshConfiguredProviders(false);
      } catch (error) {
        console.warn("[auth] Failed to refresh providers before opening model selector:", error);
      }

      closeStatusPopover();

      openModelSelectorDialog({
        models: modelRuntime.models,
        currentModel,
        onSelect: (model) => {
          void applyModelSelection(targetRuntimeId, model);
        },
      });

      void modelRefreshOwner.refresh(true).catch((error: unknown) => {
        console.warn("[models] Model refresh from selector failed:", error);
      });
    })();
  };

  registerBuiltins({
    getActiveConversation,
    getSessions: () => sessions,
    openModelSelector,
    renameActiveSession: async (title: string) => {
      const activeRuntime = getActiveRuntime();
      if (!activeRuntime) {
        showToast(t("init.noActiveSession"));
        return;
      }

      await runtimeManager.renameRuntime(activeRuntime.runtimeId, title);
    },
    createRuntime: async () => {
      await createRuntimeFromUi();
    },
    openResumeDialog: async (defaultTarget: ResumeDialogTarget = "new_tab") => {
      await openResumePicker(defaultTarget);
    },
    openRecoveryDialog,
    reopenLastClosed,
    revertLatestCheckpoint: async () => {
      try {
        await revertLatestCheckpoint();
      } catch (error) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(t("init.revertFailed", { message }));
      }
    },
    createManualFullBackup: async () => {
      return createManualFullBackup();
    },
    listManualFullBackups: async (limit?: number) => {
      return listManualFullBackups(limit);
    },
    restoreManualFullBackup: async (backupId?: string) => {
      return restoreManualFullBackup(backupId);
    },
    clearManualFullBackups: async () => {
      return clearManualFullBackups();
    },
    openInstructionsEditor: async () => {
      await openRulesEditor();
    },
    getExecutionMode: () => Promise.resolve(getExecutionMode()),
    setExecutionMode,
    openExtensionsHub,
    openFilesWorkspace: () => showFilesWorkspaceDialog(),
  });

  // Slash commands chosen from the popup menu dispatch this event.
  const onCommandRun: EventListener = (event) => {
    if (!capabilitiesReady) return;
    if (!(event instanceof CustomEvent)) return;
    if (!isTaskpaneInitPayloadShape(event.detail)) return;

    const name = typeof event.detail.name === "string" ? event.detail.name : "";
    const args = typeof event.detail.args === "string" ? event.detail.args : "";
    if (!name) return;

    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast(t("init.noActiveSession"));
      return;
    }

    const busy = activeRuntime.view.current.busy || activeRuntime.view.current.inbox.length > 0;
    if (name === "compact") {
      void activeRuntime.view.conversation.compact(args.trim() || undefined, durableContext)
        .catch(() => showToast(t("slash-command.toast.failed")));
      return;
    }
    const result = executeSlashCommand({
      name,
      args,
      busy,
      onError: () => {
        showToast(t("slash-command.toast.failed"));
      },
    });

    if (result === "not-found") {
      return;
    }

    if (result === "busy-blocked") {
      showToast(t("init.cantRunBusy", { command: name }));
      return;
    }

    if (result === "missing-queue") {
      showToast(t("init.noActiveSession"));
    }
  };
  document.addEventListener("pi:command-run", onCommandRun);

  sidebar.onSend = (text) => {
    if (!capabilitiesReady) return;
    clearErrorBanner(errorRoot);
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast(t("init.noActiveSession"));
      return;
    }
    void (async () => {
      await requireWorkbookAssociation();
      const content = await normalizePromptContent(text, activeRuntime.view.current.model, { autoResizeImages });
      await activeRuntime.view.conversation.submit({ type: "input", content, requestId: crypto.randomUUID() }, durableContext);
    })().catch((error: unknown) => showErrorBanner(errorRoot, error instanceof Error ? error.message : t("init.unknownError")));
  };

  sidebar.onAbort = () => {
    if (!capabilitiesReady) return;
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) return;
    void activeRuntime.view.conversation.abort(durableContext);
  };

  sidebar.onCreateTab = () => {
    void createRuntimeFromUi();
  };

  sidebar.onSelectTab = (runtimeId: string) => {
    runtimeManager.switchRuntime(runtimeId);
  };

  sidebar.onCloseTab = (runtimeId: string) => {
    void closeRuntimeWithRecovery(runtimeId);
  };
  sidebar.onRenameTab = (runtimeId: string) => {
    void renameRuntimeTab(runtimeId);
  };
  sidebar.onDuplicateTab = (runtimeId: string) => {
    void duplicateRuntimeTab(runtimeId);
  };
  sidebar.onMoveTabLeft = (runtimeId: string) => {
    moveRuntimeTab(runtimeId, -1);
  };
  sidebar.onMoveTabRight = (runtimeId: string) => {
    moveRuntimeTab(runtimeId, 1);
  };
  sidebar.onCloseOtherTabs = (runtimeId: string) => {
    void closeOtherRuntimes(runtimeId);
  };
  sidebar.onOpenRules = () => {
    void openRulesEditor();
  };
  sidebar.onOpenExtensions = () => {
    void openExtensionsHub();
  };
  sidebar.onOpenSettings = () => {
    void openSettings();
  };
  sidebar.onOpenFilesWorkspace = () => {
    void showFilesWorkspaceDialog();
  };
  sidebar.onFilesDrop = (files: File[]) => {
    const workspace = getFilesWorkspace();

    void workspace.importFiles(files, {
      audit: { actor: "user", source: "input-drop" },
    })
      .then((count) => {
        if (count <= 0) {
          showToast(t("init.noFilesImported"));
          return;
        }

        const importedLabel = `${count} file${count === 1 ? "" : "s"}`;
        showToast(t("init.importedIntoFiles", { label: importedLabel }));
      })
      .catch((error: unknown) => {
        showToast(t("init.importFailed", { error: error instanceof Error ? error.message : t("init.unknownError") }));
      });
  };
  sidebar.onOpenResumePicker = () => {
    void openResumePicker("new_tab");
  };
  sidebar.onOpenRecovery = () => {
    void openRecoveryDialog();
  };
  sidebar.onOpenShortcuts = () => {
    void openSettings("shortcuts");
  };


  // Bootstrap from persisted tab layout; fallback to legacy single-runtime restore.
  const restoredRuntime = await restorePersistedTabLayout();
  if (!restoredRuntime) {
    await runtimeManager.createRuntime({ activate: true, autoRestoreLatest: true });
  }

  // Install every pending conversation's capabilities before enabling scheduling.
  // Closed tabs may still own work; their projections are not added to the tab strip.
  const inspection = await durable.runtime.harness.inspect(durableContext);
  const pendingIds = new Set([
    ...inspection.tasks.map((task) => String(task.record.conversationId)),
    ...inspection.submissions.map((submission) => String(submission.conversationId)),
  ]);
  for (const conversationId of pendingIds) {
    if (runtimeManager.findRuntimeBySessionId(conversationId)) continue;
    hiddenRuntimes.push(await buildSessionRuntime({ activate: false, autoRestoreLatest: false, conversationId }));
  }

  const startupTabLayout = snapshotRuntimeTabLayout();
  tabLayoutPersistence.enable(tabLayoutReadSucceeded ? undefined : startupTabLayout);
  tabLayoutPersistence.persist(startupTabLayout);

  // ── Disclosure bar (onboarding banner) ──
  // Must run after runtime bootstrap so the sidebar has rendered .pi-messages.
  {
    const { createDisclosureBar } = await import("../ui/disclosure-bar.js");
    const disclosureEl = createDisclosureBar({
      providerCount: modelRefreshOwner.snapshot().availableProviders.length,
      onOpenSettings: () => void openSettings(),
    });
    if (disclosureEl) {
      const messagesContainer = sidebar.querySelector(".pi-messages");
      if (messagesContainer) {
        messagesContainer.parentElement?.insertBefore(disclosureEl, messagesContainer);
      }
    }
  }

  // ── Proxy warning banner (state-driven) ──
  {
    const messagesContainer = sidebar.querySelector(".pi-messages");
    const host = messagesContainer?.parentElement;
    if (host && messagesContainer) {
      const proxyBanner = createProxyBanner();
      host.insertBefore(proxyBanner.root, messagesContainer);
      proxyBanner.update(getProxyState());

      document.addEventListener("pi:proxy-state-changed", (event: Event) => {
        if (!(event instanceof CustomEvent)) return;
        const detail: unknown = event.detail;
        if (!isTaskpaneInitPayloadShape(detail)) return;
        const state = detail.state;
        if (state === "detected" || state === "not-detected" || state === "unknown") {
          proxyBanner.update(state);
        }
      });
    }
  }

  // ── Register extensions ──
  // Never let extension startup block core UI wiring (status bar + slash menu).
  // Some extension sources can be slow/unreachable; continue boot even if this
  // is still in-flight, and let extensions finish loading in the background.
  const extensionInitialization = extensionManager.initialize();

  void extensionInitialization.catch((error: unknown) => {
    console.warn("[pi] Extension initialization failed:", error);
  });

  try {
    await awaitWithTimeout("Extension initialization", 5000, extensionInitialization);
  } catch (error) {
    console.warn("[pi] Extension initialization did not complete during startup:", error);
  }

  await refreshCapabilitiesForAllRuntimes();
  await requireWorkbookAssociation();
  await durable.runtime.acknowledgeCommittedResults();
  durable.runtime.harness.resume();
  capabilitiesReady = true;

  // ── Keyboard shortcuts ──
  installKeyboardShortcuts({
    getActiveConversation,
    getActiveQueueDisplay,
    sidebar,
    onCreateTab: () => {
      void createRuntimeFromUi();
    },
    onCloseActiveTab: () => {
      const activeRuntime = getActiveRuntime();
      if (!activeRuntime) {
        showToast(t("init.noActiveSession"));
        return;
      }

      void closeRuntimeWithRecovery(activeRuntime.runtimeId);
    },
    onReopenLastClosed: () => {
      void reopenLastClosed();
    },
    canUndoCloseTab: () => recentlyClosed.peekMostRecent() !== null,
    onSwitchAdjacentTab: (direction: -1 | 1) => {
      const tabs = runtimeManager.snapshotTabs();
      if (tabs.length <= 1) {
        return;
      }

      const activeIndex = tabs.findIndex((tab) => tab.isActive);
      if (activeIndex < 0) {
        return;
      }

      const nextIndex = (activeIndex + direction + tabs.length) % tabs.length;

      suppressNextInputAutofocus = true;
      const nextTab = tabs[nextIndex];
      if (!nextTab) {
        return;
      }

      runtimeManager.switchRuntime(nextTab.runtimeId);
      requestAnimationFrame(() => {
        sidebar.focusTabNavigationAnchor();
      });
    },
  });

  // ── Status bar ──
  injectStatusBar({
    getActiveConversation,
    getLockState: getActiveLockState,
    getExecutionMode,
  });

  // ── Proxy status polling ──
  startProxyPolling(settings);

  // ── Wire command menu to textarea ──
  const wireTextarea = () => {
    const ta = sidebar.getTextarea();
    if (ta) {
      wireCommandMenu(ta);
    } else {
      requestAnimationFrame(wireTextarea);
    }
  };
  requestAnimationFrame(wireTextarea);

  const runSlashCommand = (name: string, args = ""): void => {
    document.dispatchEvent(new CustomEvent("pi:command-run", { detail: { name, args } }));
  };

  const openThinkingPopoverFrom = (target: Element): void => {
    const trigger = target.closest(".pi-status-thinking");
    if (!trigger) return;

    const activeAgent = getActiveConversation();
    if (!activeAgent) {
      showToast(t("init.noActiveSession"));
      return;
    }

    const description = trigger.getAttribute("data-tooltip") ?? "Choose how deeply Pi reasons before responding.";

    toggleThinkingPopover({
      anchor: trigger,
      description,
      levels: getThinkingLevels(activeAgent),
      activeLevel: activeAgent.current.thinkingLevel,
      onSelectLevel: (level) => {
        if (activeAgent.current.thinkingLevel === level) return;
        void activeAgent.conversation.configure({ thinkingLevel: level }, durableContext).then(() => activeAgent.refresh());
        document.dispatchEvent(new CustomEvent("pi:status-update"));
      },
    });
  };

  const openContextPopoverFrom = (target: Element): void => {
    const trigger = target.closest(".pi-status-ctx--trigger");
    if (!trigger) return;

    const description = trigger.getAttribute(STATUS_CONTEXT_DESC_ATTR)
      ?? getStatusContextPopoverFallbackDescription();

    const tokenDetail = trigger.getAttribute(STATUS_CONTEXT_TOKENS_ATTR) ?? undefined;

    const warnText = trigger.getAttribute(STATUS_CONTEXT_WARNING_ATTR) ?? "";
    const warnSeverity = trigger.getAttribute(STATUS_CONTEXT_WARNING_SEVERITY_ATTR);
    let warning: { text: string; severity: "yellow" | "red" } | undefined;
    if (warnText.length > 0) {
      warning = {
        text: warnText,
        severity: parseStatusContextWarningSeverity(warnSeverity),
      };
    }

    toggleContextPopover({
      anchor: trigger,
      description,
      ...(tokenDetail !== undefined ? { tokenDetail } : {}),
      ...(warning !== undefined ? { warning } : {}),
      onRunCommand: (command) => {
        runSlashCommand(command);
      },
    });
  };

  // ── Dev-only background verification bridge ──
  const backgroundVerificationBridge = maybeStartBackgroundVerificationBridge({
    sidebar,
    getWorkbookContext: resolveWorkbookContext,
    getActiveRuntime,
    extensionManager,
    connectionManager,
    modelRuntime,
    durable,
    refreshModels: async (allowNetwork) => {
      await modelRefreshOwner.refresh(allowNetwork);
      return { errors: [...modelRefreshOwner.snapshot().providerErrors] };
    },
  });
  if (backgroundVerificationBridge) {
    window.addEventListener("pagehide", () => backgroundVerificationBridge.stop(), { once: true });
  }

  // ── Status bar click handlers ──
  document.addEventListener("click", (e) => {
    const target = e.target;
    if (!(target instanceof Element)) return;

    const el = target;

    if (el.closest(".pi-status-popover")) {
      return;
    }

    // Model picker
    if (el.closest(".pi-status-model")) {
      openModelSelector();
      return;
    }

    // Execution mode toggle
    if (el.closest(".pi-status-mode")) {
      closeStatusPopover();
      void toggleExecutionModeFromUi();
      return;
    }

    // Context quick actions
    if (el.closest(".pi-status-ctx--trigger")) {
      openContextPopoverFrom(el);
      return;
    }

    // Thinking level selector
    if (el.closest(".pi-status-thinking")) {
      openThinkingPopoverFrom(el);
      return;
    }

    closeStatusPopover();
  });

  console.log("[pi] PiSidebar mounted");
}
