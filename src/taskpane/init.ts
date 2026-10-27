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
import { Agent } from "@earendil-works/pi-agent-core";
import { clampThinkingLevel, getCurrentSystemMessage, toToolDeclaration } from "@earendil-works/pi-ai";
import { getAppStorage } from "../storage/local/app-storage.js";
import type { SessionData } from "../storage/local/types.js";

import { createOfficeStreamFn } from "../auth/stream-proxy.js";
import {
  DEFAULT_PROXY_URL,
  validateOfficeProxyUrl,
} from "../auth/proxy-validation.js";
import { restoreCredentials } from "../auth/restore.js";
import { createOAuthRefresher } from "../auth/oauth-refresh.js";
import { getOAuthProvider } from "../auth/oauth-provider-registry.js";
import { invalidateBlueprint } from "../context/blueprint.js";
import { onCellEditWait } from "../excel/helpers.js";
import { ChangeTracker, type PiWriteWindow } from "../context/change-tracker.js";
import {
  PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT,
  PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT,
} from "../experiments/events.js";
import { createConvertToLlm } from "../messages/convert-to-llm.js";
import { installImageInputNormalization } from "../messages/image-input-limits.js";
import { effectiveToolOutputLimits } from "../context/window-budgets.js";
import { findTrailingContextOverflowError } from "../compaction/overflow-recovery.js";
import { createTabState } from "./tab-state.js";
import { getErrorMessage } from "../utils/errors.js";
import { describeTurnFailure } from "./turn-error.js";
import { createDraftStore } from "./draft-store.js";
import { submitComposerText } from "./keyboard-shortcuts/editor-actions.js";
import { readAutoCompactionEnabled } from "../compaction/settings.js";
import { runCompactCommand } from "../commands/builtins/export.js";
import { describeImportRenames, getFilesWorkspace } from "../files/workspace.js";
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
import { requestTextInputDialog } from "../ui/text-input-dialog.js";
import { TOOL_APPROVAL_OVERLAY_ID } from "../ui/overlay-ids.js";
import { showActionToast, showToast, type ActionToastHandle } from "../ui/toast.js";
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
import { createWorkbookCoordinator, getSessionLockState } from "../workbook/coordinator.js";
import { formatWorkbookLabel, type WorkbookContext } from "../workbook/context.js";
import { createRecoveryScopeResolver } from "../workbook/recovery-scope.js";
import {
  getManualFullWorkbookBackupStore,
  ManualFullWorkbookBackupStore,
  type ManualFullWorkbookBackup,
} from "../workbook/manual-full-backup.js";
import {
  getWorkbookRecoveryLog,
  WorkbookRecoveryLog,
  type WorkbookRecoveryRestorePreview,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
import { formatRelativeDate } from "../commands/builtins/overlay-relative-date.js";
import {
  getWorkbookChangeAuditLog,
  WorkbookChangeAuditLog,
} from "../audit/workbook-change-audit.js";
import { readRetentionLimit, writeRetentionLimit } from "../workbook/recovery/log-store.js";
import { MAX_RECOVERY_ENTRIES } from "../workbook/recovery/constants.js";
import { PI_BACKUP_SAVE_FAILED_EVENT } from "../tools/mutation/finalize.js";
import { requestApproval } from "../ui/approval-dialog.js";
import { humanizeToolInput } from "../ui/humanize-params.js";
import { describePendingToolCall } from "../ui/tool-renderers.js";
import {
  BackupIdentityMonitor,
  startBackupIdentityPolling,
} from "../workbook/backup-identity-monitor.js";

import { createAutoContext } from "./context-injection.js";
import { getThinkingLevels, installKeyboardShortcuts } from "./keyboard-shortcuts.js";
import { createQueueDisplay } from "./queue-display.js";
import { createActionQueue } from "./action-queue.js";
import { maybeStartBackgroundVerificationBridge } from "./background-verification-bridge.js";
import {
  RecentlyClosedStack,
  loadRecentlyClosed,
  saveRecentlyClosed,
  type RecentlyClosedItem,
} from "./recently-closed.js";
import { setupSessionPersistence } from "./sessions.js";
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
  needsRemoteProxyWarning,
  readTaskpaneLanguage,
  readTaskpaneProxySettings,
} from "./settings.js";
import {
  SessionRuntimeManager,
  isRuntimeBusy,
  type CreateRuntimeOptions,
  type OpenSessionTarget,
  type SessionRuntime,
} from "./session-runtime-manager.js";
import {
  awaitWithTimeout,
  createRuntimeToolFingerprint,
  shouldApplyRuntimeToolUpdate,
  isRuntimeAgentTool,
  normalizeRuntimeTools,
} from "./runtime-utils.js";
import { bindRuntimeSidebar } from "./runtime-sidebar-binding.js";
import { doesOverlayClaimEscape } from "../utils/escape-guard.js";

const RUNTIME_SYSTEM_PROMPT_SECTION = "pi-for-excel-runtime";

function renderErrorBanner(errorRoot: HTMLElement, message: string | null): void {
  render(message ? renderError(message) : html``, errorRoot);
}

export async function initTaskpane(opts: {
  appEl: HTMLElement;
  errorRoot: HTMLElement;
  autoResizeImages: boolean;
}): Promise<void> {
  const { appEl, errorRoot, autoResizeImages } = opts;

  // One workbook-level edit log; each runtime reads it from its own cursor.
  // User structure changes (sheets, rows, columns) invalidate the blueprint.
  const changeTracker = new ChangeTracker({
    onStructureChange: () => invalidateBlueprint(null),
    getPiWriteAddresses: async (toolCallIds) => {
      const wanted = new Set(toolCallIds);
      const written = new Map<string, string[]>();
      for (const entry of await getWorkbookChangeAuditLog().list(200)) {
        if (!wanted.has(entry.toolCallId)) continue;
        const address = entry.outputAddress ?? entry.inputAddress;
        if (address) written.set(entry.toolCallId, [...(written.get(entry.toolCallId) ?? []), address]);
      }
      return written;
    },
  });
  const spreadsheetHost = getCurrentSpreadsheetHost();

  // 1. Storage
  const { providerKeys, sessions, settings, customProviders, modelCatalogs } = initAppStorage();

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
  // The proxy pages warn again when a remote URL is saved or enabled.
  if (needsRemoteProxyWarning(await readTaskpaneProxySettings(settings))) {
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
    // The model control tags the provider when two connected providers share models.
    document.dispatchEvent(new CustomEvent("pi:status-update"));
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
  const oauthRefresher = createOAuthRefresher({ settings, providerKeys, getOAuthProvider });
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

  const streamFn = createOfficeStreamFn(
    getConfiguredProxyUrl,
    modelRuntime.models,
    (providerId) => modelRuntime.providerRoute(providerId),
  );

  const workbookCoordinator = createWorkbookCoordinator();

  // 5. Create and mount PiSidebar
  const sidebar = new PiSidebar();
  // Excel rejects every batch while the user is typing in a cell; excelRun
  // retries and this tells the user why Pi is waiting.
  onCellEditWait((waiting) => {
    sidebar.setHostWaitIndicator(
      waiting ? t("working.cellEdit") : null,
      waiting ? t("working.cellEditHint") : null,
    );
  });
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

  const runtimeManager = new SessionRuntimeManager({
    createRuntime: buildSessionRuntime,
    warnCapabilityRefresh: (error) => {
      console.warn("[pi] Failed to refresh runtime capabilities:", error);
    },
  });
  const draftStore = createDraftStore(settings);

  // Per-workbook tab state (layout, recently closed) is stored under the
  // workbook's identity. A workbook without one (never saved, or edited through
  // an anonymous web link) uses the token Pi keeps inside the document, as
  // recovery checkpoints do, so a reload keeps its tabs and another
  // unidentified workbook does not see them. Without a token the state stays in
  // memory. The standalone browser taskpane never has an identity and keeps
  // using the shared slot.
  const usesSharedUnidentifiedState = spreadsheetHost.kind === "browser";
  const workbookScope = createRecoveryScopeResolver({ getWorkbookContext: resolveWorkbookContext });
  const resolvePersistentWorkbookKey = async (
    mode: "read" | "write",
  ): Promise<{ workbookId: string | null } | null> => {
    if (usesSharedUnidentifiedState) return { workbookId: await resolveWorkbookId() };
    const scope = mode === "write" ? await workbookScope.resolveForAppend() : await workbookScope.resolveForRead();
    return scope ? { workbookId: scope.workbookId } : null;
  };

  // Recently closed tabs are kept per workbook. If the stored list cannot be
  // read, it is not overwritten: closes stay in memory for this taskpane only.
  let recentlyClosedWritable = false;
  let recentlyClosedWrites: Promise<void> = Promise.resolve();
  const recentlyClosed = new RecentlyClosedStack(10, (items) => {
    if (!recentlyClosedWritable) return;
    recentlyClosedWrites = recentlyClosedWrites.then(async () => {
      try {
        const key = await resolvePersistentWorkbookKey("write");
        if (key) await saveRecentlyClosed(settings, key.workbookId, items);
      } catch (error) {
        console.warn("[pi] Failed to persist recently closed tabs:", error);
      }
    });
  });
  try {
    const key = await resolvePersistentWorkbookKey("read");
    if (key) recentlyClosed.load(await loadRecentlyClosed(settings, key.workbookId));
    recentlyClosedWritable = true;
  } catch (error) {
    console.warn("[pi] Failed to read recently closed tabs:", error);
  }

  const getActiveRuntime = () => runtimeManager.getActiveRuntime();
  const getActiveAgent = () => getActiveRuntime()?.agent ?? null;
  const getActiveLockState = () => getActiveRuntime()?.lockState ?? "idle";

  listModelRefreshRuntimes = () => runtimeManager.listRuntimes().map((runtime) => ({
    runtimeId: runtime.runtimeId,
    model: runtime.agent.state.model,
    isBusy: isRuntimeBusy(runtime),
    applyModel: (model, thinkingLevel) => {
      runtime.agent.state.model = model;
      // A refreshed or swapped model may support fewer thinking levels than the current one.
      runtime.agent.state.thinkingLevel = clampThinkingLevel(model, thinkingLevel ?? runtime.agent.state.thinkingLevel);
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

  const stopBackupIdentityPolling = startBackupIdentityPolling({
    monitor: new BackupIdentityMonitor({
      rescopeBackups: (args) => workbookRecoveryLog.rescopeWorkbook(args),
    }),
  });

  window.addEventListener("beforeunload", () => {
    stopBackupIdentityPolling();
  }, { once: true });

  /**
   * Runs a user-started workbook operation (restore, full backup) under the
   * workbook lock, so it waits for Pi's current write instead of
   * interleaving with it. This is why backup commands are allowed while busy.
   */
  const runUserWorkbookOperation = async <T,>(toolName: string, operation: () => Promise<T>): Promise<T> => {
    const activeRuntime = getActiveRuntime();
    const sessionId = activeRuntime?.persistence.getSessionId() ?? crypto.randomUUID();
    const workbookId = await resolveWorkbookId();

    const { result } = await workbookCoordinator.runWrite(
      {
        workbookId: workbookId ?? "workbook:unknown",
        sessionId,
        opId: crypto.randomUUID(),
        toolName,
      },
      operation,
    );
    return result;
  };

  /** Restores a backup and tells every chat tab about it ("restored from a backup", #792). */
  const restoreNotingTabs = async (snapshotId: string) => {
    const window: PiWriteWindow = {
      toolCallId: `restore:${snapshotId}`,
      startedAt: Date.now(),
      settledAt: Number.POSITIVE_INFINITY,
    };
    changeTracker.notePiWrite(window, { kind: "restore" });
    try {
      const result = await workbookRecoveryLog.restore(snapshotId);
      window.address = result.address;
      return result;
    } finally {
      window.settledAt = Date.now();
    }
  };

  const formatBackupAge = (at: number): string => formatRelativeDate(new Date(at).toISOString());

  /** Every restore the user starts (/revert, the Backups page) runs here. */
  const restoreAfterConfirmation = async (preview: WorkbookRecoveryRestorePreview): Promise<void> => {
    const edits = preview.laterEdits;
    if (edits && !await requestConfirmationDialog({
      title: t("init.confirm.restoreOverEditsTitle"),
      message: t("init.confirm.restoreOverEditsMessage", {
        age: formatBackupAge(preview.at),
        address: preview.address,
        count: edits.count,
        sample: edits.sample.map((cell) => `${cell.address} = ${cell.content}`).join("; ")
          + (edits.count > edits.sample.length ? "; …" : ""),
      }),
      confirmLabel: t("init.confirm.restoreAnyway"),
      cancelLabel: t("init.confirm.keepEdits"),
      confirmButtonTone: "danger",
      restoreFocusOnClose: true,
    })) return;

    const restored = await runUserWorkbookOperation(
      "workbook_history", () => restoreNotingTabs(preview.id),
    );
    const address = restored.address;
    if (restored.inverseSnapshotError) {
      showToast(t("init.revertedWithoutRollback", { address }), { duration: 8000, variant: "error" });
      return;
    }
    showToast(t("init.revertedAddress", { address, age: formatBackupAge(preview.at) }));
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
    const tabs = runtimeManager.snapshotTabs();
    const tabNumbers: Record<string, number> = {};
    for (const tab of tabs) {
      tabNumbers[tab.sessionId] = tab.tabNumber;
    }

    return {
      sessionIds: tabs.map((tab) => tab.sessionId),
      activeSessionId: tabs.find((tab) => tab.isActive)?.sessionId ?? null,
      tabNumbers,
      nextTabNumber: runtimeManager.getNextTabNumber(),
    };
  };

  const tabLayoutPersistence = createTabLayoutPersistence({
    resolveWorkbookId,
    saveLayout: async (_workbookId, layout) => {
      const key = await resolvePersistentWorkbookKey("write");
      if (key) await saveWorkbookTabLayout(settings, key.workbookId, layout);
    },
    warn: (message, error) => {
      console.warn(message, error);
      showToast(t("sessions.toast.layoutSaveFailed", { error: getErrorMessage(error) }));
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

  bindRuntimeSidebar({
    runtimeManager,
    sidebar,
    renderError: (message) => renderErrorBanner(errorRoot, message),
  });

  let lastReconciledModelSignature = "";
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

    // A tab that registers, or restores a saved model, after the provider list
    // was published missed that pass's reconcile (P-66, P-67): check it now.
    const modelSignature = runtimeManager.listRuntimes()
      .map((runtime) => `${runtime.runtimeId}:${runtime.agent.state.model.provider}/${runtime.agent.state.model.id}`)
      .join("|");
    if (modelSignature !== lastReconciledModelSignature) {
      lastReconciledModelSignature = modelSignature;
      modelRefreshOwner.reconcileRuntimes();
    }
  });

  const refreshCapabilitiesForAllRuntimes = (): Promise<void> => {
    return runtimeManager.refreshCapabilities();
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
    getActiveAgent,
    refreshRuntimeTools: refreshCapabilitiesForAllRuntimes,
    refreshRuntimeModels: () => modelRefreshOwner.refresh(true),
    reservedToolNames,
    afterInjectAgentContext: async () => {
      const activeRuntime = getActiveRuntime();
      if (!activeRuntime) {
        return;
      }

      sidebar.syncFromAgent();
      sidebar.requestUpdate();
      document.dispatchEvent(new CustomEvent("pi:status-update"));
      await activeRuntime.persistence.saveSession({ force: true });
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

  // Convention overrides live in the system prompt; rebuild it when they change.
  document.addEventListener("pi:conventions-updated", () => {
    void refreshCapabilitiesForAllRuntimes();
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

  // Other taskpanes (browser tabs, Excel windows) share settings storage but
  // not DOM events. Tell them to rebuild their prompt when guidance changes here.
  if (typeof BroadcastChannel !== "undefined") {
    const settingsChannel = new BroadcastChannel("pi-for-excel:settings-changed");
    for (const eventName of [
      "pi:rules-updated",
      "pi:conventions-updated",
      PI_SKILLS_CHANGED_EVENT,
      PI_INTEGRATIONS_CHANGED_EVENT,
    ]) {
      document.addEventListener(eventName, () => {
        settingsChannel.postMessage(eventName);
      });
    }
    settingsChannel.addEventListener("message", () => {
      void refreshWorkbookState();
    });
  }

  document.addEventListener(PI_EXECUTION_MODE_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_BACKUP_SAVE_FAILED_EVENT, () => {
    showToast(t("init.backupSaveFailed"), { duration: 8000, variant: "error" });
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

  /**
   * Every approval goes through one queue (`requestApproval`), so parallel
   * calls and other tabs wait their turn instead of declining each other.
   */
  const requestRuntimeToolApproval = (args: {
    title: string;
    message: string;
    confirmLabel?: string;
    signal?: AbortSignal;
    source?: string;
    onStop?: () => void;
  }): Promise<boolean> => {
    return requestApproval({
      title: args.title,
      message: normalizeApprovalMessage(args.title, args.message),
      confirmLabel: args.confirmLabel ?? t("init.confirm.allowOnce"),
      ...(args.signal ? { signal: args.signal } : {}),
      ...(args.source ? { source: args.source } : {}),
      ...(args.onStop ? { onStop: args.onStop } : {}),
    });
  };

  async function buildSessionRuntime(
    optsForRuntime: CreateRuntimeOptions,
  ): Promise<SessionRuntime> {
    const runtimeId = crypto.randomUUID();
    let runtimeSessionId: string = optsForRuntime.sessionId ?? crypto.randomUUID();
    const runtimeSkillReadCache = createSkillReadCache();

    let runtimeAgent: Agent | null = null;

    /** Names this tab in its approvals when another tab is on screen. */
    /** "Stop Pi" in an approval dialog: the same as the tab's Stop button. */
    const stopRuntimeTurn = (): void => {
      const runtime = runtimeManager.getRuntime(runtimeId);
      if (runtime) stopTab(runtime);
    };

    const approvalSource = (): { source?: string } => {
      if (getActiveRuntime()?.runtimeId === runtimeId) return {};
      const runtime = runtimeManager.getRuntime(runtimeId);
      if (!runtime) return {};
      return { source: t("approval.fromTab", { title: resolveRuntimeTabTitle(runtimeId, runtime) }) };
    };

    const buildRuntimeCapabilities = async (sessionId: string): Promise<{
      tools: ReturnType<typeof withWorkbookCoordinator>;
      systemPrompt: string;
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
        getSessionId: () => runtimeAgent?.sessionId ?? runtimeSessionId,
        skillReadCache: runtimeSkillReadCache,
      }).filter(isRuntimeAgentTool);

      // Snapshot extension tools + revision together so schema-stable handler
      // reloads are compared against the same point-in-time payload.
      const extensionTools = extensionManager.getRegisteredTools();
      const extensionToolRevision = extensionManager.getToolRevision();

      const runtimeTools = normalizeRuntimeTools([
        ...coreTools,
        ...createToolsForIntegrations(activeIntegrationIds),
        ...extensionTools,
      ]);

      const requiredConnectionIds = collectRequiredConnectionIds(runtimeTools);

      const coordinatedTools = withWorkbookCoordinator(
        runtimeTools,
        workbookCoordinator,
        {
          getWorkbookId: resolveWorkbookId,
          getSessionId: () => runtimeAgent?.sessionId ?? runtimeSessionId,
        },
        {
          onWriteStarted: (window) => changeTracker.notePiWrite(window, { kind: "runtime", runtimeId }),
          onWriteCommitted: (event) => {
            if (event.impact !== "structure") return;
            invalidateBlueprint(event.workbookId);
            autoContext.noteOwnStructureChange();
          },
        },
        {
          getExecutionMode: () => executionModeController.readMode(),
          requestToolApproval: (request) => {
            const deletesBackups = request.reason === "always";
            const details = humanizeToolInput(request.toolName, request.params);
            const overwrites = request.overwrites
              ? t("approval.overwrites", {
                count: request.overwrites.count,
                sample: request.overwrites.sample
                  .map((cell) => `${cell.address} = ${cell.content.length > 40 ? `${cell.content.slice(0, 40)}…` : cell.content}`)
                  .join("; ") + (request.overwrites.count > request.overwrites.sample.length ? "; …" : ""),
              })
              : null;
            return requestApproval({
              title: deletesBackups ? t("approval.title.deleteBackups") : t("approval.title.workbookChange"),
              headline: describePendingToolCall(request.toolName, request.params) ?? request.toolName,
              ...(details ? { details } : {}),
              note: deletesBackups ? t("approval.note.deleteBackups") : t("approval.note.confirmMode"),
              confirmLabel: deletesBackups ? t("approval.allowDelete") : t("init.confirm.allowOnce"),
              confirmTone: deletesBackups ? "danger" : "primary",
              ...(request.signal ? { signal: request.signal } : {}),
              ...(overwrites ? { overwrites } : {}),
              onStop: stopRuntimeTurn,
              ...approvalSource(),
            });
          },
        },
      );

      // Gates wrap the coordinated tools, so their approvals (direct JS,
      // Python bridge) are asked before the call waits for the workbook lock.
      const gatedTools = await applyExperimentalToolGates(coordinatedTools, {
        getExecutionMode: () => executionModeController.readMode(),
        requestOfficeJsExecuteApproval: (request) => {
          const apiName = request.apiName ?? "Office.js";
          return requestRuntimeToolApproval({
            title: apiName === "WPS JSAPI"
              ? t("init.confirm.wpsJsTitle")
              : t("init.confirm.officeJsTitle"),
            message: buildOfficeJsExecuteApprovalMessage(request),
            confirmLabel: t("init.confirm.allowOnce"),
            ...(request.signal ? { signal: request.signal } : {}),
            onStop: stopRuntimeTurn,
            ...approvalSource(),
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
            ...(request.signal ? { signal: request.signal } : {}),
            onStop: stopRuntimeTurn,
            ...approvalSource(),
          });
        },
      });


      const preflightTools = withConnectionPreflight(gatedTools, {
        connectionManager,
      });

      const tools = applyToolOutputTruncation(preflightTools, {
        // Scale caps with the active model's context window (#566).
        limits: () => effectiveToolOutputLimits(runtimeAgent?.state.model.contextWindow),
        saveTruncatedOutput: saveTruncatedToolOutputToWorkspace,
      });

      const systemPrompt = await buildRuntimeSystemPrompt({
        workbookId,
        activeIntegrationIds,
        requiredConnectionIds,
      });

      return {
        tools,
        systemPrompt,
        extensionToolRevision,
      };
    };

    const initialModel = getActiveRuntime()?.agent.state.model
      ?? modelRefreshOwner.snapshot().defaultModel;
    const autoContext = createAutoContext({ changeTracker, getAgent: () => runtimeAgent, runtimeId });
    const initialCapabilities = await buildRuntimeCapabilities(runtimeSessionId);

    const agent = new Agent({
      initialState: {
        model: initialModel,
        thinkingLevel: initialModel.reasoning ? "high" : "off",
        messages: [{
          role: "system",
          content: "",
          sections: { [RUNTIME_SYSTEM_PROMPT_SECTION]: initialCapabilities.systemPrompt },
          toolsAdded: initialCapabilities.tools.map(toToolDeclaration),
          timestamp: 0,
        }],
        tools: initialCapabilities.tools,
      },
      convertToLlm: createConvertToLlm({
        getContextWindow: () => runtimeAgent?.state.model.contextWindow,
      }),
      prepareRequest: autoContext.prepareRequest,
      transformContext: autoContext.transformContext,
      streamFn,
    });

    runtimeAgent = agent;
    const uninstallImageInputNormalization = installImageInputNormalization(agent, {
      autoResizeImages,
    });
    let currentRuntimeToolsFingerprint = createRuntimeToolFingerprint(initialCapabilities.tools);
    let currentExtensionToolRevision = initialCapabilities.extensionToolRevision;

    const refreshRuntimeCapabilities = async () => {
      const nextSessionId = runtimeAgent?.sessionId ?? runtimeSessionId;
      runtimeSessionId = nextSessionId;

      const next = await buildRuntimeCapabilities(nextSessionId);
      const nextToolsFingerprint = createRuntimeToolFingerprint(next.tools);
      const nextExtensionToolRevision = next.extensionToolRevision;

      if (
        shouldApplyRuntimeToolUpdate({
          previousFingerprint: currentRuntimeToolsFingerprint,
          nextFingerprint: nextToolsFingerprint,
          previousExtensionToolRevision: currentExtensionToolRevision,
          nextExtensionToolRevision,
        })
      ) {
        agent.state.tools = next.tools;
        currentRuntimeToolsFingerprint = nextToolsFingerprint;
        currentExtensionToolRevision = nextExtensionToolRevision;
      }

      const currentRuntimeSystemPrompt = getCurrentSystemMessage(agent.state.messages)
        ?.sections?.[RUNTIME_SYSTEM_PROMPT_SECTION];
      if (next.systemPrompt !== currentRuntimeSystemPrompt) {
        agent.state.messages = [
          ...agent.state.messages,
          {
            role: "system",
            content: "",
            sections: { [RUNTIME_SYSTEM_PROMPT_SECTION]: next.systemPrompt },
            timestamp: Date.now(),
          },
        ];
      }
    };

    const tab = createTabState();
    // Set when the user dismisses the connection dialog during this turn.
    let missingCredentialsFor: string | null = null;
    // Set when this turn's login could not be refreshed and its token has expired.
    let expiredLoginFor: string | null = null;

    // API key resolution
    agent.getApiKey = async (provider: string) => {
      // OAuth access tokens expire during long sessions; refresh before the request, not at reload.
      if (await oauthRefresher.ensureFresh(provider) === "expired") {
        // Sending the stale token would only surface the provider's 401, and
        // returning no key lets pi-ai fall back to that same stored token.
        // Throwing fails the turn; the banner tells the user to sign in again.
        expiredLoginFor = provider;
        throw new Error(`The ${provider} login has expired and could not be refreshed.`);
      }
      const resolvedAuth = await modelRuntime.models.getAuth(provider);
      if (resolvedAuth) return resolvedAuth.auth.apiKey;

      // Extension provider credentials are configured through their host-owned
      // connection cards, never through the generic API-key dialog.
      if (modelRuntime.isExtensionProvider(provider)) {
        tab.setError(t("runtime.error.noApiKey", { provider }));
        return undefined;
      }

      const success = await promptForProviderConnection(provider);
      await modelRefreshOwner.refreshConfiguredProviders(false);
      if (success) {
        tab.setError(null);
        return (await getAppStorage().providerKeys.get(provider)) ?? undefined;
      }

      missingCredentialsFor = provider;
      return undefined;
    };

    const queueDisplay = createQueueDisplay({ agent, held: tab });
    // Assigned below; compaction saves the session as soon as it rewrites history.
    let savePersistence: () => Promise<void> = () => Promise.resolve();
    const actionQueue = createActionQueue({
      agent,
      tab,
      queueDisplay,
      autoCompactEnabled,
      runCompact: async (args, signal) => {
        const before = agent.state.messages;
        await runCompactCommand(agent, args, signal);
        if (agent.state.messages !== before) {
          // Redraw the context meter for the smaller history and save it now,
          // so a reload right after compaction does not bring the old one back.
          // Forced: compaction can archive every user message into the
          // summary, which a "has the chat started" save gate would skip.
          document.dispatchEvent(new CustomEvent("pi:status-update"));
          await savePersistence();
        }
      },
      onTurnSettled: ({ overflowAfterRecovery }) => {
        const credentials = missingCredentialsFor;
        missingCredentialsFor = null;
        const expiredLogin = expiredLoginFor;
        expiredLoginFor = null;
        void getConfiguredProxyUrl().then((proxyUrl) => {
          tab.setError(describeTurnFailure({
            messages: agent.state.messages,
            isContextOverflow: findTrailingContextOverflowError(agent.state) !== null,
            overflowAfterRecovery,
            missingCredentialsFor: credentials,
            expiredLoginFor: expiredLogin,
            online: navigator.onLine,
            proxyUrl: proxyUrl ?? null,
            defaultProxyUrl: DEFAULT_PROXY_URL,
          }));
        });
      },
      onLeftUnsent: (count) => {
        if (runtimeManager.getActiveRuntime()?.runtimeId !== runtimeId) return;
        showToast(t("turn.toast.leftUnsent", { count, plural: count === 1 ? "" : "s" }));
      },
    });

    const persistence = await setupSessionPersistence({
      agent,
      sessions,
      settings,
      resolveModel: (persisted) => modelRefreshOwner.resolveModel(persisted),
      initialSessionId: runtimeSessionId,
      autoRestoreLatest: optsForRuntime.autoRestoreLatest,
      onSaveError: (message) => {
        showToast(t("sessions.toast.saveFailed", { error: message }));
      },
      onSavedAsCopy: () => {
        showToast(t("init.sessionSavedAsCopy"), 8000);
      },
    });
    savePersistence = async () => { await persistence.saveSession({ force: true }); };

    runtimeSessionId = persistence.getSessionId();
    await refreshRuntimeCapabilities();

    let observedSessionId = runtimeSessionId;
    const unsubscribeSessionCapabilitySync = persistence.subscribe(() => {
      const nextSessionId = persistence.getSessionId();
      if (nextSessionId === observedSessionId) return;

      runtimeSkillReadCache.clearSession(observedSessionId);
      observedSessionId = nextSessionId;
      runtimeSessionId = nextSessionId;
      void refreshRuntimeCapabilities();
    });

    // The tab's draft is kept with its session, so a reload brings it back.
    const unbindDraft = draftStore.bind(tab, persistence);

    const unsubscribeErrorTracking = agent.subscribe((ev) => {
      // A new message clears the tab's previous error.
      if (ev.type === "message_start" && ev.message.role === "user") {
        tab.setError(null);
      }

      if (ev.type !== "agent_end") return;

      // Provider refreshes intentionally skip working runtimes. Re-run the
      // reconciliation after the action queue unwinds so an extension that
      // unloads mid-turn cannot leave its now-missing provider selected.
      window.setTimeout(() => modelRefreshOwner.reconcileRuntimes(), 0);
    });

    return {
      runtimeId,
      agent,
      actionQueue,
      queueDisplay,
      tab,
      persistence,
      lockState: "idle",
      refreshCapabilities: refreshRuntimeCapabilities,
      dispose: () => {
        runtimeSkillReadCache.clearAll();
        unsubscribeSessionCapabilitySync();
        unsubscribeErrorTracking();
        unbindDraft();
        queueDisplay.detach();
        actionQueue.shutdown();
        uninstallImageInputNormalization();
        agent.abort();
        persistence.dispose();
      },
    };
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
    const key = await resolvePersistentWorkbookKey("read");
    if (!key) {
      tabLayoutReadSucceeded = true;
      return null;
    }
    const workbookId = key.workbookId;
    let savedLayout: WorkbookTabLayout | null;
    try {
      savedLayout = await loadWorkbookTabLayout(settings, workbookId);
      tabLayoutReadSucceeded = true;
    } catch (error) {
      console.warn("[pi] Failed to read persisted tab layout:", error);
      return null;
    }
    if (!savedLayout) return null;

    // Tabs keep their "Chat N" numbers across reloads, and closed tabs' numbers stay retired.
    runtimeManager.seedNextTabNumber(savedLayout.nextTabNumber ?? 1);
    const runtimesBySessionId = new Map<string, SessionRuntime>();
    let firstRuntime: SessionRuntime | null = null;

    for (const sessionId of savedLayout.sessionIds) {
      // Keeping the saved id lets a blank tab keep its identity (and its
      // unsent draft) across reloads, and the saved number keeps its "Chat N" label.
      const tabNumber = savedLayout.tabNumbers?.[sessionId];
      const runtime = await runtimeManager.createRuntime({
        activate: false,
        autoRestoreLatest: false,
        sessionId,
        ...(tabNumber !== undefined ? { tabNumber } : {}),
      });

      const sessionData = await sessions.loadSession(sessionId);
      if (sessionData) {
        await runtime.persistence.applyLoadedSession(sessionData);
      } else {
        // Keep blank tabs durable across reloads.
        await runtime.persistence.saveSession({ force: true });
      }

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
    runtime.queueDisplay.clear();
    runtime.queueDisplay.setActionQueue([]);
    sidebar.syncFromAgent();
    sidebar.requestUpdate();
    document.dispatchEvent(new CustomEvent("pi:model-changed"));
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  };

  const getRuntimeTabTitle = (runtimeId: string): string => {
    return runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.title ?? t("session.untitled");
  };

  /**
   * Opens a saved session in a tab. A session is open in at most one tab, so
   * a session that is already open switches to its tab. Shows the refusal
   * toast itself and returns null when nothing was opened.
   */
  const openSavedSession = async (
    sessionData: SessionData,
    target: OpenSessionTarget,
  ): Promise<{ outcome: "opened" | "replaced" | "focused"; title: string } | null> => {
    const result = await runtimeManager.openSession(sessionData, target);
    if (result.outcome === "missing") {
      showToast(t("init.noActiveSession"));
      return null;
    }
    if (result.outcome === "busy") {
      showToast(t("init.currentTabBusy"));
      return null;
    }

    // An open session is not "recently closed" any more.
    recentlyClosed.removeBySessionId(result.runtime.persistence.getSessionId());
    if (result.outcome !== "focused") {
      syncRuntimeAfterSessionLoad(result.runtime);
    }
    return { outcome: result.outcome, title: getRuntimeTabTitle(result.runtime.runtimeId) };
  };

  const openResumePicker = async (defaultTarget: ResumeDialogTarget = "new_tab"): Promise<void> => {
    await showResumeDialog({
      defaultTarget,
      onOpenSession: async (sessionData: SessionData, resumeTarget: ResumeDialogTarget) => {
        if (resumeTarget === "new_tab") {
          return openSavedSession(sessionData, { kind: "new_tab" });
        }

        const activeRuntime = getActiveRuntime();
        if (!activeRuntime) {
          showToast(t("init.noActiveSession"));
          return null;
        }
        return openSavedSession(sessionData, { kind: "replace", runtimeId: activeRuntime.runtimeId });
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

  /** The undo toast shown for the most recent close, so reopening can retire it. */
  let closeUndoToast: { itemId: string; toast: ActionToastHandle } | null = null;

  const reopenRecentlyClosedItem = async (item: RecentlyClosedItem): Promise<ReopenRecentlyClosedResult> => {
    if (closeUndoToast?.itemId === item.id) {
      closeUndoToast.toast.dismiss();
      closeUndoToast = null;
    }

    try {
      const sessionData = await sessions.loadSession(item.sessionId);
      if (!sessionData) {
        showToast(t("init.couldNotReopenSession"));
        return "missing";
      }

      const opened = await openSavedSession(sessionData, {
        kind: "new_tab",
        ...(item.tabNumber !== undefined ? { tabNumber: item.tabNumber } : {}),
      });
      if (!opened) return "failed";

      showToast(opened.outcome === "focused"
        ? t("init.alreadyOpen", { title: opened.title })
        : t("init.reopened", { title: opened.title }));
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
    // "Latest" is decided under the lock: while Pi is busy, the write in
    // flight adds its backup first, and /revert undoes that one.
    const preview = await runUserWorkbookOperation("workbook_history", async () => {
      const latest = (await workbookRecoveryLog.listForCurrentWorkbook(1))[0];
      return latest ? workbookRecoveryLog.previewRestore(latest.id) : null;
    });
    if (!preview) {
      showToast(t("init.noBackupsYet"));
      return;
    }
    await restoreAfterConfirmation(preview);
  };

  const createManualFullBackup = async (): Promise<{
    id: string;
    createdAt: number;
    sizeBytes: number;
  }> => {
    const backup = await manualFullBackupStore.createAndDownload({
      runCapture: (capture) => runUserWorkbookOperation("manual_backup", capture),
    });
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

  const clearManualFullBackups = async (confirm: (count: number) => Promise<boolean>): Promise<number> => {
    return manualFullBackupStore.clearForCurrentWorkbook(confirm);
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

    if (runtime.agent.state.isStreaming || runtime.actionQueue.isBusy()) {
      const proceed = await requestConfirmationDialog({
        title: t("init.confirm.stopAndCloseTitle"),
        message: t("init.confirm.stopAndCloseMessage"),
        confirmLabel: t("init.confirm.stopAndClose"),
        cancelLabel: t("init.confirm.keepOpen"),
        confirmButtonTone: "danger",
      });
      if (!proceed) return false;

      // The tab kept running while the dialog was open; check again what closing would break.
      const current = runtimeManager.getRuntime(runtimeId);
      if (current !== runtime) return false;
      if (runtimeManager.listRuntimes().length <= 1) {
        showToast(t("init.cantCloseLastTab"));
        return false;
      }
      runtime.actionQueue.stop();
    }

    // A write that already started (it may have taken the lock while the
    // dialog was open) cannot be cancelled inside Excel. Keep the tab until
    // its run has settled, so the change lands with its backup and is saved
    // in the session instead of happening after the tab is gone.
    // Only a write makes the user wait; an aborted model request settles at once.
    if (runtime.lockState !== "idle") {
      showToast(t("init.waitingForChangeBeforeClose"));
    }
    await runtime.agent.waitForIdle();

    await runtime.persistence.saveSession({ force: true });

    const closingTab = runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId);
    const closedItem: RecentlyClosedItem = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2, 10)}`,
      sessionId: runtime.persistence.getSessionId(),
      title: closingTab?.title ?? formatSessionTitle(runtime.persistence.getSessionTitle()),
      closedAt: new Date().toISOString(),
      workbookId: await resolveWorkbookId(),
      ...(closingTab ? { tabNumber: closingTab.tabNumber } : {}),
    };

    runtimeManager.closeRuntime(runtimeId);
    recentlyClosed.push(closedItem);

    const showUndoToast = optsForClose?.showUndoToast !== false;
    if (showUndoToast) {
      const toast = showActionToast({
        message: t("init.closedTab", { title: closedItem.title }),
        actionLabel: t("init.undo"),
        duration: 9000,
        onAction: () => {
          closeUndoToast = null;
          void reopenRecentlyClosedById(closedItem.id);
        },
      });
      closeUndoToast = { itemId: closedItem.id, toast };
    }

    return true;
  };

  const renameRuntimeTab = async (runtimeId: string): Promise<void> => {
    const runtime = runtimeManager.getRuntime(runtimeId);
    if (!runtime) {
      showToast(t("init.sessionNotFound"));
      return;
    }

    if (isRuntimeBusy(runtime)) {
      showToast(t("init.waitBeforeRenaming"));
      return;
    }

    const currentTitle = runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.title
      ?? formatSessionTitle(runtime.persistence.getSessionTitle());
    const defaultTitle = runtime.persistence.hasExplicitTitle()
      ? runtime.persistence.getSessionTitle().trim()
      : currentTitle;

    // Office on the web replaces window.prompt with a function that throws, so ask in the taskpane.
    const nextTitleRaw = await requestTextInputDialog({
      title: t("init.renameTab.title"),
      initialValue: defaultTitle,
      confirmLabel: t("init.renameTab.confirm"),
    });
    if (nextTitleRaw === null) {
      return;
    }
    // The tab may have started a turn while the dialog was open.
    if (runtimeManager.getRuntime(runtimeId) !== runtime) return;

    const nextTitle = nextTitleRaw.trim();
    await runtime.persistence.renameSession(nextTitle);
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
      ?? formatSessionTitle(runtime.persistence.getSessionTitle());
  };

  type RuntimeModel = Agent["state"]["model"];

  const cloneRuntimeToNewTab = async (args: {
    sourceRuntime: SessionRuntime;
    targetModel: RuntimeModel;
    targetTitle: string;
  }): Promise<SessionRuntime> => {
    const clonedRuntime = await runtimeManager.createRuntime({
      activate: true,
      autoRestoreLatest: false,
    });

    clonedRuntime.agent.state.messages = structuredClone(args.sourceRuntime.agent.state.messages);
    clonedRuntime.agent.state.model = args.targetModel;
    clonedRuntime.agent.state.thinkingLevel = args.sourceRuntime.agent.state.thinkingLevel;

    await clonedRuntime.persistence.renameSession(args.targetTitle);
    clonedRuntime.queueDisplay.clear();
    clonedRuntime.queueDisplay.setActionQueue([]);
    sidebar.syncFromAgent();
    sidebar.requestUpdate();
    document.dispatchEvent(new CustomEvent("pi:model-changed"));
    document.dispatchEvent(new CustomEvent("pi:status-update"));
    await clonedRuntime.persistence.saveSession({ force: true });

    return clonedRuntime;
  };

  const duplicateRuntimeTab = async (runtimeId: string): Promise<void> => {
    const sourceRuntime = runtimeManager.getRuntime(runtimeId);
    if (!sourceRuntime) {
      showToast(t("init.sessionNotFound"));
      return;
    }

    if (isRuntimeBusy(sourceRuntime)) {
      showToast(t("init.waitBeforeDuplicating"));
      return;
    }

    const sourceTitle = resolveRuntimeTabTitle(runtimeId, sourceRuntime);
    const duplicateTitle = `${sourceTitle} copy`;

    await cloneRuntimeToNewTab({
      sourceRuntime,
      targetModel: sourceRuntime.agent.state.model,
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

    // Recompute every tab on this workbook: one write finishing or queuing
    // changes what the others are waiting for.
    const snapshot = workbookCoordinator.getSnapshot(event.context.workbookId);
    const finishedOpId = event.type === "completed" || event.type === "failed" ? event.context.opId : undefined;
    for (const runtime of runtimeManager.listRuntimes()) {
      runtimeManager.setRuntimeLockState(
        runtime.runtimeId,
        getSessionLockState(snapshot, runtime.persistence.getSessionId(), finishedOpId),
      );
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
      getActiveSessionId: () => getActiveRuntime()?.persistence.getSessionId() ?? null,
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
        // Every backup of this workbook (at most "Keep at most"), so counts
        // such as "Clear all" match what is stored.
        const checkpoints = await workbookRecoveryLog.listForCurrentWorkbook(MAX_RECOVERY_ENTRIES);
        return checkpoints.map((checkpoint) => toRecoveryCheckpointSummary(checkpoint));
      },
      onRestore: async (snapshotId: string) => {
        await restoreAfterConfirmation(await workbookRecoveryLog.previewRestore(snapshotId));
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
        await workbookRecoveryLog.applyRetentionLimit();
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
      sidebar.syncFromAgent();
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

    void (async () => {
      try {
        await modelRefreshOwner.refreshConfiguredProviders(false);
      } catch (error) {
        console.warn("[auth] Failed to refresh providers before opening model selector:", error);
      }

      closeStatusPopover();

      // Read the model after the refresh, which can swap a tab whose provider went away.
      const currentModel = runtimeManager.getRuntime(targetRuntimeId)?.agent.state.model
        ?? activeRuntime.agent.state.model;
      openModelSelectorDialog({
        models: modelRuntime.models,
        currentModel,
        // Same rule from every entry point (/model, status bar): a busy tab keeps
        // its model, so the list says so up front instead of refusing a pick.
        getDisabledReason: () => {
          const runtime = runtimeManager.getRuntime(targetRuntimeId);
          return runtime && isRuntimeBusy(runtime) ? t("modelSelector.busy") : undefined;
        },
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
    getActiveAgent,
    openModelSelector,
    renameActiveSession: async (title: string) => {
      const activeRuntime = getActiveRuntime();
      if (!activeRuntime) {
        showToast(t("init.noActiveSession"));
        return;
      }

      await activeRuntime.persistence.renameSession(title);
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
    clearManualFullBackups: async (confirm) => {
      return clearManualFullBackups(confirm);
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

    const busy = isRuntimeBusy(activeRuntime);
    const result = executeSlashCommand({
      name,
      args,
      busy,
      enqueueCommand: (commandName: string, commandArgs: string) => {
        activeRuntime.actionQueue.enqueueCommand(commandName, commandArgs);
      },
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

  sidebar.onSend = (text, { followUp }) => {
    const activeRuntime = getActiveRuntime();
    const accepted = submitComposerText({ text, followUp, target: activeRuntime });
    if (accepted && activeRuntime) activeRuntime.tab.setError(null);
    return accepted;
  };

  /**
   * Stop is one behavior, from the Stop button or Escape: the run, compaction
   * or retry wait is cancelled; steers, follow-ups and queued prompts stay
   * as separate "Not sent" rows; queued commands are cancelled.
   */
  const stopActiveTab = (): void => {
    const activeRuntime = getActiveRuntime();
    if (activeRuntime) stopTab(activeRuntime);
  };

  /** A tab's Stop: also used by "Stop Pi" in that tab's approval dialogs (#794). */
  function stopTab(runtime: SessionRuntime): void {
    const { texts, cancelledCommands } = runtime.actionQueue.stop();
    const count = texts.length;
    const plural = count === 1 ? "" : "s";
    const commands = cancelledCommands.join(", ");
    if (count > 0 && commands) {
      showToast(t("turn.toast.stopKeptCancelled", { count, plural, commands }));
    } else if (count > 0) {
      showToast(t("turn.toast.stopKept", { count, plural }));
    } else if (commands) {
      showToast(t("turn.toast.stopCancelled", { commands }));
    }
  }

  sidebar.onAbort = stopActiveTab;

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
      .then((result) => {
        const count = result.imported;
        if (count <= 0) {
          showToast(t("init.noFilesImported"));
          return;
        }

        const importedLabel = `${count} file${count === 1 ? "" : "s"}`;
        showToast(t("init.importedIntoFiles", { label: importedLabel }) + describeImportRenames(result));
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
    // Without an identity in a real workbook there is no "latest session" of this workbook to restore.
    const autoRestoreLatest = usesSharedUnidentifiedState || (await resolveWorkbookId()) !== null;
    await runtimeManager.createRuntime({ activate: true, autoRestoreLatest });
  }
  for (const tab of runtimeManager.snapshotTabs()) {
    recentlyClosed.removeBySessionId(tab.sessionId);
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

  // ── Keyboard shortcuts ──
  installKeyboardShortcuts({
    getActiveAgent,
    isActiveTabBusy: () => {
      const runtime = getActiveRuntime();
      return runtime !== null && (runtime.agent.state.isStreaming || runtime.actionQueue.isBusy());
    },
    sidebar,
    stopActiveTab,
    editLastQueued: () => {
      const edited = getActiveRuntime()?.actionQueue.editLast() ?? false;
      if (edited) focusChatInputSoon();
      else showToast(t("keyboard-shortcuts.toast.noQueuedMessages"));
      return edited;
    },
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
    // Ctrl/Cmd+Z reopens only the tab whose Undo toast is on screen. Recently
    // closed now survives reloads, so a bare Ctrl+Z could otherwise bring back a
    // tab closed days ago when the user meant to undo in Excel (CL-02).
    // /reopen and the resume picker still reach older tabs.
    canUndoCloseTab: () => closeUndoToast?.toast.isShowing() === true,
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
    getActiveAgent,
    getLockState: getActiveLockState,
    getExecutionMode,
  });

  // ── Proxy status polling ──
  startProxyPolling(settings);

  // ── Wire command menu to textarea ──
  const wireTextarea = () => {
    const composer = sidebar.getInput();
    if (composer) {
      wireCommandMenu(composer);
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

    const activeAgent = getActiveAgent();
    if (!activeAgent) {
      showToast(t("init.noActiveSession"));
      return;
    }

    const description = trigger.getAttribute("data-tooltip") ?? "Choose how deeply Pi reasons before responding.";

    toggleThinkingPopover({
      anchor: trigger,
      description,
      levels: getThinkingLevels(activeAgent),
      activeLevel: activeAgent.state.thinkingLevel,
      onSelectLevel: (level) => {
        if (activeAgent.state.thinkingLevel === level) return;
        activeAgent.state.thinkingLevel = level;
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
