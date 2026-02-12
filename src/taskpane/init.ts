/**
 * Taskpane initialization.
 *
 * Contains the bulk of app wiring (storage, auth restore, runtime manager,
 * sidebar mount, slash commands, keyboard shortcuts).
 */

import { html, render } from "lit";
import { Agent, type AgentTool } from "@mariozechner/pi-agent-core";
import { ApiKeyPromptDialog } from "@mariozechner/pi-web-ui/dist/dialogs/ApiKeyPromptDialog.js";
import { ModelSelector } from "@mariozechner/pi-web-ui/dist/dialogs/ModelSelector.js";
import { ApiKeysTab, ProxyTab, SettingsDialog } from "@mariozechner/pi-web-ui/dist/dialogs/SettingsDialog.js";
import { getAppStorage } from "@mariozechner/pi-web-ui/dist/storage/app-storage.js";
import type { SessionData } from "@mariozechner/pi-web-ui/dist/storage/types.js";

import { createOfficeStreamFn } from "../auth/stream-proxy.js";
import { isLoopbackProxyUrl } from "../auth/proxy-validation.js";
import { restoreCredentials } from "../auth/restore.js";
import { invalidateBlueprint } from "../context/blueprint.js";
import { ChangeTracker } from "../context/change-tracker.js";
import {
  PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT,
  PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT,
} from "../experiments/events.js";
import { isExperimentalFeatureEnabled, setExperimentalFeatureEnabled } from "../experiments/flags.js";
import { convertToLlm } from "../messages/convert-to-llm.js";
import { getFilesWorkspace } from "../files/workspace.js";
import { createAllTools } from "../tools/index.js";
import { applyExperimentalToolGates } from "../tools/experimental-tool-gates.js";
import { withWorkbookCoordinator } from "../tools/with-workbook-coordinator.js";
import { registerBuiltins } from "../commands/builtins.js";
import { showExtensionsDialog } from "../commands/builtins/extensions-overlay.js";
import { showSkillsDialog } from "../commands/builtins/skills-overlay.js";
import { ExtensionRuntimeManager } from "../extensions/runtime-manager.js";
import type { ResumeDialogTarget } from "../commands/builtins/resume-target.js";
import {
  showInstructionsDialog,
  showRecoveryDialog,
  showResumeDialog,
  showShortcutsDialog,
  type RecoveryCheckpointSummary,
} from "../commands/builtins/overlays.js";
import { wireCommandMenu } from "../commands/command-menu.js";
import { commandRegistry } from "../commands/types.js";
import {
  getUserInstructions,
  getWorkbookInstructions,
  hasAnyInstructions,
} from "../instructions/store.js";
import { getResolvedConventions } from "../conventions/store.js";
import {
  buildSkillPromptEntries,
  createToolsForSkills,
  getSkillToolNames,
  SKILL_IDS,
} from "../skills/catalog.js";
import { PI_SKILLS_CHANGED_EVENT } from "../skills/events.js";
import { getExternalToolsEnabled, resolveConfiguredSkillIds } from "../skills/store.js";
import { buildSystemPrompt } from "../prompt/system-prompt.js";
import { initAppStorage } from "../storage/init-app-storage.js";
import { renderError } from "../ui/loading.js";
import { showFilesWorkspaceDialog } from "../ui/files-dialog.js";
import { showActionToast, showToast } from "../ui/toast.js";
import { PiSidebar } from "../ui/pi-sidebar.js";
import { setActiveProviders } from "../compat/model-selector-patch.js";
import { createWorkbookCoordinator } from "../workbook/coordinator.js";
import { formatWorkbookLabel, getWorkbookContext } from "../workbook/context.js";
import {
  PI_WORKBOOK_SNAPSHOT_CREATED_EVENT,
  type WorkbookSnapshotCreatedDetail,
} from "../workbook/recovery-events.js";
import { getWorkbookRecoveryLog, type WorkbookRecoverySnapshot } from "../workbook/recovery-log.js";

import { createContextInjector } from "./context-injection.js";
import { pickDefaultModel } from "./default-model.js";
import { getThinkingLevels, installKeyboardShortcuts } from "./keyboard-shortcuts.js";
import { createQueueDisplay } from "./queue-display.js";
import { createActionQueue } from "./action-queue.js";
import { RecentlyClosedStack, type RecentlyClosedItem } from "./recently-closed.js";
import { setupSessionPersistence } from "./sessions.js";
import {
  loadWorkbookTabLayout,
  saveWorkbookTabLayout,
  type WorkbookTabLayout,
} from "./tab-layout.js";
import { injectStatusBar } from "./status-bar.js";
import {
  closeStatusPopover,
  toggleContextPopover,
  toggleThinkingPopover,
} from "./status-popovers.js";
import { showWelcomeLogin } from "./welcome-login.js";
import {
  SessionRuntimeManager,
  type RuntimeTabSnapshot,
  type SessionRuntime,
} from "./session-runtime-manager.js";
import { isRecord } from "../utils/type-guards.js";

const BUSY_ALLOWED_COMMANDS = new Set(["compact", "new", "instructions", "resume", "history", "reopen", "extensions"]);

function showErrorBanner(errorRoot: HTMLElement, message: string): void {
  render(renderError(message), errorRoot);
}

function clearErrorBanner(errorRoot: HTMLElement): void {
  render(html``, errorRoot);
}

function isRuntimeAgentTool(value: unknown): value is AgentTool {
  if (!isRecord(value)) return false;

  return typeof value.name === "string"
    && typeof value.label === "string"
    && typeof value.description === "string"
    && "parameters" in value
    && typeof value.execute === "function";
}

function isLikelyCorsErrorMessage(msg: string): boolean {
  const m = msg.toLowerCase();

  // Browser/network errors
  if (m.includes("failed to fetch")) return true;
  if (m.includes("load failed")) return true; // WebKit/Safari
  if (m.includes("networkerror")) return true;

  // Explicit CORS wording
  if (m.includes("cors") || m.includes("cross-origin")) return true;

  // Anthropic sometimes returns a JSON 401 with a CORS-specific message when direct browser access is disabled.
  if (m.includes("cors requests are not allowed")) return true;

  return false;
}

function getActiveLockNotice(tabs: RuntimeTabSnapshot[]): string | null {
  const activeTab = tabs.find((tab) => tab.isActive);
  if (!activeTab) return null;

  if (activeTab.lockState === "waiting_for_lock") {
    return "Waiting for workbook lock…";
  }

  if (activeTab.lockState === "holding_lock") {
    return "Applying workbook changes…";
  }

  return null;
}

function parseWorkbookSnapshotCreatedDetail(value: unknown): WorkbookSnapshotCreatedDetail | null {
  if (!isRecord(value)) return null;

  const snapshotId = typeof value.snapshotId === "string" ? value.snapshotId : null;
  const rawToolName = typeof value.toolName === "string" ? value.toolName : null;
  const address = typeof value.address === "string" ? value.address : null;
  const changedCount = typeof value.changedCount === "number" ? value.changedCount : null;

  const toolName =
    rawToolName === "write_cells" ||
    rawToolName === "fill_formula" ||
    rawToolName === "python_transform_range" ||
    rawToolName === "restore_snapshot"
      ? rawToolName
      : null;

  if (!snapshotId || !toolName || !address || changedCount === null) {
    return null;
  }

  return {
    snapshotId,
    toolName,
    address,
    changedCount,
  };
}

export async function initTaskpane(opts: {
  appEl: HTMLElement;
  errorRoot: HTMLElement;
}): Promise<void> {
  const { appEl, errorRoot } = opts;

  const changeTracker = new ChangeTracker();

  // 1. Storage
  const { providerKeys, sessions, settings } = initAppStorage();

  // 1b. Auto-compaction (Pi defaults to enabled)
  let autoCompactEnabled = true;
  try {
    autoCompactEnabled = (await settings.get<boolean>("compaction.enabled")) ?? true;
  } catch {
    autoCompactEnabled = true;
  }

  // 1c. Security warning: remote proxies can see your prompts + credentials.
  try {
    const proxyEnabled = await settings.get<boolean>("proxy.enabled");
    const proxyUrl = await settings.get<string>("proxy.url");
    if (
      proxyEnabled === true &&
      typeof proxyUrl === "string" &&
      proxyUrl.trim().length > 0 &&
      !isLoopbackProxyUrl(proxyUrl)
    ) {
      showToast("Security warning: proxy URL is not localhost — it can see your tokens and prompts.");
    }
  } catch {
    // ignore
  }

  // 2. Restore auth
  await restoreCredentials(providerKeys, settings);

  // 2b. Welcome/login if no providers
  const configuredProviders = await providerKeys.list();
  if (configuredProviders.length === 0) {
    await showWelcomeLogin(providerKeys);
  }

  // 3. Change tracker
  changeTracker.start().catch(() => {});

  // 4. Shared runtime dependencies
  // Workbook structure context is injected separately by transformContext.
  const availableProviders = await providerKeys.list();
  setActiveProviders(new Set(availableProviders));
  const defaultModel = pickDefaultModel(availableProviders);

  const streamFn = createOfficeStreamFn(async () => {
    // In dev mode, Vite's reverse proxy handles CORS — don't double-proxy.
    if (import.meta.env.DEV) return undefined;

    try {
      const storage = getAppStorage();
      const enabled = await storage.settings.get("proxy.enabled");
      if (!enabled) return undefined;
      const url = await storage.settings.get("proxy.url");
      return typeof url === "string" && url.trim().length > 0
        ? url.trim().replace(/\/+$/, "")
        : undefined;
    } catch {
      return undefined;
    }
  });

  const workbookCoordinator = createWorkbookCoordinator();

  // 5. Create and mount PiSidebar
  const sidebar = new PiSidebar();
  sidebar.emptyHints = [
    {
      label: "Explain this workbook",
      prompt: "Read through the entire workbook — every sheet, its structure, formulas, and named ranges. Then write a clear overview and user manual for this workbook.\nCover: what the workbook does, how it's organized, the logic flow between sheets, where inputs live and where outputs are derived.\nIf it's a model: explain the key assumptions (and where to change them), the calculation logic, and how outputs depend on inputs.\nIf it's data: explain what the data represents, the key fields, any derived columns, and notable patterns or gaps.\nStructure your explanation like documentation — start with a summary, then walk through each sheet's role.",
    },
    {
      label: "Quality check this workbook",
      prompt: "Review this workbook for errors and issues across logic, assumptions, and formatting:\n- Logic: broken or circular references, hardcoded numbers inside formulas, inconsistent formula patterns across rows/columns, missing links between sheets, #REF or #VALUE errors.\n- Assumptions: flag any key assumptions (e.g. growth rates, discount rates, margins) — are they reasonable? Are they clearly labelled and easy to find, or buried in formulas?\n- Formatting: inconsistent number formats within columns, missing or misaligned headers, unlabelled input cells, inconsistent decimal places or currency symbols, rows/columns that break the visual pattern.\nSummarize your findings as a prioritized list of recommendations, grouped by severity.",
    },
    {
      label: "Build my financial model",
      prompt: "First, read through the entire workbook — every sheet, its structure, formulas, named ranges, and any existing data. Form a clear picture of what's already here and how it's organized.\nIf the workbook is blank or mostly empty: ask me what kind of financial model I need — for example a DCF, LBO, three-statement model, budget, forecast, or comparison — then build it step by step, starting with the assumptions.\nIf there's a partially complete model: explain the current structure, the logic flow, key assumptions, and what's missing or incomplete. Offer to extend or finish it.\nIf there's data but no model: explain what the data represents, suggest what could be modelled from it, and offer to build it.",
    },
    {
      label: "Format this sheet",
      prompt: "Review this worksheet and infer the correct format for each cell from context, then apply formatting including:\n- Number formats (currency, percentages, dates, integers vs. decimals)\n- Font colour coding (e.g. blue for inputs, black for formulas)\n- Cell styles for inputs, outputs, and headers\n- Consistent headers and section labels\nEnsure formats are consistent: for example, if all other cells in a column use one decimal place, apply the same. If a row is bold or italicised, extend that to any unformatted cells in the row.\nAfter formatting, read back the sheet and verify your changes look correct.",
    },
  ];

  appEl.innerHTML = "";
  appEl.appendChild(sidebar);

  let instructionsActive = false;

  const setInstructionsActive = (next: boolean) => {
    if (instructionsActive === next) return;
    instructionsActive = next;
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  };

  const resolveWorkbookContext = async (): Promise<Awaited<ReturnType<typeof getWorkbookContext>>> => {
    try {
      return await getWorkbookContext();
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

  const resolveRuntimeSkillIds = async (args: {
    sessionId: string;
    workbookId: string | null;
  }): Promise<string[]> => {
    const configuredSkillIds = await resolveConfiguredSkillIds({
      settings,
      sessionId: args.sessionId,
      workbookId: args.workbookId,
      knownSkillIds: SKILL_IDS,
    });

    const externalToolsEnabled = await getExternalToolsEnabled(settings);
    return externalToolsEnabled ? configuredSkillIds : [];
  };

  const buildRuntimeSystemPrompt = async (args: {
    workbookId: string | null;
    activeSkillIds: readonly string[];
  }): Promise<string> => {
    try {
      const userInstructions = await getUserInstructions(settings);
      const workbookInstructions = await getWorkbookInstructions(settings, args.workbookId);
      setInstructionsActive(hasAnyInstructions({ userInstructions, workbookInstructions }));
      const conventions = await getResolvedConventions(settings);
      const activeSkills = buildSkillPromptEntries(args.activeSkillIds);
      return buildSystemPrompt({ userInstructions, workbookInstructions, activeSkills, conventions });
    } catch {
      setInstructionsActive(false);
      return buildSystemPrompt();
    }
  };

  const runtimeManager = new SessionRuntimeManager(sidebar);
  const abortedAgents = new WeakSet<Agent>();
  const runtimeCapabilityRefreshers = new Map<string, () => Promise<void>>();
  const runtimeActiveSkillIds = new Map<string, string[]>();
  const recentlyClosed = new RecentlyClosedStack(10);

  const getActiveRuntime = () => runtimeManager.getActiveRuntime();
  const getActiveAgent = () => getActiveRuntime()?.agent ?? null;
  const getActiveQueueDisplay = () => getActiveRuntime()?.queueDisplay ?? null;
  const getActiveActionQueue = () => getActiveRuntime()?.actionQueue ?? null;
  const getActiveLockState = () => getActiveRuntime()?.lockState ?? "idle";

  const workbookRecoveryLog = getWorkbookRecoveryLog();

  const restoreCheckpointById = async (snapshotId: string): Promise<void> => {
    const activeRuntime = getActiveRuntime();
    const sessionId = activeRuntime?.persistence.getSessionId() ?? crypto.randomUUID();

    const workbookId = await resolveWorkbookId();
    const coordinatorWorkbookId = workbookId ?? "workbook:unknown";

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
    showToast(`Reverted ${address}`);
  };

  const toRecoveryCheckpointSummary = (
    snapshot: WorkbookRecoverySnapshot,
  ): RecoveryCheckpointSummary => ({
    id: snapshot.id,
    at: snapshot.at,
    toolName: snapshot.toolName,
    address: snapshot.address,
    changedCount: snapshot.changedCount,
    restoredFromSnapshotId: snapshot.restoredFromSnapshotId,
  });

  const getActiveSkillTitles = (): string[] => {
    const runtime = getActiveRuntime();
    if (!runtime) return [];

    const skillIds = runtimeActiveSkillIds.get(runtime.runtimeId) ?? [];
    return buildSkillPromptEntries(skillIds).map((entry) => entry.title);
  };

  const formatSessionTitle = (title: string): string => {
    const trimmed = title.trim();
    return trimmed.length > 0 ? trimmed : "Untitled";
  };

  const snapshotRuntimeTabLayout = (): WorkbookTabLayout => {
    const runtimes = runtimeManager.listRuntimes();
    const activeRuntime = runtimeManager.getActiveRuntime();

    return {
      sessionIds: runtimes.map((runtime) => runtime.persistence.getSessionId()),
      activeSessionId: activeRuntime?.persistence.getSessionId() ?? null,
    };
  };

  const tabLayoutSignature = (layout: WorkbookTabLayout): string => JSON.stringify(layout);

  let previousActiveRuntimeId: string | null = null;
  let tabLayoutPersistenceEnabled = false;
  let lastPersistedTabLayoutSignature: string | null = null;
  let tabLayoutPersistChain: Promise<void> = Promise.resolve();

  const maybePersistTabLayout = (): void => {
    if (!tabLayoutPersistenceEnabled) return;

    const layout = snapshotRuntimeTabLayout();
    const layoutSignature = tabLayoutSignature(layout);

    tabLayoutPersistChain = tabLayoutPersistChain
      .then(
        async () => {
          const workbookId = await resolveWorkbookId();
          const persistSignature = `${workbookId ?? "__global__"}|${layoutSignature}`;
          if (persistSignature === lastPersistedTabLayoutSignature) return;

          await saveWorkbookTabLayout(settings, workbookId, layout);
          lastPersistedTabLayoutSignature = persistSignature;
        },
        () => undefined,
      )
      .catch((error: unknown) => {
        console.warn("[pi] Failed to persist tab layout:", error);
      });
  };

  runtimeManager.subscribe((tabs) => {
    sidebar.sessionTabs = tabs;
    sidebar.lockNotice = getActiveLockNotice(tabs);
    sidebar.requestUpdate();

    const activeRuntimeId = tabs.find((tab) => tab.isActive)?.runtimeId ?? null;
    if (activeRuntimeId !== previousActiveRuntimeId) {
      previousActiveRuntimeId = activeRuntimeId;
      document.dispatchEvent(new CustomEvent("pi:active-runtime-changed"));
    }

    maybePersistTabLayout();
    document.dispatchEvent(new CustomEvent("pi:status-update"));
  });

  const refreshCapabilitiesForAllRuntimes = async () => {
    const runtimes = runtimeManager.listRuntimes();

    for (const runtime of runtimes) {
      const refresh = runtimeCapabilityRefreshers.get(runtime.runtimeId);
      if (!refresh) continue;

      try {
        await refresh();
      } catch (error: unknown) {
        console.warn("[pi] Failed to refresh runtime capabilities:", error);
      }
    }

    document.dispatchEvent(new CustomEvent("pi:status-update"));
  };

  const reservedToolNames = new Set([
    ...createAllTools().map((tool) => tool.name),
    ...getSkillToolNames(),
  ]);
  const extensionManager = new ExtensionRuntimeManager({
    settings,
    getActiveAgent,
    refreshRuntimeTools: refreshCapabilitiesForAllRuntimes,
    reservedToolNames,
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

  document.addEventListener("pi:instructions-updated", () => {
    void refreshWorkbookState();
  });

  document.addEventListener(PI_EXPERIMENTAL_FEATURE_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_EXPERIMENTAL_TOOL_CONFIG_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_SKILLS_CHANGED_EVENT, () => {
    void refreshCapabilitiesForAllRuntimes();
  });

  document.addEventListener(PI_WORKBOOK_SNAPSHOT_CREATED_EVENT, (event) => {
    if (!(event instanceof CustomEvent)) return;

    const detail = parseWorkbookSnapshotCreatedDetail(event.detail);
    if (!detail) return;
    if (detail.toolName === "restore_snapshot") return;

    const changedLabel = detail.changedCount > 0
      ? `${detail.changedCount.toLocaleString()} cell${detail.changedCount === 1 ? "" : "s"}`
      : "range";

    showActionToast({
      message: `Checkpoint saved for ${detail.address} (${changedLabel})`,
      actionLabel: "Revert",
      duration: 10_000,
      onAction: () => {
        void restoreCheckpointById(detail.snapshotId)
          .catch((error: unknown) => {
            const message = error instanceof Error ? error.message : "Unknown error";
            showToast(`Revert failed: ${message}`);
          });
      },
    });
  });

  const createRuntime = async (optsForRuntime: {
    activate: boolean;
    autoRestoreLatest: boolean;
  }) => {
    const runtimeId = crypto.randomUUID();
    let runtimeSessionId: string = crypto.randomUUID();

    let runtimeAgent: Agent | null = null;

    const buildRuntimeCapabilities = async (sessionId: string): Promise<{
      tools: ReturnType<typeof withWorkbookCoordinator>;
      systemPrompt: string;
    }> => {
      const workbookId = await resolveWorkbookId();
      const activeSkillIds = await resolveRuntimeSkillIds({
        sessionId,
        workbookId,
      });

      runtimeActiveSkillIds.set(runtimeId, activeSkillIds);

      const coreTools = createAllTools().filter(isRuntimeAgentTool);
      const gatedCoreTools = await applyExperimentalToolGates(coreTools);
      const allTools = [
        ...gatedCoreTools,
        ...createToolsForSkills(activeSkillIds),
        ...extensionManager.getRegisteredTools(),
      ];

      const tools = withWorkbookCoordinator(
        allTools,
        workbookCoordinator,
        {
          getWorkbookId: resolveWorkbookId,
          getSessionId: () => runtimeAgent?.sessionId ?? runtimeSessionId,
        },
        {
          onWriteCommitted: (event) => {
            if (event.impact !== "structure") return;
            invalidateBlueprint(event.workbookId);
          },
        },
      );

      const systemPrompt = await buildRuntimeSystemPrompt({
        workbookId,
        activeSkillIds,
      });

      return {
        tools,
        systemPrompt,
      };
    };

    const initialCapabilities = await buildRuntimeCapabilities(runtimeSessionId);

    const agent = new Agent({
      initialState: {
        systemPrompt: initialCapabilities.systemPrompt,
        model: defaultModel,
        thinkingLevel: defaultModel.reasoning ? "high" : "off",
        messages: [],
        tools: initialCapabilities.tools,
      },
      convertToLlm,
      transformContext: createContextInjector(changeTracker),
      streamFn,
    });

    runtimeAgent = agent;

    const refreshRuntimeCapabilities = async () => {
      const nextSessionId = runtimeAgent?.sessionId ?? runtimeSessionId;
      runtimeSessionId = nextSessionId;

      const next = await buildRuntimeCapabilities(nextSessionId);
      agent.setTools(next.tools);
      agent.setSystemPrompt(next.systemPrompt);
    };

    runtimeCapabilityRefreshers.set(runtimeId, refreshRuntimeCapabilities);

    // API key resolution
    agent.getApiKey = async (provider: string) => {
      const key = await getAppStorage().providerKeys.get(provider);
      if (key) return key;

      const success = await ApiKeyPromptDialog.prompt(provider);
      const updated = await providerKeys.list();
      setActiveProviders(new Set(updated));
      if (success) {
        clearErrorBanner(errorRoot);
        return (await getAppStorage().providerKeys.get(provider)) ?? undefined;
      }

      showErrorBanner(errorRoot, `API key required for ${provider}.`);
      return undefined;
    };

    const queueDisplay = createQueueDisplay({ agent });
    const actionQueue = createActionQueue({
      agent,
      sidebar,
      queueDisplay,
      autoCompactEnabled,
    });

    const persistence = await setupSessionPersistence({
      agent,
      sessions,
      settings,
      initialSessionId: runtimeSessionId,
      autoRestoreLatest: optsForRuntime.autoRestoreLatest,
    });

    runtimeSessionId = persistence.getSessionId();
    await refreshRuntimeCapabilities();

    let observedSessionId = runtimeSessionId;
    const unsubscribeSessionCapabilitySync = persistence.subscribe(() => {
      const nextSessionId = persistence.getSessionId();
      if (nextSessionId === observedSessionId) return;

      observedSessionId = nextSessionId;
      runtimeSessionId = nextSessionId;
      void refreshRuntimeCapabilities();
    });

    const unsubscribeErrorTracking = agent.subscribe((ev) => {
      const isActiveRuntime = runtimeManager.getActiveRuntime()?.runtimeId === runtimeId;

      if (ev.type === "message_start" && ev.message.role === "user" && isActiveRuntime) {
        clearErrorBanner(errorRoot);
      }

      if (ev.type !== "agent_end") return;

      const wasUserAbort = abortedAgents.has(agent);
      abortedAgents.delete(agent);

      if (!isActiveRuntime) return;

      if (agent.state.error) {
        const isAbort = wasUserAbort || /abort/i.test(agent.state.error) || /cancel/i.test(agent.state.error);
        if (!isAbort) {
          const err = agent.state.error;
          if (isLikelyCorsErrorMessage(err)) {
            showErrorBanner(
              errorRoot,
              "Network error (likely CORS). Start the local HTTPS proxy (npm run proxy:https) and enable it in /settings → Proxy.",
            );
          } else {
            showErrorBanner(errorRoot, `LLM error: ${err}`);
          }
        }
      } else {
        clearErrorBanner(errorRoot);
      }
    });

    const runtime = runtimeManager.createRuntime(
      {
        runtimeId,
        agent,
        actionQueue,
        queueDisplay,
        persistence,
        lockState: "idle",
        dispose: () => {
          runtimeCapabilityRefreshers.delete(runtimeId);
          runtimeActiveSkillIds.delete(runtimeId);
          unsubscribeSessionCapabilitySync();
          unsubscribeErrorTracking();
          actionQueue.shutdown();
          agent.abort();
          persistence.dispose();
        },
      },
      { activate: optsForRuntime.activate },
    );

    return runtime;
  };

  const restorePersistedTabLayout = async (): Promise<SessionRuntime | null> => {
    const workbookId = await resolveWorkbookId();
    const savedLayout = await loadWorkbookTabLayout(settings, workbookId);
    if (!savedLayout) return null;

    const runtimesBySessionId = new Map<string, SessionRuntime>();
    let firstRuntime: SessionRuntime | null = null;

    for (const sessionId of savedLayout.sessionIds) {
      const runtime = await createRuntime({
        activate: false,
        autoRestoreLatest: false,
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

  const replaceActiveRuntimeSession = async (sessionData: SessionData): Promise<void> => {
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast("No active session");
      return;
    }

    const busy = activeRuntime.agent.state.isStreaming || activeRuntime.actionQueue.isBusy();
    if (busy) {
      showToast("Current tab is busy — use open in new tab or wait for it to finish");
      return;
    }

    await activeRuntime.persistence.applyLoadedSession(sessionData);
    syncRuntimeAfterSessionLoad(activeRuntime);
  };

  const openSessionInNewTab = async (sessionData: SessionData): Promise<SessionRuntime> => {
    const runtime = await createRuntime({
      activate: true,
      autoRestoreLatest: false,
    });

    await runtime.persistence.applyLoadedSession(sessionData);
    syncRuntimeAfterSessionLoad(runtime);
    return runtime;
  };

  const reopenRecentlyClosedItem = async (item: RecentlyClosedItem): Promise<boolean> => {
    try {
      const sessionData = await sessions.loadSession(item.sessionId);
      if (!sessionData) {
        showToast("Couldn't reopen session");
        return false;
      }

      await openSessionInNewTab(sessionData);
      showToast(`Reopened: ${formatSessionTitle(item.title)}`);
      return true;
    } catch {
      showToast("Couldn't reopen session");
      return false;
    }
  };

  const reopenLastClosed = async (): Promise<void> => {
    const item = recentlyClosed.popMostRecent();
    if (!item) {
      showToast("No recently closed tab");
      return;
    }

    await reopenRecentlyClosedItem(item);
  };

  const revertLatestCheckpoint = async (): Promise<void> => {
    const latest = await workbookRecoveryLog.listForCurrentWorkbook(1);
    const checkpoint = latest[0];

    if (!checkpoint) {
      showToast("No recovery checkpoints for this workbook yet");
      return;
    }

    await restoreCheckpointById(checkpoint.id);
  };

  const closeRuntimeWithRecovery = async (runtimeId: string): Promise<void> => {
    if (runtimeManager.listRuntimes().length <= 1) {
      showToast("Can't close the last tab");
      return;
    }

    const runtime = runtimeManager.getRuntime(runtimeId);
    if (!runtime) return;

    if (runtime.lockState === "holding_lock") {
      showToast("Wait for workbook changes to finish before closing this tab");
      return;
    }

    if (runtime.agent.state.isStreaming) {
      const proceed = window.confirm("Pi is still responding in this tab. Stop and close it?");
      if (!proceed) return;

      abortedAgents.add(runtime.agent);
      runtime.agent.abort();
    }

    await runtime.persistence.saveSession({ force: true });

    const closeTitle = runtimeManager.snapshotTabs().find((tab) => tab.runtimeId === runtimeId)?.title
      ?? formatSessionTitle(runtime.persistence.getSessionTitle());

    const closedItem: RecentlyClosedItem = {
      sessionId: runtime.persistence.getSessionId(),
      title: closeTitle,
      closedAt: new Date().toISOString(),
      workbookId: await resolveWorkbookId(),
    };

    runtimeManager.closeRuntime(runtimeId);
    recentlyClosed.push(closedItem);

    showActionToast({
      message: `Closed ${closedItem.title}`,
      actionLabel: "Undo",
      duration: 9000,
      onAction: () => {
        const itemToRestore = recentlyClosed.removeBySessionId(closedItem.sessionId);
        if (!itemToRestore) return;
        void reopenRecentlyClosedItem(itemToRestore);
      },
    });
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

  const openSkillsManager = () => {
    showSkillsDialog({
      getActiveSessionId: () => getActiveRuntime()?.persistence.getSessionId() ?? null,
      resolveWorkbookContext: async () => {
        const workbookContext = await resolveWorkbookContext();
        return {
          workbookId: workbookContext.workbookId,
          workbookLabel: formatWorkbookLabel(workbookContext),
        };
      },
    });
  };

  const openRecoveryDialog = async (): Promise<void> => {
    const workbookContext = await resolveWorkbookContext();

    await showRecoveryDialog({
      workbookLabel: formatWorkbookLabel(workbookContext),
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
    });
  };

  registerBuiltins({
    getActiveAgent,
    renameActiveSession: async (title: string) => {
      const activeRuntime = getActiveRuntime();
      if (!activeRuntime) {
        showToast("No active session");
        return;
      }

      await activeRuntime.persistence.renameSession(title);
    },
    createRuntime: async () => {
      await createRuntime({ activate: true, autoRestoreLatest: false });
    },
    openResumeDialog: async (defaultTarget: ResumeDialogTarget = "new_tab") => {
      await showResumeDialog({
        defaultTarget,
        onOpenInNewTab: async (sessionData: SessionData) => {
          await openSessionInNewTab(sessionData);
        },
        onReplaceCurrent: async (sessionData: SessionData) => {
          await replaceActiveRuntimeSession(sessionData);
        },
      });
    },
    openRecoveryDialog,
    reopenLastClosed,
    revertLatestCheckpoint: async () => {
      try {
        await revertLatestCheckpoint();
      } catch (error: unknown) {
        const message = error instanceof Error ? error.message : "Unknown error";
        showToast(`Revert failed: ${message}`);
      }
    },
    openInstructionsEditor: async () => {
      await showInstructionsDialog({
        onSaved: async () => {
          await refreshWorkbookState();
        },
      });
    },
    openExtensionsManager: () => {
      showExtensionsDialog(extensionManager);
    },
    openSkillsManager,
  });

  // Slash commands chosen from the popup menu dispatch this event.
  const onCommandRun: EventListener = (event) => {
    if (!(event instanceof CustomEvent)) return;
    if (!isRecord(event.detail)) return;

    const name = typeof event.detail.name === "string" ? event.detail.name : "";
    const args = typeof event.detail.args === "string" ? event.detail.args : "";
    if (!name) return;

    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast("No active session");
      return;
    }

    const busy = activeRuntime.agent.state.isStreaming || activeRuntime.actionQueue.isBusy();
    if (busy && !BUSY_ALLOWED_COMMANDS.has(name)) {
      showToast(`Can't run /${name} while Pi is busy`);
      return;
    }

    if (name === "compact") {
      activeRuntime.actionQueue.enqueueCommand(name, args);
      return;
    }

    const cmd = commandRegistry.get(name);
    if (cmd) void cmd.execute(args);
  };
  document.addEventListener("pi:command-run", onCommandRun);

  sidebar.onSend = (text) => {
    clearErrorBanner(errorRoot);
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) {
      showToast("No active session");
      return;
    }
    activeRuntime.actionQueue.enqueuePrompt(text);
  };

  sidebar.onAbort = () => {
    const activeRuntime = getActiveRuntime();
    if (!activeRuntime) return;
    abortedAgents.add(activeRuntime.agent);
    activeRuntime.agent.abort();
  };

  sidebar.onCreateTab = () => {
    void createRuntime({ activate: true, autoRestoreLatest: false });
  };

  sidebar.onSelectTab = (runtimeId: string) => {
    runtimeManager.switchRuntime(runtimeId);
  };

  sidebar.onCloseTab = (runtimeId: string) => {
    void closeRuntimeWithRecovery(runtimeId);
  };
  sidebar.onOpenInstructions = () => {
    void showInstructionsDialog({
      onSaved: async () => {
        await refreshWorkbookState();
      },
    });
  };
  sidebar.onOpenSkills = () => {
    openSkillsManager();
  };
  sidebar.onOpenFiles = () => {
    void showFilesWorkspaceDialog();
  };
  sidebar.onOpenSettings = () => {
    void SettingsDialog.open([new ApiKeysTab(), new ProxyTab()]);
  };
  sidebar.onFilesDrop = (files: File[]) => {
    const workspace = getFilesWorkspace();

    void workspace.importFiles(files, {
      audit: { actor: "user", source: "input-drop" },
    })
      .then((count) => {
        if (count <= 0) {
          showToast("No files were imported.");
          return;
        }

        if (!isExperimentalFeatureEnabled("files_workspace")) {
          setExperimentalFeatureEnabled("files_workspace", true);
          showToast(`Imported ${count} file${count === 1 ? "" : "s"} and enabled files-workspace.`);
          return;
        }

        showToast(`Imported ${count} file${count === 1 ? "" : "s"}.`);
      })
      .catch((error: unknown) => {
        showToast(`Import failed: ${error instanceof Error ? error.message : "Unknown error"}`);
      });
  };
  sidebar.onOpenResumePicker = () => {
    void showResumeDialog({
      defaultTarget: "new_tab",
      onOpenInNewTab: async (sessionData: SessionData) => {
        await openSessionInNewTab(sessionData);
      },
      onReplaceCurrent: async (sessionData: SessionData) => {
        await replaceActiveRuntimeSession(sessionData);
      },
    });
  };
  sidebar.onOpenRecovery = () => {
    void openRecoveryDialog();
  };
  sidebar.onOpenShortcuts = () => {
    showShortcutsDialog();
  };


  // Bootstrap from persisted tab layout; fallback to legacy single-runtime restore.
  const restoredRuntime = await restorePersistedTabLayout();
  if (!restoredRuntime) {
    await createRuntime({ activate: true, autoRestoreLatest: true });
  }

  tabLayoutPersistenceEnabled = true;
  maybePersistTabLayout();

  // ── Register extensions ──
  await extensionManager.initialize();

  document.addEventListener("pi:providers-changed", () => {
    void (async () => {
      const updated = await providerKeys.list();
      setActiveProviders(new Set(updated));
    })();
  });

  // ── Keyboard shortcuts ──
  installKeyboardShortcuts({
    getActiveAgent,
    getActiveQueueDisplay,
    getActiveActionQueue,
    sidebar,
    markUserAborted: (agent: Agent) => {
      abortedAgents.add(agent);
    },
    onReopenLastClosed: () => {
      void reopenLastClosed();
    },
  });

  // ── Status bar ──
  injectStatusBar({
    getActiveAgent,
    getLockState: getActiveLockState,
    getInstructionsActive: () => instructionsActive,
    getActiveSkills: getActiveSkillTitles,
  });

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

  const openModelSelector = (): void => {
    const activeAgent = getActiveAgent();
    if (!activeAgent) {
      showToast("No active session");
      return;
    }

    closeStatusPopover();

    void ModelSelector.open(activeAgent.state.model, (model) => {
      activeAgent.setModel(model);
      document.dispatchEvent(new CustomEvent("pi:status-update"));
      requestAnimationFrame(() => sidebar.requestUpdate());
    });
  };

  const openThinkingPopoverFrom = (target: Element): void => {
    const trigger = target.closest(".pi-status-thinking");
    if (!trigger) return;

    const activeAgent = getActiveAgent();
    if (!activeAgent) {
      showToast("No active session");
      return;
    }

    const description = trigger.getAttribute("data-tooltip") ?? "Choose how long the model thinks before responding.";

    toggleThinkingPopover({
      anchor: trigger,
      description,
      levels: getThinkingLevels(activeAgent),
      activeLevel: activeAgent.state.thinkingLevel,
      onSelectLevel: (level) => {
        if (activeAgent.state.thinkingLevel === level) return;
        activeAgent.setThinkingLevel(level);
        document.dispatchEvent(new CustomEvent("pi:status-update"));
      },
    });
  };

  const openContextPopoverFrom = (target: Element): void => {
    const trigger = target.closest(".pi-status-ctx--trigger");
    if (!trigger) return;

    const description = trigger.getAttribute("data-status-popover")
      ?? trigger.querySelector(".pi-tooltip")?.textContent
      ?? "How much of the model's context window has been used.";

    toggleContextPopover({
      anchor: trigger,
      description,
      onRunCommand: (command) => {
        runSlashCommand(command);
      },
    });
  };

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

    // Instructions editor
    if (el.closest(".pi-status-instructions")) {
      closeStatusPopover();
      void showInstructionsDialog({
        onSaved: async () => {
          await refreshWorkbookState();
        },
      });
      return;
    }

    // Skills manager
    if (el.closest(".pi-status-skills")) {
      closeStatusPopover();
      openSkillsManager();
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
