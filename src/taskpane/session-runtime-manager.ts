/**
 * UI-independent lifecycle owner for taskpane session runtimes.
 */

import type { Agent } from "@earendil-works/pi-agent-core";
import { clampThinkingLevel } from "@earendil-works/pi-ai";

import {
  shouldForkModelSwitch,
  type ModelSwitchBehavior,
} from "../models/switch-behavior.js";
import { areRuntimeModelsEquivalent } from "../models/model-refresh-owner.js";
import type { ActionQueue } from "./action-queue.js";
import type { QueueDisplay } from "./queue-display.js";
import type { SessionData } from "../storage/local/types.js";
import { countVisibleSessionMessages, type SessionPersistenceController } from "./sessions.js";
import { resolveTabTitle } from "./session-title.js";
import type { TabState } from "./tab-state.js";

export type RuntimeLockState = "idle" | "waiting_for_lock" | "holding_lock";

/** How a run ended, as far as the user needs to know. */
export type RuntimeRunOutcome = "finished" | "failed";

/**
 * Marker for a background tab whose run ended while the user was elsewhere.
 * Cleared when the tab becomes active.
 */
export type RuntimeAttention = "none" | RuntimeRunOutcome;

export interface RuntimeRunSettledEvent {
  runtimeId: string;
  title: string;
  outcome: RuntimeRunOutcome;
  /** Whether the tab was the active one when its run ended. */
  wasActive: boolean;
}

export interface SessionRuntime {
  runtimeId: string;
  agent: Agent;
  actionQueue: ActionQueue;
  queueDisplay: QueueDisplay;
  /** Draft, error banner and activity line of this tab. */
  tab: TabState;
  persistence: SessionPersistenceController;
  lockState: RuntimeLockState;
  refreshCapabilities: () => Promise<void>;
  dispose: () => void;
}

export interface CreateRuntimeOptions {
  activate: boolean;
  autoRestoreLatest: boolean;
  /** Reuse this session id (a tab restored from the saved layout). */
  sessionId?: string;
  /**
   * Default "Chat N" number to give the tab, e.g. the number it had before a
   * reload or before it was closed. Ignored when another open tab has it.
   */
  tabNumber?: number;
}

export interface RuntimeTabSnapshot {
  runtimeId: string;
  sessionId: string;
  tabNumber: number;
  title: string;
  isActive: boolean;
  isStreaming: boolean;
  isBusy: boolean;
  lockState: RuntimeLockState;
  attention: RuntimeAttention;
}

export interface RuntimeLifecycleSnapshot {
  activeRuntimeId: string | null;
  tabs: RuntimeTabSnapshot[];
}

export type RuntimeSnapshotListener = (snapshot: RuntimeLifecycleSnapshot) => void;
export type SessionRuntimeFactory = (opts: CreateRuntimeOptions) => Promise<SessionRuntime>;
export type ModelSelectionResult =
  | { outcome: "missing" | "unchanged" | "busy" | "updated" }
  | { outcome: "forked"; runtime: SessionRuntime; title: string };

/**
 * Where a saved session should open. A session is open in at most one tab, so
 * either target switches to the tab that already shows it.
 */
export type OpenSessionTarget =
  | { kind: "new_tab"; tabNumber?: number }
  | { kind: "replace"; runtimeId: string };

export type OpenSessionResult =
  | { outcome: "opened" | "replaced" | "focused"; runtime: SessionRuntime }
  | { outcome: "busy" }
  | { outcome: "missing" };

export function isRuntimeBusy(runtime: Pick<SessionRuntime, "agent" | "actionQueue">): boolean {
  return runtime.agent.state.isStreaming || runtime.actionQueue.isBusy();
}

function isValidTabNumber(value: number | undefined): value is number {
  return value !== undefined && Number.isInteger(value) && value >= 1;
}

interface RuntimeListeners {
  unsubscribeAgent: () => void;
  unsubscribePersistence: () => void;
  unsubscribeTab: () => void;
}

export class SessionRuntimeManager {
  private readonly createSessionRuntime: SessionRuntimeFactory;
  private readonly warnCapabilityRefresh: (error: unknown) => void;
  private readonly runtimes = new Map<string, SessionRuntime>();
  private readonly runtimeOrder: string[] = [];
  private readonly runtimeDefaultTabNumbers = new Map<string, number>();
  private readonly runtimeListeners = new Map<string, RuntimeListeners>();
  private readonly listeners = new Set<RuntimeSnapshotListener>();
  private readonly getRunOutcome: (runtime: SessionRuntime) => RuntimeRunOutcome | null;
  private readonly onRunSettled: (event: RuntimeRunSettledEvent) => void;
  private readonly wasBusy = new Map<string, boolean>();
  private readonly attention = new Map<string, RuntimeRunOutcome>();

  private activeRuntimeId: string | null = null;
  private nextDefaultTabNumber = 1;
  private refreshPromise: Promise<void> | null = null;
  private refreshRequested = false;
  private refreshPassRunning = false;

  constructor(opts: {
    createRuntime: SessionRuntimeFactory;
    warnCapabilityRefresh?: (error: unknown) => void;
    /** Classify a run that just ended; `null` (e.g. the user stopped it) leaves no trace. */
    getRunOutcome?: (runtime: SessionRuntime) => RuntimeRunOutcome | null;
    /** Called once per run, when the tab goes from busy to idle. */
    onRunSettled?: (event: RuntimeRunSettledEvent) => void;
  }) {
    this.createSessionRuntime = opts.createRuntime;
    this.warnCapabilityRefresh = opts.warnCapabilityRefresh ?? (() => {});
    this.getRunOutcome = opts.getRunOutcome
      ?? ((runtime) => (runtime.agent.state.errorMessage ? "failed" : "finished"));
    this.onRunSettled = opts.onRunSettled ?? (() => {});
  }

  /** Re-read runtime state, e.g. when an action queue finishes outside an agent event. */
  notifyRuntimeChanged(): void {
    this.emit();
  }

  async createRuntime(opts: CreateRuntimeOptions): Promise<SessionRuntime> {
    const runtime = await this.createSessionRuntime(opts);
    return this.registerRuntime(runtime, opts);
  }

  registerRuntime(runtime: SessionRuntime, opts?: { activate?: boolean; tabNumber?: number }): SessionRuntime {
    this.runtimes.set(runtime.runtimeId, runtime);
    this.runtimeOrder.push(runtime.runtimeId);
    const requested = opts?.tabNumber;
    const inUse = new Set(this.runtimeDefaultTabNumbers.values());
    const tabNumber = isValidTabNumber(requested) && !inUse.has(requested)
      ? requested
      : this.nextDefaultTabNumber;
    this.nextDefaultTabNumber = Math.max(this.nextDefaultTabNumber, tabNumber + 1);
    this.runtimeDefaultTabNumbers.set(runtime.runtimeId, tabNumber);

    const unsubscribeAgent = runtime.agent.subscribe(() => {
      this.emit();
    });
    const unsubscribePersistence = runtime.persistence.subscribe(() => {
      this.emit();
    });
    // Compaction, retry waits and queued work change the busy dot without an
    // agent event; draft keystrokes do not need a new snapshot.
    const isTabBusy = () => runtime.tab.busy || runtime.tab.activity !== null;
    let wasBusy = isTabBusy();
    const unsubscribeTab = runtime.tab.subscribe(() => {
      const busy = isTabBusy();
      if (busy === wasBusy) return;
      wasBusy = busy;
      this.emit();
    });
    this.runtimeListeners.set(runtime.runtimeId, { unsubscribeAgent, unsubscribePersistence, unsubscribeTab });

    const shouldActivate = opts?.activate ?? this.activeRuntimeId === null;
    if (shouldActivate) {
      this.switchRuntime(runtime.runtimeId);
    } else {
      this.emit();
    }

    return runtime;
  }

  /**
   * The next number a new tab gets. Numbers are never reused while they are
   * remembered, so a reload can seed this from the saved layout.
   */
  getNextTabNumber(): number {
    return this.nextDefaultTabNumber;
  }

  seedNextTabNumber(next: number): void {
    if (isValidTabNumber(next) && next > this.nextDefaultTabNumber) {
      this.nextDefaultTabNumber = next;
    }
  }

  getTabNumber(runtimeId: string): number | null {
    return this.runtimeDefaultTabNumbers.get(runtimeId) ?? null;
  }

  /**
   * Opens a saved session. If any tab already shows it, that tab becomes
   * active instead, so one session never backs two tabs.
   */
  async openSession(sessionData: SessionData, target: OpenSessionTarget): Promise<OpenSessionResult> {
    const existing = this.findRuntimeBySessionId(sessionData.id);
    if (existing) {
      this.switchRuntime(existing.runtimeId);
      return { outcome: "focused", runtime: existing };
    }

    if (target.kind === "replace") {
      const runtime = this.runtimes.get(target.runtimeId);
      if (!runtime) return { outcome: "missing" };
      if (isRuntimeBusy(runtime)) return { outcome: "busy" };

      await runtime.persistence.applyLoadedSession(sessionData);
      this.switchRuntime(runtime.runtimeId);
      return { outcome: "replaced", runtime };
    }

    const runtime = await this.createRuntime({
      activate: true,
      autoRestoreLatest: false,
      ...(target.tabNumber !== undefined ? { tabNumber: target.tabNumber } : {}),
    });
    await runtime.persistence.applyLoadedSession(sessionData);
    this.emit();
    return { outcome: "opened", runtime };
  }

  switchRuntime(runtimeId: string): SessionRuntime | null {
    const next = this.runtimes.get(runtimeId);
    if (!next) return null;

    this.activeRuntimeId = runtimeId;
    this.attention.delete(runtimeId);
    this.emit();
    return next;
  }

  closeRuntime(runtimeId: string): SessionRuntime | null {
    if (!this.runtimes.has(runtimeId)) return null;
    if (this.runtimeOrder.length <= 1) return this.getActiveRuntime();

    const runtime = this.runtimes.get(runtimeId);
    if (!runtime) return null;

    const index = this.runtimeOrder.indexOf(runtimeId);
    if (index === -1) return null;

    const wasActive = this.activeRuntimeId === runtimeId;

    const listeners = this.runtimeListeners.get(runtimeId);
    listeners?.unsubscribeAgent();
    listeners?.unsubscribePersistence();
    listeners?.unsubscribeTab();
    this.runtimeListeners.delete(runtimeId);
    this.wasBusy.delete(runtimeId);
    this.attention.delete(runtimeId);

    runtime.dispose();

    this.runtimes.delete(runtimeId);
    this.runtimeOrder.splice(index, 1);
    this.runtimeDefaultTabNumbers.delete(runtimeId);

    if (wasActive) {
      const fallbackId = this.runtimeOrder[Math.max(0, index - 1)] ?? this.runtimeOrder[0] ?? null;
      if (!fallbackId) {
        this.activeRuntimeId = null;
        this.emit();
        return null;
      }
      return this.switchRuntime(fallbackId);
    }

    this.emit();
    return this.getActiveRuntime();
  }

  moveRuntime(runtimeId: string, direction: -1 | 1): boolean {
    const index = this.runtimeOrder.indexOf(runtimeId);
    if (index < 0) return false;

    const targetIndex = index + direction;
    if (targetIndex < 0 || targetIndex >= this.runtimeOrder.length) {
      return false;
    }

    this.runtimeOrder.splice(index, 1);
    this.runtimeOrder.splice(targetIndex, 0, runtimeId);
    this.emit();
    return true;
  }

  setRuntimeLockState(runtimeId: string, lockState: RuntimeLockState): void {
    const runtime = this.runtimes.get(runtimeId);
    if (!runtime || runtime.lockState === lockState) return;

    runtime.lockState = lockState;
    this.emit();
  }

  async selectModel(args: {
    runtimeId: string;
    nextModel: Agent["state"]["model"];
    behavior: ModelSwitchBehavior;
  }): Promise<ModelSelectionResult> {
    const runtime = this.runtimes.get(args.runtimeId);
    if (!runtime) return { outcome: "missing" };

    const currentModel = runtime.agent.state.model;
    const sameIdentity = currentModel.provider === args.nextModel.provider
      && currentModel.id === args.nextModel.id;
    if (sameIdentity && areRuntimeModelsEquivalent(currentModel, args.nextModel)) {
      return { outcome: "unchanged" };
    }

    if (isRuntimeBusy(runtime)) {
      return { outcome: "busy" };
    }

    // The thinking level carries over only as far as the new model supports it.
    const thinkingLevel = clampThinkingLevel(args.nextModel, runtime.agent.state.thinkingLevel);

    if (sameIdentity || !shouldForkModelSwitch({
      behavior: args.behavior,
      // Every tab starts with a hidden instruction message; only what the user sees counts.
      hasMessages: countVisibleSessionMessages(runtime.agent.state.messages) > 0,
    })) {
      runtime.agent.state.model = args.nextModel;
      runtime.agent.state.thinkingLevel = thinkingLevel;
      await runtime.persistence.saveSession({ force: true });
      return { outcome: "updated" };
    }

    const sourceTitle = this.snapshotTabs()
      .find((tab) => tab.runtimeId === args.runtimeId)?.title ?? "Untitled";
    const title = `${sourceTitle} (${args.nextModel.id})`;
    const forkedRuntime = await this.createRuntime({
      activate: true,
      autoRestoreLatest: false,
    });
    forkedRuntime.agent.state.messages = structuredClone(runtime.agent.state.messages);
    forkedRuntime.agent.state.model = args.nextModel;
    forkedRuntime.agent.state.thinkingLevel = thinkingLevel;
    await forkedRuntime.persistence.renameSession(title);
    forkedRuntime.queueDisplay.clear();
    forkedRuntime.queueDisplay.setActionQueue([]);
    await forkedRuntime.persistence.saveSession({ force: true });

    return { outcome: "forked", runtime: forkedRuntime, title };
  }

  refreshCapabilities(): Promise<void> {
    this.refreshRequested = true;
    if (this.refreshPromise) return this.refreshPromise;

    const pass = this.runCapabilityRefreshes();
    // A pass with no runtime to await has already finished and released the
    // slot; holding its settled promise would swallow every later request.
    if (this.refreshPassRunning) this.refreshPromise = pass;
    return pass;
  }

  findRuntimeBySessionId(sessionId: string): SessionRuntime | null {
    for (const runtimeId of this.runtimeOrder) {
      const runtime = this.runtimes.get(runtimeId);
      if (!runtime) continue;
      if (runtime.persistence.getSessionId() === sessionId) {
        return runtime;
      }
    }
    return null;
  }

  getRuntime(runtimeId: string): SessionRuntime | null {
    return this.runtimes.get(runtimeId) ?? null;
  }

  getActiveRuntime(): SessionRuntime | null {
    if (!this.activeRuntimeId) return null;
    return this.runtimes.get(this.activeRuntimeId) ?? null;
  }

  listRuntimes(): SessionRuntime[] {
    const out: SessionRuntime[] = [];
    for (const runtimeId of this.runtimeOrder) {
      const runtime = this.runtimes.get(runtimeId);
      if (runtime) out.push(runtime);
    }
    return out;
  }

  snapshot(): RuntimeLifecycleSnapshot {
    return {
      activeRuntimeId: this.activeRuntimeId,
      tabs: this.snapshotTabs(),
    };
  }

  snapshotTabs(): RuntimeTabSnapshot[] {
    return this.listRuntimes().map((runtime, index) => ({
      runtimeId: runtime.runtimeId,
      title: resolveTabTitle({
        hasExplicitTitle: runtime.persistence.hasExplicitTitle(),
        sessionTitle: runtime.persistence.getSessionTitle(),
        defaultTabNumber: this.runtimeDefaultTabNumbers.get(runtime.runtimeId) ?? (index + 1),
      }),
      isActive: runtime.runtimeId === this.activeRuntimeId,
      isStreaming: runtime.agent.state.isStreaming,
      isBusy: runtime.agent.state.isStreaming || runtime.actionQueue.isBusy(),
      lockState: runtime.lockState,
      attention: this.attention.get(runtime.runtimeId) ?? "none",
    }));
  }

  subscribe(listener: RuntimeSnapshotListener): () => void {
    this.listeners.add(listener);
    listener(this.snapshot());
    return () => {
      this.listeners.delete(listener);
    };
  }

  private async runCapabilityRefreshes(): Promise<void> {
    this.refreshPassRunning = true;
    try {
      while (this.refreshRequested) {
        this.refreshRequested = false;
        const runtimes = this.listRuntimes();

        for (const runtime of runtimes) {
          try {
            await runtime.refreshCapabilities();
          } catch (error) {
            this.warnCapabilityRefresh(error);
          }
        }
      }
    } finally {
      // Release the in-flight slot before notifying listeners so a refresh
      // requested from a snapshot listener starts a new pass instead of being
      // coalesced into this finished one.
      this.refreshPromise = null;
      this.refreshPassRunning = false;
    }

    this.emit();
  }

  /** Find runs that ended since the last emit and mark background tabs. */
  private collectSettledRuns(): Array<{ runtimeId: string; outcome: RuntimeRunOutcome; wasActive: boolean }> {
    const settled: Array<{ runtimeId: string; outcome: RuntimeRunOutcome; wasActive: boolean }> = [];
    for (const runtime of this.runtimes.values()) {
      const busy = runtime.agent.state.isStreaming || runtime.actionQueue.isBusy();
      const wasBusy = this.wasBusy.get(runtime.runtimeId) ?? false;
      this.wasBusy.set(runtime.runtimeId, busy);
      if (busy) {
        this.attention.delete(runtime.runtimeId);
        continue;
      }
      if (!wasBusy) continue;

      const outcome = this.getRunOutcome(runtime);
      if (!outcome) continue;
      const wasActive = runtime.runtimeId === this.activeRuntimeId;
      if (!wasActive) this.attention.set(runtime.runtimeId, outcome);
      settled.push({ runtimeId: runtime.runtimeId, outcome, wasActive });
    }
    return settled;
  }

  private emit(): void {
    const settled = this.collectSettledRuns();
    const snapshot = this.snapshot();
    for (const listener of this.listeners) {
      listener(snapshot);
    }
    for (const run of settled) {
      const title = snapshot.tabs.find((tab) => tab.runtimeId === run.runtimeId)?.title ?? "";
      this.onRunSettled({ ...run, title });
    }
  }
}
