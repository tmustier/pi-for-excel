import type { Agent } from "@earendil-works/pi-agent-core";
import { clampThinkingLevel } from "@earendil-works/pi-ai";

import { areRuntimeModelsEquivalent } from "../models/model-refresh-owner.js";
import type { ActionQueue } from "./action-queue.js";
import type { QueueDisplay } from "./queue-display.js";
import type { SessionData } from "../storage/local/types.js";
import { countVisibleSessionMessages, type SessionPersistenceController } from "./sessions.js";
import { resolveTabTitle } from "./session-title.js";
import type { TabState } from "./tab-state.js";

export type RuntimeLockState = "idle" | "waiting_for_lock" | "holding_lock";

export type RuntimeRunOutcome = "finished" | "failed";

export type RuntimeAttention = "none" | RuntimeRunOutcome;

export interface RuntimeRunSettledEvent {
  runtimeId: string;
  title: string;
  outcome: RuntimeRunOutcome;
  wasActive: boolean;
}

export interface SessionRuntime {
  runtimeId: string;
  agent: Agent;
  actionQueue: ActionQueue;
  queueDisplay: QueueDisplay;
  tab: TabState;
  persistence: SessionPersistenceController;
  lockState: RuntimeLockState;
  refreshCapabilities: () => Promise<void>;
  dispose: () => void;
}

export interface CreateRuntimeOptions {
  activate: boolean;
  autoRestoreLatest: boolean;
  sessionId?: string;
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
export type ModelSelectionResult = { outcome: "missing" | "unchanged" | "updated" | "next_turn" };

export type OpenSessionTarget =
  | { kind: "new_tab"; tabNumber?: number }
  | { kind: "resume" };

export type OpenSessionResult = { outcome: "opened" | "replaced" | "focused"; runtime: SessionRuntime };

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
    getRunOutcome?: (runtime: SessionRuntime) => RuntimeRunOutcome | null;
    onRunSettled?: (event: RuntimeRunSettledEvent) => void;
  }) {
    this.createSessionRuntime = opts.createRuntime;
    this.warnCapabilityRefresh = opts.warnCapabilityRefresh ?? (() => {});
    this.getRunOutcome = opts.getRunOutcome
      ?? ((runtime) => (runtime.agent.state.errorMessage ? "failed" : "finished"));
    this.onRunSettled = opts.onRunSettled ?? (() => {});
  }

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

  async openSession(sessionData: SessionData, target: OpenSessionTarget): Promise<OpenSessionResult> {
    const existing = this.findRuntimeBySessionId(sessionData.id);
    if (existing) {
      this.switchRuntime(existing.runtimeId);
      return { outcome: "focused", runtime: existing };
    }

    if (target.kind === "resume") {
      const runtime = this.getActiveRuntime();
      if (runtime && !isRuntimeBusy(runtime) && countVisibleSessionMessages(runtime.agent.state.messages) === 0) {
        await runtime.persistence.applyLoadedSession(sessionData);
        this.switchRuntime(runtime.runtimeId);
        return { outcome: "replaced", runtime };
      }
    }

    const runtime = await this.createRuntime({
      activate: true,
      autoRestoreLatest: false,
      ...(target.kind === "new_tab" && target.tabNumber !== undefined ? { tabNumber: target.tabNumber } : {}),
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
  }): Promise<ModelSelectionResult> {
    const runtime = this.runtimes.get(args.runtimeId);
    if (!runtime) return { outcome: "missing" };

    const currentModel = runtime.agent.state.model;
    const sameIdentity = currentModel.provider === args.nextModel.provider
      && currentModel.id === args.nextModel.id;
    if (sameIdentity && areRuntimeModelsEquivalent(currentModel, args.nextModel)) {
      return { outcome: "unchanged" };
    }

    const busy = isRuntimeBusy(runtime);
    const thinkingLevel = clampThinkingLevel(args.nextModel, runtime.agent.state.thinkingLevel);
    runtime.agent.state.model = args.nextModel;
    runtime.agent.state.thinkingLevel = thinkingLevel;
    await runtime.persistence.saveSession({ force: true });
    this.emit();
    return { outcome: busy ? "next_turn" : "updated" };
  }

  refreshCapabilities(): Promise<void> {
    this.refreshRequested = true;
    if (this.refreshPromise) return this.refreshPromise;

    const pass = this.runCapabilityRefreshes();
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
    return this.listRuntimes().map((runtime, index) => {
      const tabNumber = this.runtimeDefaultTabNumbers.get(runtime.runtimeId) ?? (index + 1);
      return {
        runtimeId: runtime.runtimeId,
        sessionId: runtime.persistence.getSessionId(),
        tabNumber,
        title: resolveTabTitle({
          sessionTitle: runtime.persistence.getSessionTitle(),
          defaultTabNumber: tabNumber,
        }),
        isActive: runtime.runtimeId === this.activeRuntimeId,
        isStreaming: runtime.agent.state.isStreaming,
        isBusy: isRuntimeBusy(runtime),
        lockState: runtime.lockState,
        attention: this.attention.get(runtime.runtimeId) ?? "none",
      };
    });
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
      this.refreshPromise = null;
      this.refreshPassRunning = false;
    }

    this.emit();
  }

  private emit(): void {
    const settled = new Map<string, RuntimeRunOutcome>();
    for (const runtime of this.runtimes.values()) {
      const busy = isRuntimeBusy(runtime);
      const wasBusy = this.wasBusy.get(runtime.runtimeId) ?? false;
      this.wasBusy.set(runtime.runtimeId, busy);
      if (busy) {
        this.attention.delete(runtime.runtimeId);
        continue;
      }
      if (!wasBusy) continue;

      const outcome = this.getRunOutcome(runtime);
      if (!outcome) continue;
      if (runtime.runtimeId !== this.activeRuntimeId) this.attention.set(runtime.runtimeId, outcome);
      settled.set(runtime.runtimeId, outcome);
    }

    const snapshot = this.snapshot();
    for (const listener of this.listeners) {
      listener(snapshot);
    }
    for (const tab of snapshot.tabs) {
      const outcome = settled.get(tab.runtimeId);
      if (outcome) this.onRunSettled({ runtimeId: tab.runtimeId, title: tab.title, outcome, wasActive: tab.isActive });
    }
  }
}
