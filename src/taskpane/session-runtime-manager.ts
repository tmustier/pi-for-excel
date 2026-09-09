/**
 * UI-independent lifecycle owner for taskpane session runtimes.
 */

import type { Agent } from "@earendil-works/pi-agent-core";

import type { ActionQueue } from "./action-queue.js";
import type { QueueDisplay } from "./queue-display.js";
import type { SessionPersistenceController } from "./sessions.js";
import { resolveTabTitle } from "./session-title.js";

export type RuntimeLockState = "idle" | "waiting_for_lock" | "holding_lock";

export interface SessionRuntime {
  runtimeId: string;
  agent: Agent;
  actionQueue: ActionQueue;
  queueDisplay: QueueDisplay;
  persistence: SessionPersistenceController;
  lockState: RuntimeLockState;
  refreshCapabilities: () => Promise<void>;
  dispose: () => void;
}

export interface CreateRuntimeOptions {
  activate: boolean;
  autoRestoreLatest: boolean;
}

export interface RuntimeTabSnapshot {
  runtimeId: string;
  title: string;
  isActive: boolean;
  isStreaming: boolean;
  isBusy: boolean;
  lockState: RuntimeLockState;
}

export interface RuntimeLifecycleSnapshot {
  activeRuntimeId: string | null;
  tabs: RuntimeTabSnapshot[];
}

export type RuntimeSnapshotListener = (snapshot: RuntimeLifecycleSnapshot) => void;
export type SessionRuntimeFactory = (opts: CreateRuntimeOptions) => Promise<SessionRuntime>;

interface RuntimeListeners {
  unsubscribeAgent: () => void;
  unsubscribePersistence: () => void;
}

export class SessionRuntimeManager {
  private readonly createSessionRuntime: SessionRuntimeFactory;
  private readonly warnCapabilityRefresh: (error: DynamicValue) => void;
  private readonly runtimes = new Map<string, SessionRuntime>();
  private readonly runtimeOrder: string[] = [];
  private readonly runtimeDefaultTabNumbers = new Map<string, number>();
  private readonly runtimeListeners = new Map<string, RuntimeListeners>();
  private readonly listeners = new Set<RuntimeSnapshotListener>();

  private activeRuntimeId: string | null = null;
  private nextDefaultTabNumber = 1;
  private refreshPromise: Promise<void> | null = null;
  private refreshRequested = false;

  constructor(opts: {
    createRuntime: SessionRuntimeFactory;
    warnCapabilityRefresh?: (error: DynamicValue) => void;
  }) {
    this.createSessionRuntime = opts.createRuntime;
    this.warnCapabilityRefresh = opts.warnCapabilityRefresh ?? (() => {});
  }

  async createRuntime(opts: CreateRuntimeOptions): Promise<SessionRuntime> {
    const runtime = await this.createSessionRuntime(opts);
    return this.registerRuntime(runtime, { activate: opts.activate });
  }

  registerRuntime(runtime: SessionRuntime, opts?: { activate?: boolean }): SessionRuntime {
    this.runtimes.set(runtime.runtimeId, runtime);
    this.runtimeOrder.push(runtime.runtimeId);
    this.runtimeDefaultTabNumbers.set(runtime.runtimeId, this.nextDefaultTabNumber);
    this.nextDefaultTabNumber += 1;

    const unsubscribeAgent = runtime.agent.subscribe(() => {
      this.emit();
    });
    const unsubscribePersistence = runtime.persistence.subscribe(() => {
      this.emit();
    });
    this.runtimeListeners.set(runtime.runtimeId, { unsubscribeAgent, unsubscribePersistence });

    const shouldActivate = opts?.activate ?? this.activeRuntimeId === null;
    if (shouldActivate) {
      this.switchRuntime(runtime.runtimeId);
    } else {
      this.emit();
    }

    return runtime;
  }

  switchRuntime(runtimeId: string): SessionRuntime | null {
    const next = this.runtimes.get(runtimeId);
    if (!next) return null;

    this.activeRuntimeId = runtimeId;
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
    this.runtimeListeners.delete(runtimeId);

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

  refreshCapabilities(): Promise<void> {
    this.refreshRequested = true;
    if (this.refreshPromise) return this.refreshPromise;

    this.refreshPromise = this.runCapabilityRefreshes();
    return this.refreshPromise;
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
    }

    this.emit();
  }

  private emit(): void {
    const snapshot = this.snapshot();
    for (const listener of this.listeners) {
      listener(snapshot);
    }
  }
}
