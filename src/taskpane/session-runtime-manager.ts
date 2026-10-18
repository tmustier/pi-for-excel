/**
 * UI-independent lifecycle owner for taskpane session runtimes.
 */

import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { Api, Model } from "@earendil-works/pi-ai";
import type { ConversationPresenter } from "../durable/conversation-view.js";
import type { DurableSessions } from "../durable/sessions.js";
import type { AnyCoreTool } from "../tools/registry.js";

import {
  shouldForkModelSwitch,
  type ModelSwitchBehavior,
} from "../models/switch-behavior.js";
import { areRuntimeModelsEquivalent } from "../models/model-refresh-owner.js";
import type { QueueDisplay } from "./queue-display.js";
import { resolveTabTitle } from "./session-title.js";

export type RuntimeLockState = "idle" | "waiting_for_lock" | "holding_lock";

export interface SessionRuntime {
  runtimeId: string;
  view: ConversationPresenter;
  tools: AnyCoreTool[];
  queueDisplay: QueueDisplay;
  title: string;
  hasExplicitTitle: boolean;
  lockState: RuntimeLockState;
  refreshCapabilities: () => Promise<void>;
  /** Release UI projections only; never abort or close the Durable execution owner. */
  dispose: () => void;
}

export interface CreateRuntimeOptions {
  activate: boolean;
  autoRestoreLatest: boolean;
  conversationId?: string;
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
export type ModelSelectionResult =
  | { outcome: "missing" | "unchanged" | "busy" | "updated" }
  | { outcome: "forked"; runtime: SessionRuntime; title: string };

interface RuntimeListeners {
  unsubscribeView: () => void;
}

export class SessionRuntimeManager {
  private readonly createSessionRuntime: SessionRuntimeFactory;
  private readonly sessions: DurableSessions;
  private readonly warnCapabilityRefresh: (error: unknown) => void;
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
    sessions: DurableSessions;
    warnCapabilityRefresh?: (error: unknown) => void;
  }) {
    this.createSessionRuntime = opts.createRuntime;
    this.sessions = opts.sessions;
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

    const unsubscribeView = runtime.view.subscribe(() => {
      this.emit();
    });
    this.runtimeListeners.set(runtime.runtimeId, { unsubscribeView });

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
    listeners?.unsubscribeView();
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

  async selectModel(args: {
    runtimeId: string;
    nextModel: Model<Api>;
    behavior: ModelSwitchBehavior;
  }): Promise<ModelSelectionResult> {
    const runtime = this.runtimes.get(args.runtimeId);
    if (!runtime) return { outcome: "missing" };

    const currentModel = runtime.view.current.model;
    const sameIdentity = currentModel.provider === args.nextModel.provider
      && currentModel.id === args.nextModel.id;
    if (sameIdentity && areRuntimeModelsEquivalent(currentModel, args.nextModel)) {
      return { outcome: "unchanged" };
    }

    if (runtime.view.current.busy || runtime.view.current.inbox.length > 0) {
      return { outcome: "busy" };
    }

    if (sameIdentity || !shouldForkModelSwitch({
      behavior: args.behavior,
      hasMessages: runtime.view.current.messages.length > 0,
    })) {
      await runtime.view.conversation.configure({
        model: { provider: args.nextModel.provider, modelId: args.nextModel.id },
      }, context);
      await runtime.view.refresh();
      return { outcome: "updated" };
    }

    const sourceTitle = this.snapshotTabs()
      .find((tab) => tab.runtimeId === args.runtimeId)?.title ?? "Untitled";
    const title = `${sourceTitle} (${args.nextModel.id})`;
    const fork = await this.sessions.fork(String(runtime.view.conversation.id), args.nextModel, title);
    const forkedRuntime = await this.createRuntime({
      activate: true,
      autoRestoreLatest: false,
      conversationId: String(fork.id),
    });

    return { outcome: "forked", runtime: forkedRuntime, title };
  }

  async renameRuntime(runtimeId: string, title: string): Promise<void> {
    const runtime = this.runtimes.get(runtimeId);
    if (!runtime) return;
    await this.sessions.renameSession(String(runtime.view.conversation.id), title);
    runtime.title = title;
    runtime.hasExplicitTitle = true;
    this.emit();
  }

  refreshCapabilities(): Promise<void> {
    this.refreshRequested = true;
    if (this.refreshPromise) return this.refreshPromise;

    this.refreshPromise = this.runCapabilityRefreshes().finally(() => {
      // Even an empty pass must release the slot after its Promise is assigned.
      // Release before notifying listeners so they can request a fresh pass.
      this.refreshPromise = null;
      // A request may arrive after the loop drains but before this continuation.
      if (this.refreshRequested) return this.refreshCapabilities();
      return this.emit();
    });
    return this.refreshPromise;
  }

  findRuntimeBySessionId(sessionId: string): SessionRuntime | null {
    for (const runtimeId of this.runtimeOrder) {
      const runtime = this.runtimes.get(runtimeId);
      if (!runtime) continue;
      if (String(runtime.view.conversation.id) === sessionId) {
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
        hasExplicitTitle: runtime.hasExplicitTitle,
        sessionTitle: runtime.title,
        defaultTabNumber: this.runtimeDefaultTabNumbers.get(runtime.runtimeId) ?? (index + 1),
      }),
      isActive: runtime.runtimeId === this.activeRuntimeId,
      isStreaming: runtime.view.current.busy,
      isBusy: runtime.view.current.busy || runtime.view.current.inbox.length > 0,
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
  }

  private emit(): void {
    const snapshot = this.snapshot();
    for (const listener of this.listeners) {
      listener(snapshot);
    }
  }
}
