import type { Api, Model, MutableModels } from "@earendil-works/pi-ai";
import type { ThinkingLevel } from "@earendil-works/pi-agent-core";

import type { CustomProvider } from "../storage/local/custom-providers-store.js";
import { pickDefaultModel } from "../taskpane/default-model.js";
import { resolveRuntimeModelSwap } from "../taskpane/runtime-model-reconcile.js";
import type { BrowserModelRuntime } from "./browser-model-runtime.js";

export interface ModelRefreshRuntime {
  runtimeId: string;
  model: Model<Api>;
  isBusy: boolean;
  applyModel: (model: Model<Api>, thinkingLevel?: ThinkingLevel) => void;
}

export interface ModelRefreshSnapshot {
  availableProviders: readonly string[];
  defaultModel: Model<Api>;
  providerErrors: readonly string[];
  revision: number;
}

export type ModelRefreshListener = (snapshot: ModelRefreshSnapshot) => void;

interface RefreshRequest {
  syncConfiguredProviders: boolean;
  allowNetwork: boolean;
}

interface ComparableRuntimeModel {
  api: string;
  id: string;
  provider: string;
  baseUrl: string;
  contextWindow: number;
  maxTokens: number;
}

export function areRuntimeModelsEquivalent(
  left: ComparableRuntimeModel,
  right: ComparableRuntimeModel,
): boolean {
  return left.api === right.api
    && left.id === right.id
    && left.provider === right.provider
    && left.baseUrl === right.baseUrl
    && left.contextWindow === right.contextWindow
    && left.maxTokens === right.maxTokens;
}

/** Owns provider catalogue refresh, publication, coalescing and runtime reconciliation. */
export class ModelRefreshOwner {
  readonly models: MutableModels;

  private readonly modelRuntime: BrowserModelRuntime;
  private readonly loadCustomProviders: () => Promise<readonly CustomProvider[]>;
  private readonly getRuntimes: () => readonly ModelRefreshRuntime[];
  private readonly listeners = new Set<ModelRefreshListener>();
  private readonly warn: (message: string, error?: DynamicValue) => void;

  private currentSnapshot: ModelRefreshSnapshot;
  private refreshPromise: Promise<void> | null = null;
  private requestedSync = false;
  private requestedNetwork = false;
  private activeSync = false;
  private activeNetwork = false;

  constructor(options: {
    modelRuntime: BrowserModelRuntime;
    loadCustomProviders: () => Promise<readonly CustomProvider[]>;
    getRuntimes: () => readonly ModelRefreshRuntime[];
    warn?: (message: string, error?: DynamicValue) => void;
  }) {
    this.modelRuntime = options.modelRuntime;
    this.models = options.modelRuntime.models;
    this.loadCustomProviders = options.loadCustomProviders;
    this.getRuntimes = options.getRuntimes;
    this.warn = options.warn ?? (() => {});
    this.currentSnapshot = {
      availableProviders: [],
      defaultModel: pickDefaultModel(this.models, [], null),
      providerErrors: [],
      revision: 0,
    };
  }

  snapshot(): ModelRefreshSnapshot {
    return this.currentSnapshot;
  }

  subscribe(listener: ModelRefreshListener): () => void {
    this.listeners.add(listener);
    listener(this.currentSnapshot);
    return () => {
      this.listeners.delete(listener);
    };
  }

  /** Restore configured providers and cached catalogues without touching the network. */
  restoreCached(): Promise<void> {
    return this.requestRefresh({ syncConfiguredProviders: true, allowNetwork: false });
  }

  /** Re-read custom-provider configuration, then refresh its catalogues. */
  refreshConfiguredProviders(allowNetwork = true): Promise<void> {
    return this.requestRefresh({ syncConfiguredProviders: true, allowNetwork });
  }

  /** Refresh all currently registered providers. Concurrent calls share one provider pass. */
  refresh(allowNetwork = true): Promise<void> {
    return this.requestRefresh({ syncConfiguredProviders: false, allowNetwork });
  }

  /** Reconcile a runtime that was deliberately skipped while busy. */
  reconcileRuntimes(): void {
    this.reconcileRuntimeModels();
  }

  /**
   * Restore cached catalogues first, then perform discovery once credentials are ready.
   * A restore that settles after the startup bound schedules exactly one post-restore pass.
   */
  async startup(credentialRestore: Promise<void>, timeoutMs: number): Promise<void> {
    await this.restoreCached();

    let timeoutId: ReturnType<typeof setTimeout> | null = null;
    const timeout = new Promise<"timeout">((resolve) => {
      timeoutId = setTimeout(() => resolve("timeout"), timeoutMs);
    });

    try {
      const outcome = await Promise.race([
        credentialRestore.then(
          () => "restored" as const,
          (error: DynamicValue) => {
            this.warn("[auth] Credential restore failed:", error);
            return "failed" as const;
          },
        ),
        timeout,
      ]);

      if (outcome === "timeout") {
        void credentialRestore.then(
          () => this.refreshConfiguredProviders(true),
          (error: DynamicValue) => {
            this.warn("[auth] Credential restore failed after timeout:", error);
          },
        );
      }

      void this.refresh(true).catch((error: DynamicValue) => {
        this.warn("[models] Background model refresh failed:", error);
      });
    } finally {
      if (timeoutId !== null) clearTimeout(timeoutId);
    }
  }

  private requestRefresh(request: RefreshRequest): Promise<void> {
    if (request.syncConfiguredProviders && !this.activeSync) {
      this.requestedSync = true;
    }
    if (request.allowNetwork && (!this.activeNetwork || this.requestedSync)) {
      this.requestedNetwork = true;
    }

    if (this.refreshPromise) return this.refreshPromise;

    this.refreshPromise = this.runRefreshes().finally(() => {
      this.refreshPromise = null;
      this.activeSync = false;
      this.activeNetwork = false;
    });
    return this.refreshPromise;
  }

  private async runRefreshes(): Promise<void> {
    do {
      const syncConfiguredProviders = this.requestedSync;
      const allowNetwork = this.requestedNetwork;
      this.requestedSync = false;
      this.requestedNetwork = false;
      this.activeSync = syncConfiguredProviders;
      this.activeNetwork = allowNetwork;

      if (syncConfiguredProviders) {
        const customProviders = await this.loadCustomProviders();
        await this.modelRuntime.syncCustomProviders(customProviders);
      }

      const result = await this.modelRuntime.refresh({
        allowNetwork,
        ...(allowNetwork ? { force: true } : {}),
      });
      await this.publish(Array.from(result.errors.keys()).sort());

      this.activeSync = false;
      this.activeNetwork = false;
    } while (this.requestedSync || this.requestedNetwork);
  }

  private async publish(providerErrors: readonly string[]): Promise<void> {
    const availableModels = await this.models.getAvailable();
    const availableProviders = Array.from(
      new Set(availableModels.map((model) => model.provider)),
    );
    const defaultModel = pickDefaultModel(this.models, availableProviders, null);

    this.currentSnapshot = {
      availableProviders,
      defaultModel,
      providerErrors: [...providerErrors],
      revision: this.currentSnapshot.revision + 1,
    };
    this.reconcileRuntimeModels();

    for (const listener of this.listeners) {
      listener(this.currentSnapshot);
    }
  }

  private reconcileRuntimeModels(): void {
    const { availableProviders, defaultModel } = this.currentSnapshot;

    for (const runtime of this.getRuntimes()) {
      if (runtime.isBusy) continue;

      const refreshedModel = this.models.getModel(runtime.model.provider, runtime.model.id);
      if (refreshedModel && !areRuntimeModelsEquivalent(runtime.model, refreshedModel)) {
        runtime.applyModel(refreshedModel);
        continue;
      }

      const swap = resolveRuntimeModelSwap({
        currentModel: runtime.model,
        availableProviders,
        defaultModel,
        isBusy: runtime.isBusy,
      });
      if (swap && !areRuntimeModelsEquivalent(runtime.model, swap.model)) {
        runtime.applyModel(swap.model, swap.thinkingLevel);
      }
    }
  }
}
