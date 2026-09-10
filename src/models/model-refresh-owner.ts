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

/**
 * Canonical JSON for a model: keys sorted at every depth, `undefined` members
 * dropped. Models are JSON-shaped records (that is how they persist inside
 * sessions), so this totally orders everything that reaches the wire or the UI.
 */
function canonicalModelJson(value: DynamicValue): string {
  if (Array.isArray(value)) {
    return `[${value.map((item: DynamicValue) => canonicalModelJson(item)).join(",")}]`;
  }
  if (typeof value === "object" && value !== null) {
    // Non-array object: models only nest plain records (cost, thinking levels, headers, compat).
    const record = value as Record<string, DynamicValue>;
    const members = Object.keys(record).sort()
      .filter((key) => record[key] !== undefined)
      .map((key) => `${JSON.stringify(key)}:${canonicalModelJson(record[key])}`);
    return `{${members.join(",")}}`;
  }
  if (value === undefined) return "null";
  return JSON.stringify(value);
}

/**
 * Whole-model equality. Every field on a `Model` reaches the wire or the UI
 * (reasoning, thinking levels, input modalities, cost, headers, compat, ...),
 * so an active session must pick up any difference, not just identity and
 * token limits.
 */
export function areRuntimeModelsEquivalent(left: RuntimeModelShape, right: RuntimeModelShape): boolean {
  return canonicalModelJson(left) === canonicalModelJson(right);
}

// Structural identity floor for callers holding `Model<any>` (agent state); the
// comparison itself covers every member of the object, not just these.
type RuntimeModelShape = Pick<Model<Api>, "id" | "provider" | "api">;

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
    if (request.syncConfiguredProviders) {
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
