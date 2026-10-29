import type { Api, Model, MutableModels } from "@earendil-works/pi-ai";
import type { ThinkingLevel } from "@earendil-works/pi-agent-core";

import { resolveCustomProviderModel } from "../auth/custom-gateways.js";
import type { CustomProvider } from "../storage/local/custom-providers-store.js";
import { pickDefaultModel } from "../taskpane/default-model.js";
import { resolveRuntimeModelSwap } from "../taskpane/runtime-model-reconcile.js";
import type { BrowserModelRuntime } from "./browser-model-runtime.js";
import { canonicalJson, toJsonValue } from "../utils/json.js";

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

export function areRuntimeModelsEquivalent(left: RuntimeModelShape, right: RuntimeModelShape): boolean {
  return canonicalJson(toJsonValue(left)) === canonicalJson(toJsonValue(right));
}

type RuntimeModelShape = Pick<Model<Api>, "id" | "provider" | "api">;

export class ModelRefreshOwner {
  readonly models: MutableModels;

  private readonly modelRuntime: BrowserModelRuntime;
  private readonly loadCustomProviders: () => Promise<readonly CustomProvider[]>;
  private readonly getRuntimes: () => readonly ModelRefreshRuntime[];
  private readonly listeners = new Set<ModelRefreshListener>();
  private readonly warn: (message: string, error?: unknown) => void;

  private currentSnapshot: ModelRefreshSnapshot;
  private customProviders: readonly CustomProvider[] = [];
  private refreshPromise: Promise<void> | null = null;
  private requestedSync = false;
  private requestedNetwork = false;
  private activeNetwork = false;
  private activeProviderRevision = 0;

  constructor(options: {
    modelRuntime: BrowserModelRuntime;
    loadCustomProviders: () => Promise<readonly CustomProvider[]>;
    getRuntimes: () => readonly ModelRefreshRuntime[];
    warn?: (message: string, error?: unknown) => void;
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

  restoreCached(): Promise<void> {
    return this.requestRefresh({ syncConfiguredProviders: true, allowNetwork: false });
  }

  refreshConfiguredProviders(allowNetwork = true): Promise<void> {
    return this.requestRefresh({ syncConfiguredProviders: true, allowNetwork });
  }

  refresh(allowNetwork = true): Promise<void> {
    return this.requestRefresh({ syncConfiguredProviders: false, allowNetwork });
  }

  resolveModel(persisted: Model<Api>): Model<Api> | undefined {
    const direct = this.models.getModel(persisted.provider, persisted.id);
    if (direct) return direct;

    const custom = resolveCustomProviderModel([...this.customProviders], persisted);
    if (!custom) return undefined;
    return this.models.getModel(custom.provider, custom.id) ?? undefined;
  }

  reconcileRuntimes(): void {
    this.reconcileRuntimeModels();
  }

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
          (error: unknown) => {
            this.warn("[auth] Credential restore failed:", error);
            return "failed" as const;
          },
        ),
        timeout,
      ]);

      if (outcome === "timeout") {
        void credentialRestore.then(
          () => this.refreshConfiguredProviders(true),
          (error: unknown) => {
            this.warn("[auth] Credential restore failed after timeout:", error);
          },
        );
      }

      void this.refresh(true).catch((error: unknown) => {
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
    if (
      request.allowNetwork
      && (!this.activeNetwork || this.requestedSync || this.providersChangedDuringPass())
    ) {
      this.requestedNetwork = true;
    }

    if (this.refreshPromise) return this.refreshPromise;

    this.refreshPromise = this.runRefreshes().finally(() => {
      this.refreshPromise = null;
      this.activeNetwork = false;
    });
    return this.refreshPromise;
  }

  private providersChangedDuringPass(): boolean {
    return this.modelRuntime.providerRevision !== this.activeProviderRevision;
  }

  private async runRefreshes(): Promise<void> {
    do {
      const syncConfiguredProviders = this.requestedSync;
      const allowNetwork = this.requestedNetwork;
      this.requestedSync = false;
      this.requestedNetwork = false;
      this.activeNetwork = allowNetwork;

      if (syncConfiguredProviders) {
        const customProviders = await this.loadCustomProviders();
        this.customProviders = customProviders;
        await this.modelRuntime.syncCustomProviders(customProviders);
      }
      this.activeProviderRevision = this.modelRuntime.providerRevision;

      await this.publishCachedCatalogues();
      if (allowNetwork) await this.publishDiscoveredCatalogues();

      this.activeNetwork = false;
    } while (this.requestedSync || this.requestedNetwork || this.providersChangedDuringPass());
  }

  private async publishCachedCatalogues(): Promise<void> {
    const cached = await this.modelRuntime.refresh({ allowNetwork: false });
    await this.publish(Array.from(cached.errors.keys()).sort());
  }

  private async publishDiscoveredCatalogues(): Promise<void> {
    const discovered = await this.modelRuntime.refresh({ allowNetwork: true, force: true });
    await this.publish(Array.from(discovered.errors.keys()).sort());
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

      const refreshedModel = this.resolveModel(runtime.model);
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
