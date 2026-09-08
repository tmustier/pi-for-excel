/**
 * Store for custom LLM providers (auto-discovery servers + manual providers).
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import type { Api, Model } from "@earendil-works/pi-ai";

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

export type AutoDiscoveryProviderType = "ollama" | "llama.cpp" | "vllm" | "lmstudio";

export type CustomProviderType =
  | AutoDiscoveryProviderType
  | "openai-completions"
  | "openai-responses"
  | "anthropic-messages";

export interface CustomProvider {
  id: string;
  name: string;
  type: CustomProviderType;
  baseUrl: string;
  apiKey?: string;
  models?: Model<Api>[];
}

export class CustomProvidersStore extends Store {
  getConfig(): StoreConfig {
    return {
      name: "custom-providers",
    };
  }

  async get(id: string): Promise<CustomProvider | null> {
    return this.getBackend().get<CustomProvider>("custom-providers", id);
  }

  async set(provider: CustomProvider): Promise<void> {
    await this.getBackend().set("custom-providers", provider.id, provider);
  }

  async delete(id: string): Promise<void> {
    await this.getBackend().delete("custom-providers", id);
  }

  async getAll(): Promise<CustomProvider[]> {
    const keys = await this.getBackend().keys("custom-providers");
    const providers: CustomProvider[] = [];
    for (const key of keys) {
      const provider = await this.get(key);
      if (provider) {
        providers.push(provider);
      }
    }
    return providers;
  }

  async has(id: string): Promise<boolean> {
    return this.getBackend().has("custom-providers", id);
  }
}
