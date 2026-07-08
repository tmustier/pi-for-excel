/**
 * Store for LLM provider API keys (Anthropic, OpenAI, etc.).
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

export class ProviderKeysStore extends Store {
  getConfig(): StoreConfig {
    return {
      name: "provider-keys",
    };
  }

  async get(provider: string): Promise<string | null> {
    return this.getBackend().get<string>("provider-keys", provider);
  }

  async set(provider: string, key: string): Promise<void> {
    await this.getBackend().set("provider-keys", provider, key);
  }

  async delete(provider: string): Promise<void> {
    await this.getBackend().delete("provider-keys", provider);
  }

  async list(): Promise<string[]> {
    return this.getBackend().keys("provider-keys");
  }

  async has(provider: string): Promise<boolean> {
    return this.getBackend().has("provider-keys", provider);
  }
}
