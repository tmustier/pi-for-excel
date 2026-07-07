/**
 * Store for application settings (theme, proxy config, etc.).
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

export class SettingsStore extends Store {
  getConfig(): StoreConfig {
    return {
      name: "settings",
      // No keyPath - uses out-of-line keys
    };
  }

  async get<T = DynamicValue>(key: string): Promise<T | null> {
    return this.getBackend().get<T>("settings", key);
  }

  async set<T = DynamicValue>(key: string, value: T): Promise<void> {
    await this.getBackend().set("settings", key, value);
  }

  async delete(key: string): Promise<void> {
    await this.getBackend().delete("settings", key);
  }

  async list(): Promise<string[]> {
    return this.getBackend().keys("settings");
  }

  async clear(): Promise<void> {
    await this.getBackend().clear("settings");
  }
}
