/**
 * Store for application settings (theme, proxy config, etc.).
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

/**
 * Read side of the settings bag. Settings are a key-value bag written by many
 * features, so a read is `unknown` by contract: the feature that owns the key
 * parses it at this seam. Features depend on this contract rather than the
 * class so tests can pass an in-memory store.
 */
export interface SettingsReader {
  get(key: string): Promise<unknown>;
}

/** Read and write side of the settings bag. */
export interface SettingsWriter extends SettingsReader {
  set(key: string, value: unknown): Promise<void>;
}

/** Full access to the settings bag, including key removal. */
export interface SettingsAccess extends SettingsWriter {
  delete(key: string): Promise<void>;
}

export class SettingsStore extends Store implements SettingsAccess {
  getConfig(): StoreConfig {
    return {
      name: "settings",
      // No keyPath - uses out-of-line keys
    };
  }

  async get(key: string): Promise<unknown> {
    return this.getBackend().get<unknown>("settings", key);
  }

  async set(key: string, value: unknown): Promise<void> {
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
