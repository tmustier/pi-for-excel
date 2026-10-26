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

/**
 * Atomic read-modify-write of one key. Several taskpanes share the settings
 * bag, so a record that more than one of them appends to (backups, audit
 * log, file tags) must be changed with `update`: `parse` decodes the value
 * stored now (not a copy loaded earlier) and `change` returns the value to
 * store (`undefined` deletes the key). Both must be synchronous. Rejects
 * when the write does not commit (for example when storage is full).
 */
export interface SettingsUpdater extends SettingsWriter {
  update<T, U>(key: string, parse: SettingsParser<T>, change: (current: T) => U | undefined): Promise<void>;
}

/** Decodes a raw settings value (the seam where `unknown` becomes a domain type). */
export type SettingsParser<T> = (raw: unknown) => T;

export class SettingsStore extends Store implements SettingsAccess, SettingsUpdater {
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

  async update<T, U>(key: string, parse: SettingsParser<T>, change: (current: T) => U | undefined): Promise<void> {
    await this.getBackend().transaction(["settings"], "readwrite", async (tx) => {
      const next = change(parse(await tx.get<unknown>("settings", key)));
      if (next === undefined) {
        await tx.delete("settings", key);
      } else {
        await tx.set("settings", key, next);
      }
    });
  }

  async list(): Promise<string[]> {
    return this.getBackend().keys("settings");
  }

  async clear(): Promise<void> {
    await this.getBackend().clear("settings");
  }
}
