/**
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono).
 */

import { Store } from "./store.js";
import type { StoreConfig } from "./types.js";

export interface SettingsReader {
  get(key: string): Promise<unknown>;
}

export interface SettingsWriter extends SettingsReader {
  set(key: string, value: unknown): Promise<void>;
}

export interface SettingsAccess extends SettingsWriter {
  delete(key: string): Promise<void>;
}

export interface SettingsUpdater extends SettingsWriter {
  update<T, U>(key: string, parse: SettingsParser<T>, change: (current: T) => U | undefined): Promise<void>;
}

export type SettingsParser<T> = (raw: unknown) => T;

export class SettingsStore extends Store implements SettingsAccess, SettingsUpdater {
  getConfig(): StoreConfig {
    return {
      name: "settings",
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
