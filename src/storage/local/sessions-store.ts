/**
 * Read-only migration source for pre-Durable chat snapshots.
 * Keep both original object stores intact; all new session writes use Durable.
 *
 * Adapted from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import { Store } from "./store.js";
import type { SessionData, StoreConfig } from "./types.js";

export class SessionsStore extends Store {
  getConfig(): StoreConfig {
    return {
      name: "sessions",
      keyPath: "id",
      indices: [{ name: "lastModified", keyPath: "lastModified" }],
    };
  }

  /** Preserve the existing schema without deleting original metadata. */
  static getMetadataConfig(): StoreConfig {
    return {
      name: "sessions-metadata",
      keyPath: "id",
      indices: [{ name: "lastModified", keyPath: "lastModified" }],
    };
  }

  async get(id: string): Promise<SessionData | null> {
    return this.getBackend().get<SessionData>("sessions", id);
  }

  /** Include original snapshots whose metadata row is missing. */
  async getSessionIds(): Promise<string[]> {
    return this.getBackend().keys("sessions");
  }
}
