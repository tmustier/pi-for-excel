/**
 * Store for chat sessions (data and metadata).
 * Uses two object stores: sessions (full data) and sessions-metadata (lightweight).
 *
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono). See docs/ui-ownership.md.
 */

import type { AgentState } from "@earendil-works/pi-agent-core";
import type { Api, Model } from "@earendil-works/pi-ai";

import { Store } from "./store.js";
import type { SessionData, SessionMetadata, StoreConfig } from "./types.js";

export class SessionsStore extends Store {
  getConfig(): StoreConfig {
    return {
      name: "sessions",
      keyPath: "id",
      indices: [{ name: "lastModified", keyPath: "lastModified" }],
    };
  }

  /**
   * Additional config for sessions-metadata store.
   * Must be included when creating the backend.
   */
  static getMetadataConfig(): StoreConfig {
    return {
      name: "sessions-metadata",
      keyPath: "id",
      indices: [{ name: "lastModified", keyPath: "lastModified" }],
    };
  }

  async save(data: SessionData, metadata: SessionMetadata): Promise<void> {
    await this.getBackend().transaction(
      ["sessions", "sessions-metadata"],
      "readwrite",
      async (tx) => {
        await tx.set("sessions", data.id, data);
        await tx.set("sessions-metadata", metadata.id, metadata);
      },
    );
  }

  /**
   * Saves unless the stored session has changed since `baseRevision` (the
   * revision this writer last loaded or saved; `undefined` for a session it
   * has not seen stored). Check and write are one transaction, so another
   * taskpane cannot save in between. Returns `false` on conflict, writing
   * nothing.
   */
  async saveIfUnchanged(
    data: SessionData,
    metadata: SessionMetadata,
    baseRevision: string | undefined,
  ): Promise<boolean> {
    return this.getBackend().transaction(
      ["sessions", "sessions-metadata"],
      "readwrite",
      async (tx) => {
        const stored = await tx.get<SessionMetadata>("sessions-metadata", metadata.id);
        if (stored && stored.revision !== baseRevision) return false;
        await tx.set("sessions", data.id, data);
        await tx.set("sessions-metadata", metadata.id, metadata);
        return true;
      },
    );
  }

  async get(id: string): Promise<SessionData | null> {
    return this.getBackend().get<SessionData>("sessions", id);
  }

  async getMetadata(id: string): Promise<SessionMetadata | null> {
    return this.getBackend().get<SessionMetadata>("sessions-metadata", id);
  }

  async getAllMetadata(): Promise<SessionMetadata[]> {
    // Use the lastModified index to get sessions sorted by most recent first
    return this.getBackend().getAllFromIndex<SessionMetadata>(
      "sessions-metadata",
      "lastModified",
      "desc",
    );
  }

  async delete(id: string): Promise<void> {
    await this.getBackend().transaction(
      ["sessions", "sessions-metadata"],
      "readwrite",
      async (tx) => {
        await tx.delete("sessions", id);
        await tx.delete("sessions-metadata", id);
      },
    );
  }

  async updateTitle(id: string, title: string): Promise<void> {
    const metadata = await this.getMetadata(id);
    if (metadata) {
      metadata.title = title;
      await this.getBackend().set("sessions-metadata", id, metadata);
    }
    // Also update in full session data
    const data = await this.get(id);
    if (data) {
      data.title = title;
      await this.getBackend().set("sessions", id, data);
    }
  }

  async getQuotaInfo(): Promise<{ usage: number; quota: number; percent: number }> {
    return this.getBackend().getQuotaInfo();
  }

  async requestPersistence(): Promise<boolean> {
    return this.getBackend().requestPersistence();
  }

  // Alias methods for backward compatibility
  /**
   * Saves agent state. With `revisions`, saves only over `revisions.base`
   * (see `saveIfUnchanged`) and returns `false` on conflict.
   */
  async saveSession(
    id: string,
    state: AgentState,
    metadata: SessionMetadata | undefined,
    title?: string,
    revisions?: { base: string | undefined; next: string },
  ): Promise<boolean> {
    // If metadata is provided, use it; otherwise create it from state
    const meta: SessionMetadata = metadata ?? {
      id,
      title: title ?? "",
      createdAt: new Date().toISOString(),
      lastModified: new Date().toISOString(),
      messageCount: state.messages.length,
      usage: {
        input: 0,
        output: 0,
        cacheRead: 0,
        cacheWrite: 0,
        totalTokens: 0,
        cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
      },
      thinkingLevel: state.thinkingLevel,
      preview: "",
    };
    // AgentState.model is typed Model<any> upstream; every model that enters
    // our runtime comes from pi-ai's typed registry or a parsed custom-provider
    // boundary, so this narrows an upstream `any` back to the concrete Api union.
    const model = state.model as Model<Api>;
    const data: SessionData = {
      id,
      title: title ?? meta.title,
      model,
      thinkingLevel: state.thinkingLevel,
      messages: state.messages,
      createdAt: meta.createdAt,
      lastModified: meta.lastModified,
    };
    if (!revisions) {
      await this.save(data, meta);
      return true;
    }
    return this.saveIfUnchanged(
      { ...data, revision: revisions.next },
      { ...meta, revision: revisions.next },
      revisions.base,
    );
  }

  async loadSession(id: string): Promise<SessionData | null> {
    return this.get(id);
  }

  async getLatestSessionId(): Promise<string | null> {
    const allMetadata = await this.getAllMetadata();
    return allMetadata[0]?.id ?? null;
  }
}
