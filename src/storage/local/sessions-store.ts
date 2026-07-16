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

  // Alias for backward compatibility
  async deleteSession(id: string): Promise<void> {
    return this.delete(id);
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
  async saveSession(
    id: string,
    state: AgentState,
    metadata: SessionMetadata | undefined,
    title?: string,
  ): Promise<void> {
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
      lastModified: new Date().toISOString(),
    };
    await this.save(data, meta);
  }

  async loadSession(id: string): Promise<SessionData | null> {
    return this.get(id);
  }

  async getLatestSessionId(): Promise<string | null> {
    const allMetadata = await this.getAllMetadata();
    // Sort by lastModified descending
    allMetadata.sort((a, b) => b.lastModified.localeCompare(a.lastModified));
    return allMetadata[0]?.id ?? null;
  }
}
