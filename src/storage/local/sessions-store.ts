/**
 * Vendored from @earendil-works/pi-web-ui 0.75.3 (MIT, © Mario Zechner,
 * https://github.com/badlogic/pi-mono).
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

  async saveSession(
    id: string,
    state: AgentState,
    metadata: SessionMetadata | undefined,
    title?: string,
    revisions?: { base: string | undefined; next: string },
  ): Promise<boolean> {
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
      ...(meta.titleChosen !== undefined ? { titleChosen: meta.titleChosen } : {}),
      ...(meta.titleGenerationAttempted !== undefined ? { titleGenerationAttempted: meta.titleGenerationAttempted } : {}),
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
