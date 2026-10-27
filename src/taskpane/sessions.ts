/**
 * Session persistence wiring for one runtime.
 *
 * Owns:
 * - auto-saving agent state to IndexedDB
 * - optional latest-session restore on startup
 * - session identity lifecycle (new / rename / resume)
 */

import type { Api, Model } from "@earendil-works/pi-ai";
import { clampThinkingLevel } from "@earendil-works/pi-ai";
import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";
import type { SessionData, SessionMetadata } from "../storage/local/types.js";
import type { SessionsStore } from "../storage/local/sessions-store.js";
import type { SettingsStore } from "../storage/local/settings-store.js";
import type { SpreadsheetHost } from "../host/types.js";

import { getCurrentSpreadsheetHost } from "../host/index.js";
import { extractTextFromContent } from "../utils/content.js";
import { getErrorMessage } from "../utils/errors.js";

export interface SessionPersistenceController {
  getSessionId: () => string;
  getSessionTitle: () => string;
  getSessionCreatedAt: () => string;
  /** True when the user chose the title (rename, `/name`, duplicate, fork). Persisted with the session. */
  hasExplicitTitle: () => boolean;
  startNewSession: () => void;
  /** Sets a chosen title (empty clears it) and saves the session. */
  renameSession: (title: string) => Promise<void>;
  applyLoadedSession: (sessionData: SessionData) => Promise<void>;
  restoreLatestSession: () => Promise<boolean>;
  /**
   * Saves the session once it holds a user message, or always with `force`.
   * Resolves false when there was nothing to save or the save failed; failures
   * are also reported through `onSaveError`.
   */
  saveSession: (opts?: { force?: boolean }) => Promise<boolean>;
  subscribe: (listener: () => void) => () => void;
  dispose: () => void;
}

type SessionId = string;

type UserLikeMessage = AgentMessage & {
  role: "user" | "user-with-attachments";
  content: unknown;
};

/**
 * Resolves a persisted model against the current registry so metadata such as
 * `contextWindow` picks up upstream changes, and a renamed custom gateway is
 * found again. Returns undefined when the registry no longer has the model.
 */
export type PersistedModelResolver = (persisted: Model<Api>) => Model<Api> | undefined;

function isChatUserMessage(message: AgentMessage): message is UserLikeMessage {
  if (!(message.role === "user" || message.role === "user-with-attachments")) {
    return false;
  }

  return "content" in message;
}

function isSessionId(value: string): value is SessionId {
  return value.split("-").length === 5;
}

function normalizeSessionId(value: string | null): string | null {
  if (typeof value !== "string") return null;

  const trimmed = value.trim();
  return trimmed.length > 0 ? trimmed : null;
}

/**
 * Messages the user sees in the transcript. Hidden instruction (`system`)
 * messages and internal bookkeeping roles are not counted, so a blank session
 * has zero messages.
 */
export function countVisibleSessionMessages(messages: readonly AgentMessage[]): number {
  let count = 0;
  for (const message of messages) {
    if (
      message.role === "user"
      || message.role === "user-with-attachments"
      || message.role === "assistant"
      || message.role === "toolResult"
    ) {
      count += 1;
    }
  }
  return count;
}

export function getRestoreCandidateSessionIds(args: {
  workbookId: string | null;
  workbookLatestSessionId: string | null;
  globalLatestSessionId: string | null;
}): string[] {
  const candidates: string[] = [];

  const pushUnique = (sessionId: string | null) => {
    if (!sessionId || candidates.includes(sessionId)) return;
    candidates.push(sessionId);
  };

  const workbookLatest = normalizeSessionId(args.workbookLatestSessionId);
  const globalLatest = normalizeSessionId(args.globalLatestSessionId);

  if (args.workbookId) {
    // Workbook is known: restore only the workbook-linked latest session.
    pushUnique(workbookLatest);
    return candidates;
  }

  // Workbook identity unavailable: fall back to global latest behavior.
  pushUnique(globalLatest);
  return candidates;
}

export async function setupSessionPersistence(opts: {
  agent: Agent;
  sessions: SessionsStore;
  settings: SettingsStore;
  resolveModel: PersistedModelResolver;
  initialSessionId?: string;
  autoRestoreLatest?: boolean;
  spreadsheetHost?: Pick<SpreadsheetHost, "getWorkbookContext" | "sessionStorage">;
  /** Reports a save that did not reach storage (quota, blocked IndexedDB). */
  onSaveError?: (message: string) => void;
  /**
   * Called when another taskpane saved this session since this one loaded
   * it: this taskpane's copy was saved as a new session instead of
   * replacing theirs.
   */
  onSavedAsCopy?: (args: { fromSessionId: string; toSessionId: string; title: string }) => void;
}): Promise<SessionPersistenceController> {
  const { agent, sessions, settings } = opts;
  const spreadsheetHost = opts.spreadsheetHost ?? getCurrentSpreadsheetHost();

  async function resolveWorkbookId(): Promise<string | null> {
    const ctx = await spreadsheetHost.getWorkbookContext();
    return ctx.workbookId;
  }

  const listeners = new Set<() => void>();
  const initialSessionId = normalizeSessionId(opts.initialSessionId ?? null);
  let sessionId: SessionId = initialSessionId && isSessionId(initialSessionId)
    ? initialSessionId
    : crypto.randomUUID();
  let sessionTitle = "";
  let sessionCreatedAt = new Date().toISOString();
  let explicitTitle = false;
  let canWriteLatestSessionPointer = !opts.autoRestoreLatest;
  /** Stored revision this taskpane last loaded or saved; saves only go over it. */
  let baseRevision: string | undefined;
  /** Saves run one at a time so each one's base is the previous one's result. */
  let saveQueue: Promise<void> = Promise.resolve();

  agent.sessionId = sessionId;

  function emitChange(): void {
    for (const listener of listeners) {
      listener();
    }
  }

  async function updateWorkbookAssociation(savedSessionId: string): Promise<void> {
    try {
      const workbookId = await resolveWorkbookId();
      if (!workbookId) return;

      await spreadsheetHost.sessionStorage.linkSessionToWorkbook(settings, savedSessionId, workbookId);
      if (canWriteLatestSessionPointer) {
        await spreadsheetHost.sessionStorage.setLatestSessionForWorkbook(settings, workbookId, savedSessionId);
      }
    } catch (err) {
      console.warn("[pi] Workbook/session association update failed:", err);
    }
  }

  function saveSession(optsForSave?: { force?: boolean }): Promise<boolean> {
    const run = saveQueue.then(() => saveSessionNow(optsForSave));
    saveQueue = run.then(() => undefined);
    return run;
  }

  async function saveSessionNow(optsForSave?: { force?: boolean }): Promise<boolean> {
    const messages = agent.state.messages;
    // A session exists from the user's first message, so a reload during the
    // first turn keeps it. Blank tabs are saved only on request (close, restore).
    // Compaction moves the user's messages into the summary and archive, which still count.
    const hasConversation = messages.some((m) =>
      isChatUserMessage(m) || m.role === "compactionSummary" || m.role === "archivedMessages");
    if (!optsForSave?.force && !hasConversation) return false;

    try {
      const now = new Date().toISOString();

      if (!sessionTitle && messages.length > 0) {
        const firstUser = messages.find((m) => isChatUserMessage(m));
        if (firstUser) {
          const text = extractTextFromContent(firstUser.content);
          sessionTitle = text.slice(0, 80) || "Untitled";
        }
      }

      let preview = "";
      for (const message of messages) {
        let text = "";

        if (message.role === "compactionSummary") {
          text = message.summary;
        } else if (message.role === "user" || message.role === "assistant") {
          text = extractTextFromContent(message.content);
        } else {
          continue;
        }

        preview += text + "\n";
        if (preview.length > 2048) {
          preview = preview.slice(0, 2048);
          break;
        }
      }

      let inputTokens = 0;
      let outputTokens = 0;
      let cacheReadTokens = 0;
      let cacheWriteTokens = 0;
      let totalTokens = 0;

      let costInput = 0;
      let costOutput = 0;
      let costCacheRead = 0;
      let costCacheWrite = 0;
      let costTotal = 0;

      for (const message of messages) {
        if (message.role !== "assistant") continue;
        const usage = message.usage;

        inputTokens += usage.input;
        outputTokens += usage.output;
        cacheReadTokens += usage.cacheRead;
        cacheWriteTokens += usage.cacheWrite;
        totalTokens += usage.totalTokens;

        costInput += usage.cost.input;
        costOutput += usage.cost.output;
        costCacheRead += usage.cost.cacheRead;
        costCacheWrite += usage.cost.cacheWrite;
        costTotal += usage.cost.total;
      }

      const buildMetadata = (id: string): SessionMetadata => ({
        id,
        title: sessionTitle,
        createdAt: sessionCreatedAt,
        lastModified: now,
        messageCount: countVisibleSessionMessages(messages),
        titleChosen: explicitTitle,
        usage: {
          input: inputTokens,
          output: outputTokens,
          cacheRead: cacheReadTokens,
          cacheWrite: cacheWriteTokens,
          totalTokens,
          cost: {
            input: costInput,
            output: costOutput,
            cacheRead: costCacheRead,
            cacheWrite: costCacheWrite,
            total: costTotal,
          },
        },
        thinkingLevel: agent.state.thinkingLevel || "off",
        preview,
      });

      let savedSessionId = sessionId;
      let nextRevision = crypto.randomUUID();
      const saved = await sessions.saveSession(
        savedSessionId,
        agent.state,
        buildMetadata(savedSessionId),
        sessionTitle,
        { base: baseRevision, next: nextRevision },
      );

      if (!saved) {
        // Another taskpane saved this session since we loaded it. Keep both
        // conversations: ours becomes a new session instead of replacing theirs.
        const fromSessionId = savedSessionId;
        savedSessionId = crypto.randomUUID();
        sessionId = savedSessionId;
        sessionCreatedAt = now;
        agent.sessionId = savedSessionId;
        nextRevision = crypto.randomUUID();
        await sessions.saveSession(
          savedSessionId,
          agent.state,
          { ...buildMetadata(savedSessionId), createdAt: now },
          sessionTitle,
          { base: undefined, next: nextRevision },
        );
        opts.onSavedAsCopy?.({ fromSessionId, toSessionId: savedSessionId, title: sessionTitle });
      }
      baseRevision = nextRevision;

      await updateWorkbookAssociation(savedSessionId);
      emitChange();
      return true;
    } catch (err) {
      console.warn("[pi] Session save failed:", err);
      opts.onSaveError?.(getErrorMessage(err));
      return false;
    }
  }

  function startNewSession(): void {
    sessionId = crypto.randomUUID();
    baseRevision = undefined;
    sessionTitle = "";
    sessionCreatedAt = new Date().toISOString();
    explicitTitle = false;
    canWriteLatestSessionPointer = true;
    agent.sessionId = sessionId;
    emitChange();
  }

  async function renameSession(title: string): Promise<void> {
    sessionTitle = title.trim();
    explicitTitle = sessionTitle.length > 0;
    emitChange();
    // A chosen title is a decision the user expects to keep, even before Pi answers.
    await saveSession({ force: true });
  }

  async function applyLoadedSession(sessionData: SessionData): Promise<void> {
    canWriteLatestSessionPointer = true;

    if (isSessionId(sessionData.id)) {
      sessionId = sessionData.id;
    } else {
      sessionId = crypto.randomUUID();
    }

    sessionTitle = sessionData.title || "";
    explicitTitle = sessionData.titleChosen === true && sessionTitle.trim().length > 0;
    sessionCreatedAt = sessionData.createdAt;
    baseRevision = sessionData.revision;

    agent.sessionId = sessionId;
    agent.state.messages = sessionData.messages;

    if (sessionData.model) {
      agent.state.model = opts.resolveModel(sessionData.model) ?? sessionData.model;
    }
    if (sessionData.thinkingLevel) {
      // The saved model may have been swapped for one with fewer thinking levels.
      agent.state.thinkingLevel = clampThinkingLevel(agent.state.model, sessionData.thinkingLevel);
    }

    await updateWorkbookAssociation(sessionId);
    emitChange();
  }

  async function restoreLatestSession(): Promise<boolean> {
    try {
      const workbookId = await resolveWorkbookId();
      const workbookLatest = workbookId
        ? await spreadsheetHost.sessionStorage.getLatestSessionForWorkbook(settings, workbookId)
        : null;
      const globalLatest = workbookId ? null : await sessions.getLatestSessionId();
      canWriteLatestSessionPointer = true;

      const candidates = getRestoreCandidateSessionIds({
        workbookId,
        workbookLatestSessionId: workbookLatest,
        globalLatestSessionId: globalLatest,
      });

      for (const candidateId of candidates) {
        const sessionData = await sessions.loadSession(candidateId);
        if (!sessionData || sessionData.messages.length === 0) continue;

        await applyLoadedSession(sessionData);
        console.log(`[pi] Restored session: ${sessionData.title || candidateId}`);
        return true;
      }
    } catch (err) {
      console.warn("[pi] Session restore failed:", err);
    }

    return false;
  }

  const unsubscribeAgent = agent.subscribe((event) => {
    if (event.type !== "message_end") return;
    void saveSession();
  });

  if (opts.autoRestoreLatest) {
    await restoreLatestSession();
  }

  return {
    getSessionId: () => sessionId,
    getSessionTitle: () => sessionTitle,
    getSessionCreatedAt: () => sessionCreatedAt,
    hasExplicitTitle: () => explicitTitle,
    startNewSession,
    renameSession,
    applyLoadedSession,
    restoreLatestSession,
    saveSession,
    subscribe(listener: () => void): () => void {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    dispose(): void {
      unsubscribeAgent();
      listeners.clear();
    },
  };
}
