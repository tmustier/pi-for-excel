/**
 * Auto-context: what Pi tells the agent about the workbook without being asked.
 *
 * Two layers, chosen for prompt caching (see docs/context-management-policy.md):
 *
 * 1. Notes stored in the transcript (`prepareRequest`). Snapshots that stay
 *    true until something changes (the workbook blueprint, the Files summary)
 *    and events (cells edited by someone other than Pi). Before each model
 *    request, a note is added only when its subject differs from the last
 *    note of that kind in the transcript. Stored notes become part of the
 *    cached prefix, remain visible for the rest of the turn and in later
 *    turns, survive reloads, and are re-sent only when compaction removed
 *    them. They are inserted before the user's new prompt, or after the latest
 *    tool results mid-turn.
 *
 * 2. The current selection (`transformContext`). State that can change at any
 *    moment, so it is re-read for every request and appended as the last
 *    message. It is never stored, is bounded in size, and sits after the
 *    cached prefix.
 */

import type { Agent, AgentMessage, PrepareRequest } from "@earendil-works/pi-agent-core";

import { getBlueprint } from "../context/blueprint.js";
import type { ChangeTracker } from "../context/change-tracker.js";
import { readSelectionContext, type SelectionContext } from "../context/selection.js";
import { getFilesWorkspace } from "../files/workspace.js";
import { AUTO_CONTEXT_PREFIX, type AutoContextKind, type AutoContextMessage } from "../messages/auto-context.js";
import { getErrorMessage } from "../utils/errors.js";
import { getWorkbookContext } from "../workbook/context.js";

interface FilesSummary {
  summary: string;
  relevantSignature: string;
  hasRelevantFiles: boolean;
}

export interface AutoContextSources {
  changeTracker: Pick<ChangeTracker, "cursor" | "describeSince">;
  /** The runtime's agent: stored notes are added to its transcript. */
  getAgent: () => Pick<Agent, "state"> | null | undefined;
  getWorkbookId?: () => Promise<string | null>;
  getBlueprint?: (workbookId: string | null) => Promise<string>;
  getFilesSummary?: (workbookId: string | null) => Promise<FilesSummary | null>;
  readSelection?: () => Promise<SelectionContext>;
}

export interface AutoContext {
  prepareRequest: PrepareRequest;
  transformContext: (messages: AgentMessage[], signal?: AbortSignal) => Promise<AgentMessage[]>;
}

const WORKBOOK_ID_TIMEOUT_MS = 1200;
const BLUEPRINT_TIMEOUT_MS = 2500;
const FILES_TIMEOUT_MS = 1200;
const SELECTION_TIMEOUT_MS = 2000;

const TIMED_OUT = Symbol("timed out");

async function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T | typeof TIMED_OUT> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  const timeout = new Promise<typeof TIMED_OUT>((resolve) => {
    timeoutId = setTimeout(() => resolve(TIMED_OUT), timeoutMs);
  });
  try {
    return await Promise.race([promise, timeout]);
  } finally {
    if (timeoutId !== undefined) clearTimeout(timeoutId);
  }
}

/** Run one auto-context source; a failure or timeout is logged and yields null. */
async function attempt<T>(what: string, run: () => Promise<T>, timeoutMs: number): Promise<T | null> {
  try {
    const result = await withTimeout(run(), timeoutMs);
    if (result !== TIMED_OUT) return result;
    console.warn(`[auto-context] ${what} timed out after ${timeoutMs} ms`);
  } catch (error) {
    console.warn(`[auto-context] ${what} failed:`, getErrorMessage(error));
  }
  return null;
}

function hashText(value: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193) >>> 0;
  }
  return hash.toString(16).padStart(8, "0");
}

function latestNote(messages: readonly AgentMessage[], kind: AutoContextKind): AutoContextMessage | undefined {
  for (let i = messages.length - 1; i >= 0; i -= 1) {
    const message = messages[i];
    if (message?.role === "autoContext" && message.kind === kind) return message;
  }
  return undefined;
}

function isPrompt(message: AgentMessage | undefined): boolean {
  return message?.role === "user" || message?.role === "user-with-attachments";
}

function note(kind: AutoContextKind, text: string, extra: Partial<AutoContextMessage> = {}): AutoContextMessage {
  return { role: "autoContext", kind, text, timestamp: Date.now(), ...extra };
}

async function defaultWorkbookId(): Promise<string | null> {
  return (await getWorkbookContext()).workbookId;
}

async function defaultFilesSummary(workbookId: string | null): Promise<FilesSummary | null> {
  return getFilesWorkspace().getContextSummary(workbookId);
}

export function createAutoContext(sources: AutoContextSources): AutoContext {
  const loadWorkbookId = sources.getWorkbookId ?? defaultWorkbookId;
  const loadBlueprint = sources.getBlueprint ?? getBlueprint;
  const loadFilesSummary = sources.getFilesSummary ?? defaultFilesSummary;
  const loadSelection = sources.readSelection ?? readSelectionContext;
  // Each runtime (chat tab) reads the shared edit log from its own cursor.
  let changeCursor = sources.changeTracker.cursor();

  async function collectNotes(messages: readonly AgentMessage[]): Promise<AutoContextMessage[]> {
    const notes: AutoContextMessage[] = [];
    const compacted = messages.some((message) => message.role === "compactionSummary");
    const workbookId = await attempt("workbook identity", loadWorkbookId, WORKBOOK_ID_TIMEOUT_MS);

    const blueprint = await attempt("workbook blueprint", () => loadBlueprint(workbookId), BLUEPRINT_TIMEOUT_MS);
    if (blueprint && blueprint.trim().length > 0) {
      const key = `${workbookId ?? "unknown"}:${hashText(blueprint)}`;
      const last = latestNote(messages, "workbook");
      let reason: string | null = null;
      if (!last) reason = compacted ? "re-sent after the conversation was compacted" : "initial";
      else if (last.workbookId !== workbookId) reason = "workbook switched";
      else if (last.key !== key) reason = "structure changed";
      if (reason) {
        notes.push(note("workbook", [
          `Workbook structure (${reason}). Sizes are as of this note; a new note follows when sheets, tables, names, rows or columns change.`,
          blueprint,
        ].join("\n\n"), { key, workbookId }));
      }
    }

    const files = await attempt("Files workspace summary", () => loadFilesSummary(workbookId), FILES_TIMEOUT_MS);
    if (files) {
      const last = latestNote(messages, "files");
      let reason: string | null = null;
      if (!last && files.hasRelevantFiles) reason = compacted ? "re-sent after the conversation was compacted" : "initial";
      else if (last && last.key !== files.relevantSignature) reason = "files changed";
      if (reason) {
        notes.push(note("files", `Files workspace (${reason}):\n\n${files.summary}`, { key: files.relevantSignature }));
      }
    }

    const changes = sources.changeTracker.describeSince(changeCursor);
    changeCursor = changes.cursor;
    if (changes.text) notes.push(note("changes", changes.text));

    return notes;
  }

  const prepareRequest: PrepareRequest = async (request) => {
    const messages = request.context.messages;
    const notes = await collectNotes(messages);
    if (notes.length === 0) return undefined;

    // Before the user's new prompt(s); after tool results mid-turn.
    let insertAt = messages.length;
    while (insertAt > 0 && isPrompt(messages[insertAt - 1])) insertAt -= 1;
    const anchor = messages[insertAt];

    // Store the notes in the runtime's transcript at the same place, so they
    // persist with the session and later runs see them.
    const agent = sources.getAgent();
    if (agent) {
      const transcript = agent.state.messages;
      const index = anchor ? transcript.indexOf(anchor) : -1;
      agent.state.messages = index >= 0
        ? [...transcript.slice(0, index), ...notes, ...transcript.slice(index)]
        : [...transcript, ...notes];
    }

    return {
      context: {
        ...request.context,
        messages: [...messages.slice(0, insertAt), ...notes, ...messages.slice(insertAt)],
      },
    };
  };

  const transformContext = async (messages: AgentMessage[]): Promise<AgentMessage[]> => {
    const selection = await attempt("selection", loadSelection, SELECTION_TIMEOUT_MS);
    const text = selection?.text
      ?? "Current selection: unavailable (Excel did not answer in time).";
    return [
      ...messages,
      { role: "user", content: [{ type: "text", text: `${AUTO_CONTEXT_PREFIX} ${text}` }], timestamp: Date.now() },
    ];
  };

  return { prepareRequest, transformContext };
}
