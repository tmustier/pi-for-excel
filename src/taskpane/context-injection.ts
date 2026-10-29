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
  getAgent: () => Pick<Agent, "state"> | null | undefined;
  runtimeId: string;
  getWorkbookId?: () => Promise<string | null>;
  getBlueprint?: (workbookId: string | null) => Promise<string>;
  getFilesSummary?: (workbookId: string | null) => Promise<FilesSummary | null>;
  readSelection?: () => Promise<SelectionContext>;
}

export interface AutoContext {
  prepareRequest: PrepareRequest;
  transformContext: (messages: AgentMessage[], signal?: AbortSignal) => Promise<AgentMessage[]>;
  noteOwnStructureChange: () => void;
}

const WORKBOOK_ID_TIMEOUT_MS = 1200;
const BLUEPRINT_TIMEOUT_MS = 2500;
const FILES_TIMEOUT_MS = 1200;
const SELECTION_TIMEOUT_MS = 2000;

const TIMED_OUT = Symbol("timed out");

async function attempt<T>(what: string, run: () => Promise<T>, timeoutMs: number): Promise<T | null> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  try {
    const promise = run();
    const timeout = new Promise<typeof TIMED_OUT>((resolve) => {
      timeoutId = setTimeout(() => resolve(TIMED_OUT), timeoutMs);
    });
    const result = await Promise.race([promise, timeout]);
    if (result !== TIMED_OUT) return result;
    console.warn(`[auto-context] ${what} timed out after ${timeoutMs} ms`);
  } catch (error) {
    console.warn(`[auto-context] ${what} failed:`, getErrorMessage(error));
  } finally {
    if (timeoutId !== undefined) clearTimeout(timeoutId);
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

function note(kind: AutoContextKind, text: string, extra: Pick<AutoContextMessage, "key" | "workbookId"> = {}): AutoContextMessage {
  return { role: "autoContext", kind, text, timestamp: Date.now(), ...extra };
}

export function createAutoContext(sources: AutoContextSources): AutoContext {
  const loadWorkbookId = sources.getWorkbookId ?? (async () => (await getWorkbookContext()).workbookId);
  const loadBlueprint = sources.getBlueprint ?? getBlueprint;
  const loadFilesSummary = sources.getFilesSummary ?? ((workbookId) => getFilesWorkspace().getContextSummary(workbookId));
  const loadSelection = sources.readSelection ?? readSelectionContext;
  let changeCursor = sources.changeTracker.cursor();
  let ownStructureChangePending = false;

  async function collectNotes(messages: readonly AgentMessage[]): Promise<AutoContextMessage[]> {
    const notes: AutoContextMessage[] = [];
    const compacted = messages.some((message) => message.role === "compactionSummary");
    const workbookId = await attempt("workbook identity", loadWorkbookId, WORKBOOK_ID_TIMEOUT_MS);
    const midTurn = messages.length > 0 && !isPrompt(messages[messages.length - 1]);
    const deferStructure = midTurn && ownStructureChangePending;
    if (!midTurn) ownStructureChangePending = false;

    const blueprint = await attempt("workbook blueprint", () => loadBlueprint(workbookId), BLUEPRINT_TIMEOUT_MS);
    if (blueprint && blueprint.trim().length > 0) {
      const key = `${workbookId ?? "unknown"}:${hashText(blueprint)}`;
      const last = latestNote(messages, "workbook");
      let reason: string | null = null;
      if (!last) reason = compacted ? "re-sent after the conversation was compacted" : "initial";
      else if (last.workbookId !== workbookId) reason = "workbook switched";
      else if (last.key !== key && !deferStructure) reason = "structure changed";
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

    const changes = await sources.changeTracker.describeSince(changeCursor, sources.runtimeId);
    changeCursor = changes.cursor;
    if (changes.text) notes.push(note("changes", changes.text));

    return notes;
  }

  const prepareRequest: PrepareRequest = async (request) => {
    const messages = request.context.messages;
    const notes = await collectNotes(messages);
    if (notes.length === 0) return undefined;

    let insertAt = messages.length;
    while (insertAt > 0 && isPrompt(messages[insertAt - 1])) insertAt -= 1;
    const anchor = messages[insertAt];

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
    const midTurn = messages.length > 0 && !isPrompt(messages[messages.length - 1]);
    const reminder = midTurn
      ? "(Background from Excel, not a message from the user and not a new request. Continue from the tool results above; do not repeat steps that already succeeded.)"
      : "(Background from Excel, not a message from the user.)";
    return [
      ...messages,
      { role: "user", content: [{ type: "text", text: `${AUTO_CONTEXT_PREFIX} ${text}\n\n${reminder}` }], timestamp: Date.now() },
    ];
  };

  return {
    prepareRequest,
    transformContext,
    noteOwnStructureChange: () => { ownStructureChangePending = true; },
  };
}
