function isCommandsBuiltinsExportPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/**
 * Builtin export/compaction commands.
 */

import {
  getCurrentSystemMessage,
  normalizeContext,
  type Api,
  type Model,
  type StopReason,
  type Usage,
} from "@earendil-works/pi-ai";
import type { Agent, AgentMessage } from "@earendil-works/pi-agent-core";

import type { SlashCommand } from "../types.js";
import type { ActiveAgentProvider } from "./model.js";
import { copyText } from "../../ui/copy-text.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import { createCompactionSummaryMessage } from "../../messages/compaction.js";
import { compactToolResult } from "../../messages/tool-result-shaping.js";
import {
  createArchivedMessagesMessage,
  splitArchivedMessages,
} from "../../messages/archived-history.js";
import { getErrorMessage } from "../../utils/errors.js";
import { extractTextBlocks, summarizeContentForTranscript } from "../../utils/content.js";
import type { PiSidebar } from "../../ui/pi-sidebar.js";
import { getWorkbookChangeAuditLog } from "../../audit/workbook-change-audit.js";
import { getWorkbookContext } from "../../workbook/context.js";
import { effectiveKeepRecentTokens, effectiveReserveTokens } from "../../compaction/defaults.js";
import {
  buildCompactionMemoryFocusInstruction,
  collectCompactionMemoryCues,
  mergeCompactionAdditionalFocus,
} from "../../compaction/memory-nudge.js";

type TranscriptEntry = {
  role: AgentMessage["role"];
  text: string;
  usage?: Usage;
  stopReason?: StopReason;
};

function isApiModel(model: unknown): model is Model<Api> {
  if (!isCommandsBuiltinsExportPayloadShape(model)) return false;

  return (
    typeof model.id === "string" &&
    typeof model.name === "string" &&
    typeof model.provider === "string" &&
    typeof model.api === "string"
  );
}

function hasContent(message: AgentMessage): message is AgentMessage & { content: unknown } {
  return isCommandsBuiltinsExportPayloadShape(message) && "content" in message;
}

function messageToTranscriptText(message: AgentMessage): string {
  if (message.role === "archivedMessages") {
    return `[archived history: ${message.archivedChatMessageCount} chat messages]`;
  }

  if (message.role === "compactionSummary") return message.summary;
  if (message.role === "autoContext") return message.text;
  if (hasContent(message)) return summarizeContentForTranscript(message.content);
  return "";
}

function countChatMessages(messages: AgentMessage[]): number {
  let count = 0;
  for (const m of messages) {
    const role = m.role;
    if (role === "user" || role === "assistant" || role === "user-with-attachments") {
      count += 1;
    }
  }
  return count;
}

type ExportDestination = "clipboard" | "download";

function parseExportDestination(raw: string, fallback: ExportDestination): ExportDestination {
  const normalized = raw.trim().toLowerCase();
  if (normalized === "clipboard") return "clipboard";
  if (normalized === "file" || normalized === "download") return "download";
  return fallback;
}

function triggerJsonDownload(fileName: string, content: string): void {
  const blob = new Blob([content], { type: "application/json" });
  const url = URL.createObjectURL(blob);

  // Try window.open() first for Office WebView (WKWebView) compatibility.
  // Fall back to <a download> if popup is blocked (lost user activation).
  const opened = window.open(url, "_blank");
  if (!opened) {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.rel = "noopener";
    anchor.hidden = true;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }

  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

async function exportWorkbookAuditLog(rawArgs: string): Promise<void> {
  const destination = parseExportDestination(rawArgs, "download");

  // Only this workbook's changes: the log holds addresses and values from every workbook.
  // With no workbook identity, only entries that were recorded without one are included.
  const { workbookId } = await getWorkbookContext();
  const entries = (await getWorkbookChangeAuditLog().list(500))
    .filter((entry) => (entry.workbookId ?? null) === workbookId);
  const payload = {
    exported: new Date().toISOString(),
    workbookId,
    count: entries.length,
    entries,
  };

  const json = JSON.stringify(payload, null, 2);

  if (destination === "clipboard") {
    if (await copyText(json) === "manual") return;
    showToast(t("export.toast.audit_copied", {
      count: String(entries.length),
      size: (json.length / 1024).toFixed(0),
    }));
    return;
  }

  triggerJsonDownload(`pi-audit-log-${new Date().toISOString().slice(0, 10)}.json`, json);
  showToast(t("export.toast.audit_downloaded", { count: String(entries.length) }));
}

// =============================================================================
// Compaction helpers
// =============================================================================

// Mirrors pi-coding-agent defaults (see docs/compaction.md in pi-coding-agent).

const SUMMARIZATION_SYSTEM_PROMPT =
  "You are a context summarization assistant. Your task is to read a conversation between a user and an AI assistant, then produce a structured summary following the exact format specified.\n\n" +
  "Do NOT continue the conversation. Do NOT respond to any questions in the conversation. ONLY output the structured summary.";

const SUMMARIZATION_PROMPT = `The messages above are a conversation to summarize. Create a structured context checkpoint summary that another LLM will use to continue the work.

Use this EXACT format:

## Goal
[What is the user trying to accomplish? Can be multiple items if the session covers different tasks.]

## Constraints & Preferences
- [Any constraints, preferences, or requirements mentioned by user]
- [Or "(none)" if none were mentioned]

## Progress
### Done
- [x] [Completed tasks/changes]

### In Progress
- [ ] [Current work]

### Blocked
- [Issues preventing progress, if any]

## Key Decisions
- **[Decision]**: [Brief rationale]

## Next Steps
1. [Ordered list of what should happen next]

## Critical Context
- [Any data, examples, or references needed to continue]
- [Or "(none)" if not applicable]

Keep each section concise. Preserve exact identifiers (sheet names, cell addresses, tool names, error messages).`;

/** Appended to both prompts: requests are labelled with whether they were answered. */
const ANSWERED_REQUESTS_RULE = `Each user request is labelled "answered" or "not answered yet". An answered request is finished: put it under Done with its outcome, and never under In Progress or Next Steps; do not suggest repeating its tool calls. A reply given after the summarized messages is shown right after them. Only requests not answered yet can be In Progress.`;

const UPDATE_SUMMARIZATION_PROMPT = `The messages above are NEW conversation messages to incorporate into the existing summary provided in <previous-summary> tags.

Update the existing structured summary with new information. RULES:
- PRESERVE all existing information from the previous summary
- ADD new progress, decisions, and context from the new messages
- UPDATE the Progress section: move items from "In Progress" to "Done" when completed
- UPDATE "Next Steps" based on what was accomplished
- PRESERVE exact identifiers (sheet names, cell addresses, tool names, error messages)
- If something is no longer relevant, you may remove it

Use this EXACT format:

## Goal
[Preserve existing goals, add new ones if the task expanded]

## Constraints & Preferences
- [Preserve existing, add new ones discovered]

## Progress
### Done
- [x] [Include previously done items AND newly completed items]

### In Progress
- [ ] [Current work - update based on progress]

### Blocked
- [Current blockers - remove if resolved]

## Key Decisions
- **[Decision]**: [Brief rationale] (preserve all previous, add new)

## Next Steps
1. [Update based on current state]

## Critical Context
- [Preserve important context, add new if needed]

Keep each section concise. Preserve exact identifiers (sheet names, cell addresses, tool names, error messages).`;

type SerializeLimits = {
  maxUserChars: number;
  maxAssistantChars: number;
  maxToolResultChars: number;
};

function truncateMiddle(text: string, maxChars: number): string {
  if (maxChars <= 0) return "";
  if (text.length <= maxChars) return text;

  const marker = "\n…[truncated]…\n";
  const keep = Math.max(0, maxChars - marker.length);
  const head = Math.floor(keep / 2);
  const tail = keep - head;

  return text.slice(0, head) + marker + text.slice(text.length - tail);
}

function estimateTokens(message: AgentMessage): number {
  // Conservative heuristic from pi-coding-agent: tokens ≈ chars / 4
  const charsPerToken = 4;
  let chars = 0;

  if (message.role === "artifact") {
    // UI-only, not part of LLM context (defaultConvertToLlm filters it out).
    return 0;
  }

  if (message.role === "compactionSummary") {
    return Math.ceil(message.summary.length / charsPerToken);
  }

  if (message.role === "autoContext") {
    return Math.ceil(message.text.length / charsPerToken);
  }

  if (message.role === "user" || message.role === "user-with-attachments") {
    const content = message.content;
    if (typeof content === "string") {
      chars += content.length;
    } else if (Array.isArray(content)) {
      for (const block of content) {
        if (block.type === "text") chars += block.text.length;
        if (block.type === "image") chars += 4800; // ~1200 tokens
      }
    }
    return Math.ceil(chars / charsPerToken);
  }

  if (message.role === "assistant") {
    for (const block of message.content) {
      if (block.type === "text") chars += block.text.length;
      else if (block.type === "thinking") chars += block.thinking.length;
      else if (block.type === "toolCall") {
        chars += block.name.length;
        try {
          chars += JSON.stringify(block.arguments).length;
        } catch {
          // ignore
        }
      }
    }
    return Math.ceil(chars / charsPerToken);
  }

  if (message.role === "toolResult") {
    for (const block of message.content) {
      if (block.type === "text") chars += block.text.length;
      if (block.type === "image") chars += 4800;
    }
    return Math.ceil(chars / charsPerToken);
  }

  // Unknown custom message types: ignore.
  return 0;
}

function getPreviousCompaction(messages: AgentMessage[]): { boundaryStart: number; previousSummary?: string } {
  for (let i = messages.length - 1; i >= 0; i--) {
    const m = messages[i];
    if (m?.role === "compactionSummary") {
      return { boundaryStart: i + 1, previousSummary: m.summary };
    }
  }
  return { boundaryStart: 0 };
}

function findCutIndex(
  messages: AgentMessage[],
  boundaryStart: number,
  keepRecentTokens: number,
  estimate: (message: AgentMessage) => number = estimateTokens,
): number {
  let accumulated = 0;

  for (let i = messages.length - 1; i >= boundaryStart; i--) {
    const message = messages[i];
    if (!message) {
      continue;
    }

    accumulated += estimate(message);
    if (accumulated >= keepRecentTokens) {
      let cut = i;
      // Never start kept context with a tool result.
      while (cut > boundaryStart && messages[cut]?.role === "toolResult") {
        cut -= 1;
      }
      return cut;
    }
  }

  return boundaryStart;
}

/** Kept results above this size are compacted to a preview when their request was answered. */
const KEPT_TOOL_RESULT_MAX_CHARS = 8_000;

function serializeConversation(
  messages: AgentMessage[],
  limits: SerializeLimits,
  answered: ReadonlyMap<AgentMessage, boolean> = new Map(),
): string {
  const parts: string[] = [];

  for (const msg of messages) {
    if (msg.role === "artifact") continue;

    if (msg.role === "autoContext") {
      // Workbook and Files notes are re-sent after compaction; user edits are not.
      if (msg.kind === "changes") parts.push(`[Workbook edits by the user]: ${truncateMiddle(msg.text, limits.maxUserChars)}`);
      continue;
    }

    if (msg.role === "user" || msg.role === "user-with-attachments") {
      const raw = typeof msg.content === "string" ? msg.content : extractTextBlocks(msg.content);
      const text = truncateMiddle(raw, limits.maxUserChars);
      const label = answered.has(msg) ? (answered.get(msg) ? "User request, answered" : "User request, not answered yet") : "User";
      if (text.trim().length > 0) parts.push(`[${label}]: ${text}`);
      continue;
    }

    if (msg.role === "assistant") {
      const textParts: string[] = [];
      const thinkingParts: string[] = [];
      const toolCalls: string[] = [];

      for (const block of msg.content) {
        if (block.type === "text") {
          textParts.push(block.text);
        } else if (block.type === "thinking") {
          thinkingParts.push(block.thinking);
        } else if (block.type === "toolCall") {
          let args = "";
          try {
            args = JSON.stringify(block.arguments);
          } catch {
            args = "{}";
          }
          toolCalls.push(`${block.name}(${args})`);
        }
      }

      if (thinkingParts.length > 0) {
        const t = truncateMiddle(thinkingParts.join("\n"), limits.maxAssistantChars);
        parts.push(`[Assistant thinking]: ${t}`);
      }

      if (textParts.length > 0) {
        const t = truncateMiddle(textParts.join("\n"), limits.maxAssistantChars);
        parts.push(`[Assistant]: ${t}`);
      }

      if (toolCalls.length > 0) {
        const t = truncateMiddle(toolCalls.join("; "), limits.maxAssistantChars);
        parts.push(`[Assistant tool calls]: ${t}`);
      }

      continue;
    }

    if (msg.role === "toolResult") {
      const raw = extractTextBlocks(msg.content);
      const text = truncateMiddle(raw, limits.maxToolResultChars);
      const label = `${msg.toolName}${msg.isError ? " (error)" : ""}`;
      if (text.trim().length > 0) {
        parts.push(`[Tool result ${label}]: ${text}`);
      }
      continue;
    }

    // Ignore other message types.
  }

  return parts.join("\n\n");
}

function buildSummarizationPrompt(args: {
  conversationText: string;
  previousSummary?: string;
  customInstructions?: string;
}): string {
  const base = args.previousSummary ? UPDATE_SUMMARIZATION_PROMPT : SUMMARIZATION_PROMPT;
  const withRule = `${base}\n\n${ANSWERED_REQUESTS_RULE}`;
  const withFocus = args.customInstructions
    ? `${withRule}\n\nAdditional focus: ${args.customInstructions}`
    : withRule;

  let prompt = `<conversation>\n${args.conversationText}\n</conversation>\n\n`;

  if (args.previousSummary) {
    prompt += `<previous-summary>\n${args.previousSummary}\n</previous-summary>\n\n`;
  }

  prompt += withFocus;
  return prompt;
}

function isPromptTooLongError(err: unknown): boolean {
  const msg = getErrorMessage(err).toLowerCase();
  return (
    msg.includes("prompt is too long") ||
    msg.includes("context_length_exceeded") ||
    (msg.includes("maximum") && msg.includes("tokens"))
  );
}

export function createExportCommands(getActiveAgent: ActiveAgentProvider): SlashCommand[] {
  return [
    {
      name: "export",
      description: t("command.export.json"),
      source: "builtin",
      execute: async (args: string) => {
        const parts = args.trim().split(/\s+/u).filter((part) => part.length > 0);
        const mode = parts[0]?.toLowerCase();

        if (mode === "audit" || mode === "audit-log") {
          try {
            await exportWorkbookAuditLog(parts.slice(1).join(" "));
          } catch (error) {
            showToast(t("export.toast.audit_export_failed", { error: getErrorMessage(error) }));
          }
          return;
        }

        const agent = getActiveAgent();
        if (!agent) {
          showToast(t("export.toast.no_session"));
          return;
        }

        const msgs = agent.state.messages;
        if (msgs.length === 0) {
          showToast(t("export.toast.no_messages"));
          return;
        }

        const transcript: TranscriptEntry[] = msgs.map((m) => {
          const text = messageToTranscriptText(m);
          if (m.role === "assistant") {
            return {
              role: m.role,
              text,
              usage: m.usage,
              stopReason: m.stopReason,
            };
          }
          return { role: m.role, text };
        });

        const exportData = {
          exported: new Date().toISOString(),
          model: agent.state.model
            ? {
              id: agent.state.model.id,
              name: agent.state.model.name,
              provider: agent.state.model.provider,
            }
            : null,
          thinkingLevel: agent.state.thinkingLevel,
          messageCount: msgs.length,
          transcript,
          // Also include raw messages for full fidelity debugging
          raw: msgs,
        };

        const json = JSON.stringify(exportData, null, 2);
        const destination = parseExportDestination(args, "clipboard");

        if (destination === "clipboard") {
          try {
            if (await copyText(json) === "manual") return;
            showToast(t("export.toast.transcript_copied", {
              count: String(countChatMessages(msgs)),
              size: (json.length / 1024).toFixed(0),
            }));
          } catch (error) {
            showToast(t("export.toast.copy_failed", { error: getErrorMessage(error) }));
          }
          return;
        }

        triggerJsonDownload(`pi-session-${new Date().toISOString().slice(0, 10)}.json`, json);
        showToast(t("export.toast.transcript_downloaded", { count: String(countChatMessages(msgs)) }));
      },
    },
  ];
}

class CompactionStoppedError extends Error {
  constructor() {
    super("Compaction stopped");
  }
}

function untilStopped<T>(work: Promise<T>, signal: AbortSignal | undefined): Promise<T> {
  if (!signal) return work;
  if (signal.aborted) return Promise.reject(new CompactionStoppedError());
  return new Promise<T>((resolve, reject) => {
    const onAbort = () => reject(new CompactionStoppedError());
    signal.addEventListener("abort", onAbort, { once: true });
    work.then(
      (value) => {
        signal.removeEventListener("abort", onAbort);
        resolve(value);
      },
      (error: Error) => {
        signal.removeEventListener("abort", onAbort);
        reject(error);
      },
    );
  });
}

/**
 * Run compaction for a specific agent.
 *
 * Used by the `/compact` slash command (with the active agent) and by
 * auto-compaction / overflow recovery (with the agent that owns the run, which
 * may not be the active tab).
 */
export async function runCompactCommand(agent: Agent, args: string, signal?: AbortSignal): Promise<void> {
  const allMessages = agent.state.messages;
  const currentSystemMessage = getCurrentSystemMessage(allMessages);
  const conversationMessages = allMessages.filter((message) => message.role !== "system");
  const {
    archivedMessages: existingArchivedMessages,
    messagesWithoutArchived,
  } = splitArchivedMessages(conversationMessages);

  if (messagesWithoutArchived.length < 4) {
    showToast(t("export.toast.compact.few_messages"));
    return;
  }

  showToast(t("export.toast.compact.compacting"), 60000);

  const now = Date.now();
  const model = agent.state.model;
  if (!isApiModel(model)) {
    showToast(t("export.toast.compact.no_model"));
    return;
  }

  // IMPORTANT: use the agent's configured stream function + API key resolver.
  // Calling pi-ai's completeSimple() directly bypasses:
  // - our CORS proxy logic (streamFunction)
  // - our API key/OAuth resolution (agent.getApiKey)
  // and can crash in browser WebViews due to env key fallbacks using `process`.
  const apiKey = agent.getApiKey ? await agent.getApiKey(model.provider) : undefined;
  if (!apiKey) {
    showToast(t("export.toast.compact.no_api_key", { provider: model.provider }));
    return;
  }

  const contextWindow = model.contextWindow || 200000;

  // Pi uses reserveTokens to ensure we don't run out of room for the model's response.
  const reserveTokens = effectiveReserveTokens(contextWindow);
  const keepRecentTokens = effectiveKeepRecentTokens(contextWindow, reserveTokens);

  const maxTokens = Math.max(
    256,
    Math.min(model.maxTokens, Math.floor(0.8 * reserveTokens)),
  );

  const { boundaryStart, previousSummary } = getPreviousCompaction(messagesWithoutArchived);
  const userCompactionFocus = args.trim() || undefined;
  let memoryNudgeShown = false;

  const prompts = messagesWithoutArchived.flatMap((message, start) =>
    message.role === "user" || message.role === "user-with-attachments" ? [{ start, prompt: message }] : []);
  const turns = prompts.map(({ start, prompt }, n) => {
    const end = prompts[n + 1]?.start ?? messagesWithoutArchived.length;
    let reply: { index: number; text: string } | null = null;
    for (let i = end - 1; i > start; i -= 1) {
      const message = messagesWithoutArchived[i];
      if (message?.role !== "assistant") continue;
      const text = message.content.flatMap((block) => (block.type === "text" ? [block.text] : [])).join("\n").trim();
      if (message.stopReason === "stop" && text.length > 0) reply = { index: i, text };
      break;
    }
    return { start, end, prompt, reply };
  });
  const compactable = new Set<string>();
  for (const turn of turns) {
    if (!turn.reply) continue;
    for (let i = Math.max(turn.start, boundaryStart); i < turn.end; i += 1) {
      const message = messagesWithoutArchived[i];
      if (message?.role !== "toolResult") continue;
      const chars = message.content.reduce((total, block) => total + (block.type === "text" ? block.text.length : 4_800), 0);
      if (chars > KEPT_TOOL_RESULT_MAX_CHARS) compactable.add(message.toolCallId);
    }
  }

  const runOnce = async (limits: SerializeLimits, keepRecentOverride?: number): Promise<{
    summary: string;
    keptMessages: AgentMessage[];
    compactedToolCallIds: string[];
    messagesToArchive: AgentMessage[];
    summarizedCount: number;
    summarizedTokens: number;
  }> => {
    const keepRecent = keepRecentOverride ?? keepRecentTokens;
    // Large results of answered requests are kept as previews, so they no
    // longer fill the kept tail on their own.
    const estimateKept = (message: AgentMessage) => estimateTokens(
      message.role === "toolResult" && compactable.has(message.toolCallId) ? compactToolResult(message) : message,
    );
    // When the previews leave nothing to summarize, cut as before: compacting
    // the kept results still frees their space.
    const shrunkCut = findCutIndex(messagesWithoutArchived, boundaryStart, keepRecent, estimateKept);
    const cutIndex = shrunkCut > boundaryStart ? shrunkCut : findCutIndex(messagesWithoutArchived, boundaryStart, keepRecent);
    const messagesToSummarize = messagesWithoutArchived.slice(boundaryStart, cutIndex);
    const keptMessages = messagesWithoutArchived.slice(cutIndex);
    const compactedToolCallIds = keptMessages.flatMap((message) => (
      message.role === "toolResult" && compactable.has(message.toolCallId) ? [message.toolCallId] : []
    ));

    // Tell the summarizer which requests were answered, including by a reply
    // that is kept after the summary.
    const answered = new Map<AgentMessage, boolean>();
    const laterReplies: string[] = [];
    for (const turn of turns) {
      if (turn.start < boundaryStart || turn.start >= cutIndex) continue;
      answered.set(turn.prompt, turn.reply !== null);
      if (turn.reply && turn.reply.index >= cutIndex) laterReplies.push(turn.reply.text);
    }

    if (messagesToSummarize.length === 0) {
      throw new Error("Nothing to compact");
    }

    const memoryCues = collectCompactionMemoryCues(messagesToSummarize);
    if (memoryCues.cueCount > 0 && !memoryNudgeShown) {
      const cueLabel = memoryCues.cueCount === 1 ? "cue" : "cues";
      showToast(
        t("export.toast.compact.memory_nudge", { count: memoryCues.cueCount, cue: cueLabel }),
        12000,
      );
      memoryNudgeShown = true;
    }

    const conversationText = [
      serializeConversation(messagesToSummarize, limits, answered),
      ...laterReplies.map((text) => `[Assistant, final reply to the request above, given after these messages]: ${truncateMiddle(text, limits.maxAssistantChars)}`),
    ].join("\n\n");
    const memoryFocus = buildCompactionMemoryFocusInstruction(memoryCues);
    const customInstructions = mergeCompactionAdditionalFocus(
      userCompactionFocus,
      memoryFocus,
    );
    const promptText = buildSummarizationPrompt({
      conversationText,
      ...(previousSummary !== undefined ? { previousSummary } : {}),
      ...(customInstructions !== undefined ? { customInstructions } : {}),
    });

    const request = async () => {
      const stream = await agent.streamFunction(
        model,
        normalizeContext({
          systemPrompt: SUMMARIZATION_SYSTEM_PROMPT,
          messages: [
            {
              role: "user",
              content: [{ type: "text", text: promptText }],
              timestamp: Date.now(),
            },
          ],
        }),
        {
          apiKey,
          ...(agent.sessionId !== undefined ? { sessionId: agent.sessionId } : {}),
          maxTokens,
          ...(signal ? { signal } : {}),
          // Match pi-coding-agent: don't force temperature when using reasoning,
          // since Anthropic requires temperature=1 when thinking is enabled.
          reasoning: "high",
        },
      );

      return stream.result();
    };
    // Stop ends compaction at once, even if the provider stream does not
    // honor the signal.
    const result = await untilStopped(request(), signal);

    if (result.stopReason === "aborted" || signal?.aborted) {
      throw new CompactionStoppedError();
    }

    if (result.stopReason === "error") {
      throw new Error(result.errorMessage || t("export.toast.compact.failed_error"));
    }

    const summary = extractTextBlocks(result.content).trim() || t("export.toast.compact.summary_unavailable");

    return {
      summary,
      keptMessages,
      compactedToolCallIds,
      messagesToArchive: messagesToSummarize,
      summarizedCount: countChatMessages(messagesToSummarize),
      summarizedTokens: messagesToSummarize.reduce(
        (total, message) => total + estimateTokens(message),
        0,
      ),
    };
  };

  const defaultLimits: SerializeLimits = {
    maxUserChars: 4000,
    maxAssistantChars: 8000,
    maxToolResultChars: 8000,
  };

  const aggressiveLimits: SerializeLimits = {
    maxUserChars: 1200,
    maxAssistantChars: 2500,
    maxToolResultChars: 2500,
  };

  try {
    let out: {
      summary: string;
      keptMessages: AgentMessage[];
      compactedToolCallIds: string[];
      messagesToArchive: AgentMessage[];
      summarizedCount: number;
      summarizedTokens: number;
    };

    try {
      out = await runOnce(defaultLimits);
    } catch (e) {
      if (e instanceof CompactionStoppedError || !isPromptTooLongError(e)) throw e;

      // Retry once with more aggressive truncation + keeping a larger recent tail.
      showToast(t("export.toast.compact.retrying"), 60000);

      const keepMoreRecent = Math.min(contextWindow, keepRecentTokens * 2);
      out = await runOnce(aggressiveLimits, keepMoreRecent);
    }

    const archived = createArchivedMessagesMessage({
      existingArchivedMessages,
      newlyArchivedMessages: out.messagesToArchive,
      timestamp: now,
    });

    const compacted = createCompactionSummaryMessage({
      summary: out.summary,
      tokensBefore: out.summarizedTokens,
      timestamp: now,
      compactedToolCallIds: out.compactedToolCallIds,
    });

    if (signal?.aborted) throw new CompactionStoppedError();

    agent.state.messages = [
      ...(currentSystemMessage ? [currentSystemMessage] : []),
      archived,
      compacted,
      ...out.keptMessages,
    ];

    const iface = document.querySelector<PiSidebar>("pi-sidebar");
    iface?.requestUpdate();

    showToast(t("export.toast.compact.summarized", {
      count: out.summarizedCount,
      plural: out.summarizedCount === 1 ? "" : "s",
    }));
  } catch (e) {
    if (e instanceof CompactionStoppedError) {
      showToast(t("export.toast.compact.stopped"));
      return;
    }
    const msg = getErrorMessage(e);
    if (msg === "Nothing to compact") {
      showToast(t("export.toast.compact.nothing"));
      return;
    }
    showToast(t("export.toast.compact.failed", { msg }));
  }
}

export function createCompactCommands(getActiveAgent: ActiveAgentProvider): SlashCommand[] {
  return [
    {
      name: "compact",
      description: t("command.export.summarize"),
      source: "builtin",
      execute: async (args: string) => {
        const agent = getActiveAgent();
        if (!agent) {
          showToast(t("export.toast.no_session"));
          return;
        }

        await runCompactCommand(agent, args);
      },
    },
  ];
}
