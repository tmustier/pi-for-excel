/**
 * Builtin export/compaction commands.
 */

import type { Api, Model, StopReason, Usage } from "@mariozechner/pi-ai";
import type { AgentMessage } from "@mariozechner/pi-agent-core";

import type { SlashCommand } from "../types.js";
import type { ActiveAgentProvider } from "./model.js";
import { showToast } from "../../ui/toast.js";
import { createCompactionSummaryMessage } from "../../messages/compaction.js";
import {
  createArchivedMessagesMessage,
  splitArchivedMessages,
} from "../../messages/archived-history.js";
import { getErrorMessage } from "../../utils/errors.js";
import { extractTextBlocks, summarizeContentForTranscript } from "../../utils/content.js";
import { isRecord } from "../../utils/type-guards.js";
import type { PiSidebar } from "../../ui/pi-sidebar.js";
import { getWorkbookChangeAuditLog } from "../../audit/workbook-change-audit.js";
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
  if (!isRecord(model)) return false;

  return (
    typeof model.id === "string" &&
    typeof model.name === "string" &&
    typeof model.provider === "string" &&
    typeof model.api === "string"
  );
}

function hasContent(message: AgentMessage): message is AgentMessage & { content: unknown } {
  return isRecord(message) && "content" in message;
}

function messageToTranscriptText(message: AgentMessage): string {
  if (message.role === "archivedMessages") {
    return `[archived history: ${message.archivedChatMessageCount} chat messages]`;
  }

  if (message.role === "compactionSummary") return message.summary;
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

  const entries = await getWorkbookChangeAuditLog().list(500);
  const payload = {
    exported: new Date().toISOString(),
    count: entries.length,
    entries,
  };

  const json = JSON.stringify(payload, null, 2);

  if (destination === "clipboard") {
    await navigator.clipboard.writeText(json);
    showToast(
      `Audit log copied (${entries.length} entries, ${(json.length / 1024).toFixed(0)}KB)`,
    );
    return;
  }

  triggerJsonDownload(`pi-audit-log-${new Date().toISOString().slice(0, 10)}.json`, json);
  showToast(`Downloaded audit log (${entries.length} entries)`);
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
    if (m.role === "compactionSummary") {
      return { boundaryStart: i + 1, previousSummary: m.summary };
    }
  }
  return { boundaryStart: 0 };
}

function findCutIndex(messages: AgentMessage[], boundaryStart: number, keepRecentTokens: number): number {
  let accumulated = 0;

  for (let i = messages.length - 1; i >= boundaryStart; i--) {
    accumulated += estimateTokens(messages[i]);
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

function serializeConversation(messages: AgentMessage[], limits: SerializeLimits): string {
  const parts: string[] = [];

  for (const msg of messages) {
    if (msg.role === "artifact") continue;

    if (msg.role === "user" || msg.role === "user-with-attachments") {
      const raw = typeof msg.content === "string" ? msg.content : extractTextBlocks(msg.content);
      const text = truncateMiddle(raw, limits.maxUserChars);
      if (text.trim().length > 0) parts.push(`[User]: ${text}`);
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
  const withFocus = args.customInstructions
    ? `${base}\n\nAdditional focus: ${args.customInstructions}`
    : base;

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
      description: "Export JSON (session transcript or audit log)",
      source: "builtin",
      execute: async (args: string) => {
        const parts = args.trim().split(/\s+/u).filter((part) => part.length > 0);
        const mode = parts[0]?.toLowerCase();

        if (mode === "audit" || mode === "audit-log") {
          try {
            await exportWorkbookAuditLog(parts.slice(1).join(" "));
          } catch (error: unknown) {
            showToast(`Audit export failed: ${getErrorMessage(error)}`);
          }
          return;
        }

        const agent = getActiveAgent();
        if (!agent) {
          showToast("No active session");
          return;
        }

        const msgs = agent.state.messages;
        if (msgs.length === 0) {
          showToast("No messages to export");
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
            await navigator.clipboard.writeText(json);
            showToast(
              `Transcript copied (${msgs.length} messages, ${(json.length / 1024).toFixed(0)}KB)`,
            );
          } catch (error: unknown) {
            showToast(`Copy failed: ${getErrorMessage(error)}`);
          }
          return;
        }

        triggerJsonDownload(`pi-session-${new Date().toISOString().slice(0, 10)}.json`, json);
        showToast(`Downloaded transcript (${msgs.length} messages)`);
      },
    },
  ];
}

export function createCompactCommands(getActiveAgent: ActiveAgentProvider): SlashCommand[] {
  return [
    {
      name: "compact",
      description: "Summarize older messages to free context",
      source: "builtin",
      execute: async (args: string) => {
        const agent = getActiveAgent();
        if (!agent) {
          showToast("No active session");
          return;
        }

        const allMessages = agent.state.messages;
        const {
          archivedMessages: existingArchivedMessages,
          messagesWithoutArchived,
        } = splitArchivedMessages(allMessages);

        if (messagesWithoutArchived.length < 4) {
          showToast("Too few messages to compact");
          return;
        }

        showToast("Compacting to free up context", 60000);

        const now = Date.now();
        const model = agent.state.model;
        if (!isApiModel(model)) {
          showToast("No model configured for compaction");
          return;
        }

        // IMPORTANT: use the agent's configured streamFn + api key resolver.
        // Calling pi-ai's completeSimple() directly bypasses:
        // - our CORS proxy logic (streamFn)
        // - our API key/OAuth resolution (agent.getApiKey)
        // and can crash in browser WebViews due to env key fallbacks using `process`.
        const apiKey = agent.getApiKey ? await agent.getApiKey(model.provider) : undefined;
        if (!apiKey) {
          showToast(`No API key available for ${model.provider}. Use /login or /settings.`);
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

        const runOnce = async (limits: SerializeLimits, keepRecentOverride?: number): Promise<{
          summary: string;
          keptMessages: AgentMessage[];
          messagesToArchive: AgentMessage[];
          summarizedCount: number;
        }> => {
          const keepRecent = keepRecentOverride ?? keepRecentTokens;
          const cutIndex = findCutIndex(messagesWithoutArchived, boundaryStart, keepRecent);
          const messagesToSummarize = messagesWithoutArchived.slice(boundaryStart, cutIndex);
          const keptMessages = messagesWithoutArchived.slice(cutIndex);

          if (messagesToSummarize.length === 0) {
            throw new Error("Nothing to compact");
          }

          const memoryCues = collectCompactionMemoryCues(messagesToSummarize);
          if (memoryCues.cueCount > 0 && !memoryNudgeShown) {
            const cueLabel = memoryCues.cueCount === 1 ? "cue" : "cues";
            showToast(
              `Compaction reminder: found ${memoryCues.cueCount} memory ${cueLabel} in older messages. Save durable facts to notes/ (rules via instructions) if needed.`,
              12000,
            );
            memoryNudgeShown = true;
          }

          const conversationText = serializeConversation(messagesToSummarize, limits);
          const memoryFocus = buildCompactionMemoryFocusInstruction(memoryCues);
          const customInstructions = mergeCompactionAdditionalFocus(
            userCompactionFocus,
            memoryFocus,
          );
          const promptText = buildSummarizationPrompt({
            conversationText,
            previousSummary,
            customInstructions,
          });

          const stream = await agent.streamFn(
            model,
            {
              systemPrompt: SUMMARIZATION_SYSTEM_PROMPT,
              messages: [
                {
                  role: "user",
                  content: [{ type: "text", text: promptText }],
                  timestamp: Date.now(),
                },
              ],
            },
            {
              apiKey,
              sessionId: agent.sessionId,
              maxTokens,
              // Match pi-coding-agent: don't force temperature when using reasoning,
              // since Anthropic requires temperature=1 when thinking is enabled.
              reasoning: "high",
            },
          );

          const result = await stream.result();

          if (result.stopReason === "error") {
            throw new Error(result.errorMessage || "Compaction failed");
          }

          const summary = extractTextBlocks(result.content).trim() || "Summary unavailable";

          return {
            summary,
            keptMessages,
            messagesToArchive: messagesToSummarize,
            summarizedCount: countChatMessages(messagesToSummarize),
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
            messagesToArchive: AgentMessage[];
            summarizedCount: number;
          };

          try {
            out = await runOnce(defaultLimits);
          } catch (e: unknown) {
            if (!isPromptTooLongError(e)) throw e;

            // Retry once with more aggressive truncation + keeping a larger recent tail.
            showToast("Compaction input too large — retrying with stronger truncation", 60000);

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
            messageCountBefore: out.summarizedCount,
            timestamp: now,
          });

          agent.replaceMessages([archived, compacted, ...out.keptMessages]);

          const iface = document.querySelector<PiSidebar>("pi-sidebar");
          iface?.requestUpdate();

          showToast(`Summarized ${out.summarizedCount} messages`);
        } catch (e: unknown) {
          const msg = getErrorMessage(e);
          if (msg === "Nothing to compact") {
            showToast("Nothing to compact");
            return;
          }
          showToast(`Compact failed: ${msg}`);
        }
      },
    },
  ];
}
