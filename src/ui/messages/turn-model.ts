/**
 * The transcript as turns.
 *
 * A turn is a user message and everything the agent did in reply: one or
 * more assistant responses, their tool results, and the workbook-change
 * receipt. The transcript draws each turn as the lead (what the model said
 * before its first tool call), one <turn-group> holding every tool call and
 * the prose between them, then the answer (what it said after the last
 * call), so the record of what the agent did reads as one block no matter
 * how many responses it took.
 */

import type { AgentMessage, AgentTool } from "@earendil-works/pi-agent-core";
import type { AssistantMessage, TextContent, ThinkingContent, ToolCall, ToolResultMessage } from "@earendil-works/pi-ai";

import type { TurnChangesMessage } from "../../messages/turn-changes.js";
import { decodeToolDetails } from "../../tools/tool-details.js";
import { isFailedToolResult } from "../tool-renderers.js";
import type { ToolRowTense } from "../tool-rows.js";
import { thinkingKey } from "./transcript-state.js";

export interface TurnToolStep {
  type: "tool";
  /** Disclosure key for the row's drill-in. */
  key: string;
  toolCall: ToolCall;
  tool: AgentTool | undefined;
  result: ToolResultMessage | undefined;
  state: ToolRowTense;
}

/** Prose or thinking between tool calls: what the model said before its next step. */
export interface TurnNoteStep {
  type: "note";
  key: string;
  content: TextContent | ThinkingContent;
  /** The block is still being streamed. */
  streaming: boolean;
}

export type TurnStep = TurnToolStep | TurnNoteStep;

export interface TurnGroupModel {
  /** Disclosure key for the group's open state. */
  key: string;
  steps: TurnStep[];
  startedAt: number;
  /** When the last step finished; null while the turn runs. */
  endedAt: number | null;
  /** The answer has started, so the group collapses itself. */
  answerStarted: boolean;
  changes: TurnChangesMessage | null;
}

/** A response, or the slice of one, drawn as a message. */
export interface MessageBlock {
  type: "message";
  key: string;
  message: AgentMessage;
  streaming: boolean;
  /** Content blocks to draw, for a response split around its tool calls. */
  contentFrom?: number;
  contentTo?: number;
}

export type TranscriptBlock =
  | MessageBlock
  | { type: "turnGroup"; key: string; group: TurnGroupModel };

export interface BuildTurnsOptions {
  messages: readonly AgentMessage[];
  /** The response being streamed, not yet in `messages`. */
  streamingMessage: AssistantMessage | null;
  tools: readonly AgentTool[];
  results: ReadonlyMap<string, ToolResultMessage>;
  pendingToolCalls: ReadonlySet<string>;
  /** The agent is working: calls without a result are running or about to. */
  running: boolean;
}

interface Response {
  message: AssistantMessage;
  index: number;
  firstCall: number;
  lastCall: number;
}

/**
 * Splits the transcript into blocks. Assistant responses that follow one
 * another without a user message in between form one turn; its tool calls
 * (across every response) make one group.
 */
export function buildTranscriptBlocks(options: BuildTurnsOptions): TranscriptBlock[] {
  const { messages, streamingMessage, tools, results, pendingToolCalls, running } = options;
  const blocks: TranscriptBlock[] = [];
  const all: AgentMessage[] = streamingMessage ? [...messages, streamingMessage] : [...messages];
  const streamingIndex = streamingMessage ? all.length - 1 : -1;

  let index = 0;
  while (index < all.length) {
    const message = all[index];
    if (!message) break;
    if (message.role !== "assistant") {
      // Tool results are drawn with their turn; artifacts are persistence-only.
      if (message.role !== "toolResult" && message.role !== "artifact") {
        blocks.push({ type: "message", key: `msg:${index}`, message, streaming: false });
      }
      index++;
      continue;
    }

    // Context-only updates do not interrupt a turn; a receipt closes it.
    const responses: Response[] = [];
    let changes: TurnChangesMessage | null = null;
    let end = index;
    while (end < all.length) {
      const candidate = all[end];
      if (!candidate) break;
      if (candidate.role === "assistant") {
        const calls = candidate.content.flatMap((chunk, position) => (chunk.type === "toolCall" ? [position] : []));
        responses.push({ message: candidate, index: end, firstCall: calls[0] ?? -1, lastCall: calls[calls.length - 1] ?? -1 });
      } else if (candidate.role === "turnChanges") {
        changes = candidate;
        end++;
        break;
      } else if (candidate.role !== "toolResult" && candidate.role !== "artifact"
        && candidate.role !== "autoContext" && candidate.role !== "system") {
        break;
      }
      end++;
    }

    const draw = (response: Response, contentFrom: number, contentTo: number): void => {
      if (contentFrom >= contentTo && contentFrom > 0) {
        // An empty tail is drawn only for the response's error or stop marker.
        const { stopReason } = response.message;
        if (contentTo < response.message.content.length || (stopReason !== "error" && stopReason !== "aborted")) return;
      }
      blocks.push({
        type: "message",
        key: `msg:${response.index}:${contentFrom}`,
        message: response.message,
        streaming: response.index === streamingIndex,
        contentFrom,
        contentTo,
      });
    };

    const firstTool = responses.find((response) => response.firstCall >= 0);
    const lastTool = [...responses].reverse().find((response) => response.firstCall >= 0);
    const last = responses[responses.length - 1];
    if (!firstTool || !lastTool || !last) {
      for (const response of responses) draw(response, 0, response.message.content.length);
      if (changes) blocks.push({ type: "message", key: `msg:${index}:changes`, message: changes, streaming: false });
      index = end;
      continue;
    }

    // Lead: everything before the first tool call.
    for (const response of responses) {
      if (response.index < firstTool.index) draw(response, 0, response.message.content.length);
    }
    if (firstTool.firstCall > 0) draw(firstTool, 0, firstTool.firstCall);

    // Group: every tool call, and the prose between the first and the last.
    const steps: TurnStep[] = [];
    let endedAt = message.timestamp;
    let hasLiveTool = false;
    const terminated = last.message.stopReason === "error" || last.message.stopReason === "aborted";
    const active = running && end === all.length && !changes && !terminated;
    for (const response of responses) {
      if (response.index < firstTool.index) continue;
      const { message: assistant } = response;
      const streaming = response.index === streamingIndex;
      const from = response.index === firstTool.index ? firstTool.firstCall : 0;
      const to = response.index === lastTool.index ? lastTool.lastCall + 1
        : response.index > lastTool.index ? 0 : assistant.content.length;
      for (let contentIndex = from; contentIndex < to; contentIndex++) {
        const chunk = assistant.content[contentIndex];
        if (!chunk) break;
        if (chunk.type === "toolCall") {
          const result = results.get(chunk.id);
          const details = decodeToolDetails(result?.details);
          const stopped = details !== undefined && "stopped" in details && details.stopped === true;
          const state: ToolRowTense = result
            ? (stopped ? "stopped" : isFailedToolResult(result) ? "failed" : "done")
            : assistant.stopReason === "error" ? "failed"
            : assistant.stopReason === "aborted" ? "stopped"
            : active && (streaming || pendingToolCalls.has(chunk.id) || assistant.stopReason === "toolUse") ? "live"
            : "stopped";
          if (state === "live") hasLiveTool = true;
          if (result) endedAt = Math.max(endedAt, result.timestamp);
          steps.push({
            type: "tool",
            key: `tool:${assistant.timestamp}:${chunk.id}`,
            toolCall: chunk,
            tool: tools.find((tool) => tool.name === chunk.name),
            result,
            state,
          });
        } else if ((chunk.type === "text" && chunk.text.trim() !== "")
          || (chunk.type === "thinking" && chunk.thinking.trim() !== "")) {
          steps.push({
            type: "note",
            key: thinkingKey(assistant, contentIndex),
            content: chunk,
            streaming: streaming && contentIndex === assistant.content.length - 1,
          });
        }
      }
      endedAt = Math.max(endedAt, assistant.timestamp);
    }
    // Text after the last call starts the answer; thinking alone does not.
    const answerStarted = !hasLiveTool && responses.some((response) => response.index >= lastTool.index
      && response.message.content.slice(response.index === lastTool.index ? lastTool.lastCall + 1 : 0)
        .some((chunk) => chunk.type === "text" && chunk.text.trim() !== ""));
    const live = active && (hasLiveTool || (!answerStarted
      && (last.index === streamingIndex || last.message.stopReason === "toolUse")));

    blocks.push({
      type: "turnGroup",
      key: `turn:${firstTool.message.timestamp}`,
      group: {
        key: `turn:${firstTool.message.timestamp}`,
        steps,
        startedAt: message.timestamp,
        endedAt: live ? null : endedAt,
        answerStarted,
        changes,
      },
    });

    for (const response of responses) {
      if (response.index >= lastTool.index) {
        draw(response, response.index === lastTool.index ? lastTool.lastCall + 1 : 0, response.message.content.length);
      }
    }
    index = end;
  }

  return blocks;
}
