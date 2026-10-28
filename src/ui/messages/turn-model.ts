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
import type { AssistantMessage, ToolCall, ToolResultMessage } from "@earendil-works/pi-ai";

import type { TurnChangesMessage } from "../../messages/turn-changes.js";
import { decodeToolDetails } from "../../tools/tool-details.js";
import { isFailedToolResult } from "../tool-renderers.js";

export type TurnStepState = "live" | "done" | "failed" | "stopped";

export interface TurnToolStep {
  type: "tool";
  /** Disclosure key for the row's drill-in. */
  key: string;
  toolCall: ToolCall;
  tool: AgentTool | undefined;
  result: ToolResultMessage | undefined;
  state: TurnStepState;
}

/** Prose or thinking between tool calls: what the model said before its next step. */
export interface TurnNoteStep {
  type: "note";
  key: string;
  message: AssistantMessage;
  contentIndex: number;
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
  /** The turn is still running: more steps may come. */
  live: boolean;
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

/** The user pressed Stop during the call; the result says how far it got. */
function wasStopped(result: ToolResultMessage): boolean {
  const details = decodeToolDetails(result.details);
  return details !== undefined && "stopped" in details && details.stopped === true;
}

function isNoteChunk(message: AssistantMessage, index: number): boolean {
  const chunk = message.content[index];
  if (!chunk) return false;
  return (chunk.type === "text" && chunk.text.trim() !== "") || (chunk.type === "thinking" && chunk.thinking.trim() !== "");
}

function hasText(message: AssistantMessage, from: number, to: number): boolean {
  return message.content.slice(from, to).some((chunk) => chunk.type === "text" && chunk.text.trim() !== "");
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

    // The turn: assistant responses, their results and the receipt, up to the next other message.
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
      } else if (candidate.role !== "toolResult" && candidate.role !== "artifact") {
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
    let endedAt = firstTool.message.timestamp;
    let live = false;
    const turnIsLast = end >= all.length;
    for (const response of responses) {
      if (response.index < firstTool.index) continue;
      const { message: assistant } = response;
      const streaming = response.index === streamingIndex;
      const from = response.index === firstTool.index ? firstTool.firstCall : 0;
      const to = response.index === lastTool.index && response.index === last.index
        ? lastTool.lastCall + 1
        : response.index > lastTool.index ? 0 : assistant.content.length;
      for (let contentIndex = from; contentIndex < to; contentIndex++) {
        const chunk = assistant.content[contentIndex];
        if (!chunk) break;
        if (chunk.type === "toolCall") {
          const result = results.get(chunk.id);
          const state: TurnStepState = result
            ? (wasStopped(result) ? "stopped" : isFailedToolResult(result) ? "failed" : "done")
            : assistant.stopReason === "aborted" ? "stopped"
            : running && turnIsLast && (streaming || pendingToolCalls.has(chunk.id) || assistant.stopReason === "toolUse") ? "live"
            : "stopped";
          if (state === "live") live = true;
          if (result) endedAt = Math.max(endedAt, result.timestamp);
          steps.push({
            type: "tool",
            key: `tool:${assistant.timestamp}:${chunk.id}`,
            toolCall: chunk,
            tool: tools.find((tool) => tool.name === chunk.name),
            result,
            state,
          });
        } else if (isNoteChunk(assistant, contentIndex)) {
          steps.push({
            type: "note",
            key: `note:${assistant.timestamp}:${contentIndex}`,
            message: assistant,
            contentIndex,
            streaming: streaming && contentIndex === assistant.content.length - 1,
          });
        }
      }
      endedAt = Math.max(endedAt, assistant.timestamp);
    }
    // Between responses (the last one asked for tools, the next has not started) the turn is still running.
    if (running && turnIsLast && last.index !== streamingIndex && last.message.stopReason === "toolUse" && !changes) live = true;
    if (running && turnIsLast && last.index === streamingIndex && last.index === lastTool.index) live = true;

    // Answer: everything after the last tool call.
    const answerStarted = last.index > lastTool.index
      ? responses.some((response) => response.index > lastTool.index && hasText(response.message, 0, response.message.content.length))
      : hasText(last.message, lastTool.lastCall + 1, last.message.content.length);

    blocks.push({
      type: "turnGroup",
      key: `turn:${firstTool.message.timestamp}`,
      group: {
        key: `turn:${firstTool.message.timestamp}`,
        steps,
        startedAt: firstTool.message.timestamp,
        endedAt: live ? null : endedAt,
        live,
        answerStarted,
        changes,
      },
    });

    if (last.index === lastTool.index) {
      draw(last, lastTool.lastCall + 1, last.message.content.length);
    } else {
      for (const response of responses) {
        if (response.index > lastTool.index) draw(response, 0, response.message.content.length);
      }
    }
    index = end;
  }

  return blocks;
}
