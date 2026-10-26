/**
 * Transcript view state that must outlive any one rendered element.
 *
 * The same message is drawn by different elements over its life (streaming
 * container → message list, tab switches, tool results arriving), so what
 * the user chose to see cannot live in those elements:
 *
 * - Details mode (Ctrl/Cmd+O): whether tool cards, thinking blocks and tool
 *   groups start expanded. Persisted across reloads.
 * - Disclosure overrides: a card, block or group the user toggled by hand,
 *   until the next Ctrl/Cmd+O.
 * - Thinking timings observed while a response streamed, so the finished
 *   block can say "Thought for 12s".
 *
 * Keys combine the assistant message timestamp with a content position or
 * tool-call id, because tool-call ids are only unique per provider response.
 */

import type { AssistantMessage } from "@earendil-works/pi-ai";

const DETAILS_STORAGE_KEY = "pi.transcript.details-expanded";

/** Groups this large start collapsed while details are collapsed. */
const GROUP_COLLAPSE_THRESHOLD = 3;

type DisclosureListener = () => void;

let detailsExpanded: boolean | null = null;
const overrides = new Map<string, boolean>();
const listeners = new Set<DisclosureListener>();

const thinkingStartedAt = new Map<string, number>();
const thinkingDurations = new Map<string, number>();

function readStoredDetailsExpanded(): boolean {
  try {
    return localStorage.getItem(DETAILS_STORAGE_KEY) === "1";
  } catch {
    return false;
  }
}

function notify(): void {
  for (const listener of listeners) listener();
}

/* ── Keys ──────────────────────────────────────────────── */

export function toolCardKey(message: AssistantMessage, toolCallId: string): string {
  return `tool:${message.timestamp}:${toolCallId}`;
}

export function toolGroupKey(message: AssistantMessage, leaderToolCallId: string): string {
  return `group:${message.timestamp}:${leaderToolCallId}`;
}

export function thinkingKey(message: AssistantMessage, contentIndex: number): string {
  return `thinking:${message.timestamp}:${contentIndex}`;
}

/* ── Details mode + disclosure ─────────────────────────── */

export function areDetailsExpanded(): boolean {
  detailsExpanded ??= readStoredDetailsExpanded();
  return detailsExpanded;
}

/** Flip details mode; every card, block and group follows it again. */
export function toggleDetailsExpanded(): boolean {
  detailsExpanded = !areDetailsExpanded();
  overrides.clear();
  try {
    localStorage.setItem(DETAILS_STORAGE_KEY, detailsExpanded ? "1" : "0");
  } catch {
    // Storage unavailable (private mode): the mode still applies this session.
  }
  notify();
  return detailsExpanded;
}

/** Tool cards and thinking blocks follow details mode unless toggled by hand. */
export function isDisclosureExpanded(key: string): boolean {
  return overrides.get(key) ?? areDetailsExpanded();
}

/** Groups of three or more start collapsed while details are collapsed. */
export function isToolGroupExpanded(key: string, memberCount: number): boolean {
  return overrides.get(key) ?? (areDetailsExpanded() || memberCount < GROUP_COLLAPSE_THRESHOLD);
}

export function setDisclosureExpanded(key: string, expanded: boolean): void {
  overrides.set(key, expanded);
  notify();
}

export function subscribeDisclosure(listener: DisclosureListener): () => void {
  listeners.add(listener);
  return () => listeners.delete(listener);
}

/** "⌘O" on macOS, "Ctrl+O" elsewhere (the handler accepts either modifier). */
export function formatDetailsShortcut(): string {
  if (typeof navigator === "undefined") return "Ctrl+O";
  const mac = navigator.platform.startsWith("Mac") || navigator.userAgent.includes("Macintosh");
  return mac ? "⌘O" : "Ctrl+O";
}

/* ── Thinking timings ──────────────────────────────────── */

export function recordThinkingStart(key: string, at: number): void {
  thinkingStartedAt.set(key, at);
}

export function recordThinkingEnd(key: string, at: number): void {
  const startedAt = thinkingStartedAt.get(key);
  if (startedAt === undefined) return;
  thinkingStartedAt.delete(key);
  thinkingDurations.set(key, at - startedAt);
}

/** End every block still open for this message (response ended or aborted mid-thought). */
export function endOpenThinking(message: AssistantMessage, at: number): void {
  const prefix = `thinking:${message.timestamp}:`;
  // Deleting the current entry while iterating a Map is safe.
  for (const key of thinkingStartedAt.keys()) {
    if (key.startsWith(prefix)) recordThinkingEnd(key, at);
  }
}

/** Forget starts we can no longer see end (the sidebar switched to another agent). */
export function forgetOpenThinking(): void {
  thinkingStartedAt.clear();
}

/** Duration of a block whose start and end were both observed, if any. */
export function getThinkingDurationMs(key: string): number | undefined {
  return thinkingDurations.get(key);
}
