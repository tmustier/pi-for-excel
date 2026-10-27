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

type DisclosureListener = () => void;

let detailsExpanded: boolean | null = null;
const overrides = new Map<string, boolean>();
const listeners = new Set<DisclosureListener>();

const thinkingStartedAt = new Map<string, number>();
const thinkingDurations = new Map<string, number>();

function notify(): void {
  for (const listener of listeners) listener();
}

/* ── Keys ──────────────────────────────────────────────── */

export function thinkingKey(message: AssistantMessage, contentIndex: number): string {
  return `thinking:${message.timestamp}:${contentIndex}`;
}

/* ── Details mode + disclosure ─────────────────────────── */

export function areDetailsExpanded(): boolean {
  if (detailsExpanded === null) {
    try {
      detailsExpanded = localStorage.getItem(DETAILS_STORAGE_KEY) === "1";
    } catch {
      detailsExpanded = false;
    }
  }
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

/** Hand toggles override details mode and the element's initial state. */
export function isDisclosureExpanded(key: string, initiallyExpanded = false): boolean {
  return overrides.get(key) ?? (areDetailsExpanded() || initiallyExpanded);
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
