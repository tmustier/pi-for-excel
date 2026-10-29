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

export function thinkingKey(message: AssistantMessage, contentIndex: number): string {
  return `thinking:${message.timestamp}:${contentIndex}`;
}

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

export function toggleDetailsExpanded(): boolean {
  detailsExpanded = !areDetailsExpanded();
  overrides.clear();
  try {
    localStorage.setItem(DETAILS_STORAGE_KEY, detailsExpanded ? "1" : "0");
  } catch {
  }
  notify();
  return detailsExpanded;
}

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

export function formatDetailsShortcut(): string {
  if (typeof navigator === "undefined") return "Ctrl+O";
  const mac = navigator.platform.startsWith("Mac") || navigator.userAgent.includes("Macintosh");
  return mac ? "⌘O" : "Ctrl+O";
}

export function recordThinkingStart(key: string, at: number): void {
  thinkingStartedAt.set(key, at);
}

export function recordThinkingEnd(key: string, at: number): void {
  const startedAt = thinkingStartedAt.get(key);
  if (startedAt === undefined) return;
  thinkingStartedAt.delete(key);
  thinkingDurations.set(key, at - startedAt);
}

export function endOpenThinking(message: AssistantMessage, at: number): void {
  const prefix = `thinking:${message.timestamp}:`;
  for (const key of thinkingStartedAt.keys()) {
    if (key.startsWith(prefix)) recordThinkingEnd(key, at);
  }
}

export function forgetOpenThinking(): void {
  thinkingStartedAt.clear();
}

export function getThinkingDurationMs(key: string): number | undefined {
  return thinkingDurations.get(key);
}
