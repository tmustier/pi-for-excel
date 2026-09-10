/**
 * Lightweight debug flags.
 *
 * Keep this module dependency-free and synchronous so UI code (status bar)
 * can read flags without awaiting IndexedDB.
 */

import type { Usage } from "@earendil-works/pi-ai";

const STORAGE_KEY = "pi-excel.debug";

type DebugState = {
  enabled: boolean;
};

function safeGetItem(key: string): string | null {
  try {
    return window.localStorage.getItem(key);
  } catch {
    return null;
  }
}

function safeSetItem(key: string, value: string): void {
  try {
    window.localStorage.setItem(key, value);
  } catch {
    // ignore (private mode / disabled storage)
  }
}

function readState(): DebugState {
  const raw = safeGetItem(STORAGE_KEY);
  if (!raw) return { enabled: false };

  if (raw === "1" || raw === "true") return { enabled: true };
  if (raw === "0" || raw === "false") return { enabled: false };

  // Backwards/forward compatibility for potential JSON expansion.
  try {
    const parsed = JSON.parse(raw) as unknown;
    if (typeof parsed === "object" && parsed !== null && "enabled" in parsed) {
      const enabled = (parsed as { enabled?: unknown }).enabled;
      return { enabled: enabled === true };
    }
  } catch {
    // ignore
  }

  return { enabled: false };
}

function writeState(state: DebugState): void {
  safeSetItem(STORAGE_KEY, state.enabled ? "1" : "0");
}

export function isDebugEnabled(): boolean {
  return readState().enabled;
}

export function setDebugEnabled(enabled: boolean): void {
  writeState({ enabled });
  // Status bar listens to this.
  document.dispatchEvent(new Event("pi:status-update"));
  // For future use by other UI components.
  document.dispatchEvent(new Event("pi:debug-changed"));
}

export function toggleDebugEnabled(): boolean {
  const next = !isDebugEnabled();
  setDebugEnabled(next);
  return next;
}

export function formatK(n: number): string {
  if (!Number.isFinite(n)) return "?";
  if (Math.abs(n) >= 100_000) return `${(n / 1000).toFixed(0)}k`;
  if (Math.abs(n) >= 10_000) return `${(n / 1000).toFixed(1)}k`;
  if (Math.abs(n) >= 1000) return `${(n / 1000).toFixed(2)}k`;
  return String(Math.round(n));
}

export function formatUsageDebug(usage: Usage): string {
  // Keep this short — it renders in the status bar.
  return `in:${formatK(usage.input)} out:${formatK(usage.output)} cr:${formatK(usage.cacheRead)} cw:${formatK(usage.cacheWrite)}`;
}
