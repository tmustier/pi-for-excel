/**
 * Recently closed tabs, newest first, kept per workbook so they survive a reload.
 */

import { Type, type Static } from "typebox";
import { Value } from "typebox/value";

import type { SettingsStore } from "../storage/local/settings-store.js";

const RECENTLY_CLOSED_PREFIX = "workbook.recentlyClosed.v1.";
const GLOBAL_WORKBOOK_KEY = "__global__";

const recentlyClosedItemSchema = Type.Object({
  id: Type.String(),
  sessionId: Type.String({ minLength: 1 }),
  title: Type.String(),
  closedAt: Type.String(),
  workbookId: Type.Union([Type.String(), Type.Null()]),
  /** The tab's "Chat N" number when it was closed, so reopening restores the same label. */
  tabNumber: Type.Optional(Type.Integer({ minimum: 1 })),
});

export type RecentlyClosedItem = Static<typeof recentlyClosedItemSchema>;

export function recentlyClosedKey(workbookId: string | null): string {
  const suffix = workbookId && workbookId.trim().length > 0 ? workbookId : GLOBAL_WORKBOOK_KEY;
  return `${RECENTLY_CLOSED_PREFIX}${suffix}`;
}

/**
 * Reads the recently closed list for a workbook. Malformed entries are dropped;
 * a failed read rejects so the caller does not overwrite a list it could not see.
 */
export async function loadRecentlyClosed(
  settings: Pick<SettingsStore, "get">,
  workbookId: string | null,
): Promise<RecentlyClosedItem[]> {
  const stored: unknown = await settings.get(recentlyClosedKey(workbookId));
  if (!Array.isArray(stored)) return [];
  return stored.filter((item): item is RecentlyClosedItem => Value.Check(recentlyClosedItemSchema, item));
}

export async function saveRecentlyClosed(
  settings: Pick<SettingsStore, "set" | "delete">,
  workbookId: string | null,
  items: readonly RecentlyClosedItem[],
): Promise<void> {
  const key = recentlyClosedKey(workbookId);
  if (items.length === 0) {
    await settings.delete(key);
    return;
  }
  await settings.set(key, [...items]);
}

export class RecentlyClosedStack {
  private readonly limit: number;
  private readonly items: RecentlyClosedItem[] = [];
  private readonly onChange: (items: readonly RecentlyClosedItem[]) => void;

  constructor(limit = 10, onChange: (items: readonly RecentlyClosedItem[]) => void = () => {}) {
    this.limit = Math.max(1, Math.floor(limit));
    this.onChange = onChange;
  }

  /** Replaces the list with persisted items without reporting a change. */
  load(items: readonly RecentlyClosedItem[]): void {
    this.items.length = 0;
    const seenSessions = new Set<string>();
    for (const item of items) {
      if (seenSessions.has(item.sessionId)) continue;
      seenSessions.add(item.sessionId);
      this.items.push(item);
      if (this.items.length >= this.limit) break;
    }
  }

  push(item: RecentlyClosedItem): void {
    // One entry per session: closing it again replaces the older entry.
    const existing = this.items.findIndex((candidate) => candidate.sessionId === item.sessionId);
    if (existing >= 0) this.items.splice(existing, 1);

    this.items.unshift(item);

    if (this.items.length > this.limit) {
      this.items.length = this.limit;
    }
    this.onChange(this.snapshot());
  }

  popMostRecent(): RecentlyClosedItem | null {
    const item = this.items.shift() ?? null;
    if (item) this.onChange(this.snapshot());
    return item;
  }

  removeById(id: string): RecentlyClosedItem | null {
    const idx = this.items.findIndex((item) => item.id === id);
    return this.removeAt(idx);
  }

  removeBySessionId(sessionId: string): RecentlyClosedItem | null {
    const idx = this.items.findIndex((item) => item.sessionId === sessionId);
    return this.removeAt(idx);
  }

  peekMostRecent(): RecentlyClosedItem | null {
    return this.items[0] ?? null;
  }

  snapshot(): readonly RecentlyClosedItem[] {
    return [...this.items];
  }

  get size(): number {
    return this.items.length;
  }

  private removeAt(idx: number): RecentlyClosedItem | null {
    if (idx < 0) return null;
    const [removed] = this.items.splice(idx, 1);
    this.onChange(this.snapshot());
    return removed ?? null;
  }
}
