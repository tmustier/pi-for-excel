/**
 * In-memory stack for recently closed tabs/sessions.
 */

export interface RecentlyClosedItem {
  id: string;
  sessionId: string;
  title: string;
  closedAt: string;
  workbookId: string | null;
}

export class RecentlyClosedStack {
  private readonly limit: number;
  private readonly items: RecentlyClosedItem[] = [];

  constructor(limit = 10) {
    this.limit = Math.max(1, Math.floor(limit));
  }

  push(item: RecentlyClosedItem): void {
    this.items.unshift(item);

    if (this.items.length > this.limit) {
      this.items.length = this.limit;
    }
  }

  popMostRecent(): RecentlyClosedItem | null {
    return this.items.shift() ?? null;
  }

  removeById(id: string): RecentlyClosedItem | null {
    const idx = this.items.findIndex((item) => item.id === id);
    if (idx < 0) return null;

    const [removed] = this.items.splice(idx, 1);
    return removed ?? null;
  }

  removeBySessionId(sessionId: string): RecentlyClosedItem | null {
    const idx = this.items.findIndex((item) => item.sessionId === sessionId);
    if (idx < 0) return null;

    const [removed] = this.items.splice(idx, 1);
    return removed ?? null;
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
}
