/**
 * Per-tab turn state that is not part of the agent: the composer draft, the
 * error banner, and the non-streaming activity line (compaction, retry wait).
 *
 * Each session runtime owns one. The visible sidebar renders the active tab's
 * state (see `runtime-sidebar-binding.ts`), so switching tabs swaps the draft,
 * the banner and the activity line together with the transcript.
 */

export interface TabActivity {
  label: string;
  hint: string | null;
}

export interface TabState {
  readonly draft: string;
  /** The tab's action queue is running something (a turn, compaction, a command). */
  readonly busy: boolean;
  readonly error: string | null;
  readonly activity: TabActivity | null;
  setDraft(text: string): void;
  setError(message: string | null): void;
  setActivity(activity: TabActivity | null): void;
  setBusy(busy: boolean): void;
  /** Puts `texts` back above the current draft, separated by blank lines. */
  returnToDraft(texts: readonly string[]): void;
  subscribe(listener: () => void): () => void;
}

export function mergeIntoDraft(texts: readonly string[], draft: string): string {
  return [...texts, draft]
    .filter((text) => text.trim().length > 0)
    .join("\n\n");
}

export function createTabState(): TabState {
  let draft = "";
  let error: string | null = null;
  let activity: TabActivity | null = null;
  let busy = false;
  const listeners = new Set<() => void>();

  const emit = () => {
    for (const listener of listeners) listener();
  };

  const setDraft = (text: string) => {
    if (text === draft) return;
    draft = text;
    emit();
  };

  return {
    get draft() { return draft; },
    get error() { return error; },
    get activity() { return activity; },
    get busy() { return busy; },
    setDraft,
    setError(message) {
      if (message === error) return;
      error = message;
      emit();
    },
    setActivity(next) {
      if (next === activity) return;
      if (next && activity && next.label === activity.label && next.hint === activity.hint) return;
      activity = next;
      emit();
    },
    setBusy(next) {
      if (next === busy) return;
      busy = next;
      emit();
    },
    returnToDraft(texts) {
      if (texts.length === 0) return;
      setDraft(mergeIntoDraft(texts, draft));
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
