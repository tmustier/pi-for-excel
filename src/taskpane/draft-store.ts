import type { SettingsAccess } from "../storage/local/settings-store.js";
import type { TabState } from "./tab-state.js";

const DRAFT_KEY_PREFIX = "composer.draft.";
const HELD_KEY_PREFIX = "composer.held.";
const ATTACHMENTS_KEY_PREFIX = "composer.attachments.";
const SAVE_DELAY_MS = 400;

interface SessionIdentity {
  getSessionId: () => string;
  subscribe: (listener: () => void) => () => void;
}

export interface DraftStore {
  bind(tab: TabState, session: SessionIdentity): () => void;
}

export function createDraftStore(settings: SettingsAccess): DraftStore {

  const write = async (sessionId: string, draft: string, held: readonly string[], attachments: readonly string[]): Promise<void> => {
    const draftKey = `${DRAFT_KEY_PREFIX}${sessionId}`;
    const heldKey = `${HELD_KEY_PREFIX}${sessionId}`;
    await (draft.trim().length > 0 ? settings.set(draftKey, draft) : settings.delete(draftKey));
    await (held.length > 0 ? settings.set(heldKey, [...held]) : settings.delete(heldKey));
    const attachmentsKey = `${ATTACHMENTS_KEY_PREFIX}${sessionId}`;
    await (attachments.length > 0 ? settings.set(attachmentsKey, [...attachments]) : settings.delete(attachmentsKey));
  };

  return {
    bind(tab, session) {
      let sessionId = session.getSessionId();
      let timer: ReturnType<typeof setTimeout> | undefined;
      let disposed = false;

      const flush = () => {
        if (timer !== undefined) clearTimeout(timer);
        timer = undefined;
        void write(sessionId, tab.draft, tab.held, tab.attachments).catch((error: Error) => {
          console.warn("[pi] Composer draft could not be saved:", error.message);
        });
      };

      const load = (id: string) => {
        void settings.get(`${DRAFT_KEY_PREFIX}${id}`).then((saved) => {
          if (disposed || id !== sessionId) return;
          if (typeof saved === "string" && saved.length > 0 && tab.draft.trim().length === 0) {
            tab.setDraft(saved);
          }
        }, () => undefined);
        void settings.get(`${HELD_KEY_PREFIX}${id}`).then((saved) => {
          if (disposed || id !== sessionId || !Array.isArray(saved) || tab.held.length > 0) return;
          tab.setHeld(saved.filter((text): text is string => typeof text === "string"));
        }, () => undefined);
        void settings.get(`${ATTACHMENTS_KEY_PREFIX}${id}`).then((saved) => {
          if (disposed || id !== sessionId || !Array.isArray(saved) || tab.attachments.length > 0) return;
          tab.setAttachments(saved.filter((path): path is string => typeof path === "string"));
        }, () => undefined);
      };

      let observedDraft = tab.draft;
      let observedHeld = tab.held;
      let observedAttachments = tab.attachments;
      const unsubscribeTab = tab.subscribe(() => {
        if (tab.draft === observedDraft && tab.held === observedHeld && tab.attachments === observedAttachments) return;
        observedDraft = tab.draft;
        observedHeld = tab.held;
        observedAttachments = tab.attachments;
        if (timer !== undefined) clearTimeout(timer);
        timer = setTimeout(flush, SAVE_DELAY_MS);
      });

      const unsubscribeSession = session.subscribe(() => {
        const next = session.getSessionId();
        if (next === sessionId) return;
        const previous = sessionId;
        sessionId = next;
        if (tab.draft.trim().length > 0 || tab.held.length > 0 || tab.attachments.length > 0) {
          void write(previous, "", [], []).catch(() => undefined);
          flush();
        } else {
          load(next);
        }
      });

      load(sessionId);

      return () => {
        disposed = true;
        if (timer !== undefined) flush();
        unsubscribeTab();
        unsubscribeSession();
      };
    },
  };
}
