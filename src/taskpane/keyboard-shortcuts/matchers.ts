/**
 * Pure keyboard shortcut matchers + predicate helpers.
 */

export interface TabShortcutEventLike {
  key: string;
  code?: string;
  keyCode?: number;
  metaKey: boolean;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

export type ReopenShortcutEventLike = TabShortcutEventLike;
export type UndoCloseShortcutEventLike = TabShortcutEventLike;
export type CreateTabShortcutEventLike = TabShortcutEventLike;
export type CloseTabShortcutEventLike = TabShortcutEventLike;

export interface FocusInputShortcutEventLike {
  key: string;
  metaKey: boolean;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

export interface AdjacentTabShortcutEventLike {
  key: string;
  code?: string;
  keyCode?: number;
  repeat: boolean;
  metaKey: boolean;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

export interface RestoreQueuedShortcutEventLike {
  key: string;
  code?: string;
  keyCode?: number;
  repeat: boolean;
  metaKey: boolean;
  ctrlKey: boolean;
  shiftKey: boolean;
  altKey: boolean;
}

function isShortcutLetter(event: TabShortcutEventLike, letter: "t" | "w" | "z"): boolean {
  if (event.key.toLowerCase() === letter) {
    return true;
  }

  const upper = letter.toUpperCase();
  if (event.code === `Key${upper}`) {
    return true;
  }

  const keyCode = letter === "t" ? 84 : letter === "w" ? 87 : 90;
  if (event.keyCode === keyCode) {
    return true;
  }

  return false;
}

export function isCreateTabShortcut(event: CreateTabShortcutEventLike): boolean {
  if (!(event.metaKey || event.ctrlKey)) return false;
  if (event.shiftKey || event.altKey) return false;

  return isShortcutLetter(event, "t");
}

export function isCloseActiveTabShortcut(event: CloseTabShortcutEventLike): boolean {
  if (!(event.metaKey || event.ctrlKey)) return false;
  if (event.shiftKey || event.altKey) return false;

  return isShortcutLetter(event, "w");
}

export function isUndoCloseTabShortcut(event: UndoCloseShortcutEventLike): boolean {
  if (!(event.metaKey || event.ctrlKey)) return false;
  if (event.shiftKey || event.altKey) return false;

  return isShortcutLetter(event, "z");
}

export function isReopenLastClosedShortcut(event: ReopenShortcutEventLike): boolean {
  if (!(event.metaKey || event.ctrlKey)) return false;
  if (!event.shiftKey || event.altKey) return false;

  return isShortcutLetter(event, "t");
}

export function isFocusInputShortcut(event: FocusInputShortcutEventLike): boolean {
  if (event.metaKey || event.ctrlKey || event.altKey || event.shiftKey) return false;
  return event.key === "F2";
}

export function getAdjacentTabDirectionFromShortcut(
  event: AdjacentTabShortcutEventLike,
): -1 | 1 | null {
  if (event.repeat) return null;

  const key = event.key;
  const code = event.code;
  const keyCode = event.keyCode;

  // Fallback chords for hosts that swallow plain arrow keys.
  if (event.metaKey && event.shiftKey && !event.ctrlKey && !event.altKey) {
    if (key === "[") return -1;
    if (key === "]") return 1;
    return null;
  }

  if (event.ctrlKey && !event.metaKey && !event.altKey && !event.shiftKey) {
    if (key === "PageUp") return -1;
    if (key === "PageDown") return 1;
  }

  if (event.metaKey || event.ctrlKey || event.altKey || event.shiftKey) return null;

  if (key === "ArrowLeft" || key === "Left" || code === "ArrowLeft" || keyCode === 37) {
    return -1;
  }

  if (key === "ArrowRight" || key === "Right" || code === "ArrowRight" || keyCode === 39) {
    return 1;
  }

  return null;
}

export function isRestoreQueuedMessagesShortcut(event: RestoreQueuedShortcutEventLike): boolean {
  if (event.repeat) return false;
  if (!event.altKey || event.metaKey || event.ctrlKey || event.shiftKey) return false;

  if (event.key === "ArrowUp" || event.key === "Up" || event.code === "ArrowUp" || event.keyCode === 38) {
    return true;
  }

  // Some hosts/webviews translate Alt/Option+↑ into PageUp.
  return event.key === "PageUp" || event.code === "PageUp" || event.keyCode === 33;
}

export function shouldBlurEditorFromEscape(opts: {
  key: string;
  isInEditor: boolean;
  isStreaming: boolean;
}): boolean {
  if (opts.key !== "Escape" && opts.key !== "Esc") return false;
  if (!opts.isInEditor) return false;
  if (opts.isStreaming) return false;
  return true;
}

export function shouldAbortFromEscape(opts: {
  isStreaming: boolean;
  hasAgent: boolean;
  escapeClaimedByOverlay: boolean;
}): boolean {
  if (!opts.isStreaming) return false;
  if (!opts.hasAgent) return false;
  if (opts.escapeClaimedByOverlay) return false;
  return true;
}
