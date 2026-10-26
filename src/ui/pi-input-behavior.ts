export interface SubmitKeyEvent {
  key: string;
  shiftKey: boolean;
  isComposing: boolean;
  keyCode: number;
}

export function getSendText(value: string): string | null {
  const text = value.trim();
  return text.length > 0 ? text : null;
}

/**
 * Enter submits the draft; Shift+Enter is a newline. An Enter that commits an
 * IME composition (WebKit reports `isComposing`, Chromium keyCode 229) only
 * commits the candidate.
 */
export function isSubmitKey(event: SubmitKeyEvent): boolean {
  if (event.key !== "Enter" || event.shiftKey) return false;
  return !event.isComposing && event.keyCode !== 229;
}

export function resolveInputAutoGrowHeight(options: {
  scrollHeight: number;
  viewportHeight: number;
  cssMaxHeight?: number;
}): number {
  const cssMaxHeight = options.cssMaxHeight;
  const viewportFallback = options.viewportHeight * (options.viewportHeight <= 520 ? 0.28 : 0.4);
  const maxHeight = typeof cssMaxHeight === "number" && Number.isFinite(cssMaxHeight) && cssMaxHeight > 0
    ? cssMaxHeight
    : viewportFallback;
  return Math.min(options.scrollHeight, maxHeight);
}
