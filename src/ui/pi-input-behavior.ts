export interface InputEnterState {
  key: string;
  shiftKey: boolean;
  isStreaming: boolean;
  value: string;
}

export function getSendText(value: string): string | null {
  const text = value.trim();
  return text.length > 0 ? text : null;
}

export function shouldSendOnEnter(state: InputEnterState): boolean {
  if (state.key !== "Enter" || state.shiftKey) return false;
  if (state.isStreaming) return false;
  if (!getSendText(state.value)) return false;
  return !state.value.startsWith("/");
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
