export const PI_REQUEST_INPUT_FOCUS_EVENT = "pi:request-input-focus";

export function requestChatInputFocus(): void {
  if (typeof document === "undefined") {
    return;
  }

  document.dispatchEvent(new CustomEvent(PI_REQUEST_INPUT_FOCUS_EVENT));
}

export function moveCursorToEnd(textarea: HTMLTextAreaElement): void {
  const cursor = textarea.value.length;
  textarea.setSelectionRange(cursor, cursor);
}
