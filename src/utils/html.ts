export function escapeHtml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

export function escapeAttr(text: string): string {
  return escapeHtml(text).replace(/`/g, "&#96;");
}

export function setSafeInnerHTML(target: Element, html: string, safetyReason: string): void {
  if (safetyReason.trim().length < 12) {
    throw new Error("setSafeInnerHTML requires a concrete safety reason");
  }
  target.innerHTML = html;
}
