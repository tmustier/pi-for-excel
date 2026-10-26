/** Platform helpers for key labels (⌘ on macOS, Ctrl elsewhere). */

export function isMacPlatform(): boolean {
  if (typeof navigator === "undefined") return false;

  // navigator.platform is deprecated but widely supported; userAgentData
  // is the modern replacement but not available in all WebViews yet.
  const ua = navigator.userAgent ?? "";
  const platform = navigator.platform ?? "";

  return platform.startsWith("Mac") || ua.includes("Macintosh");
}

/**
 * True in Excel on the web, where the browser keeps tab chords (Ctrl/⌘+T, W,
 * Shift+T, Ctrl+PageUp/PageDown) for itself before the taskpane sees them.
 */
export function isOfficeOnline(): boolean {
  try {
    return typeof Office !== "undefined" && String(Office.context?.platform) === "OfficeOnline";
  } catch {
    return false;
  }
}
