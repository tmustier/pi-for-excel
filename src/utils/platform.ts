/**
 * Desktop platform detection for user-facing copy (shortcut glyphs, terminal names).
 */

function readPlatformStrings(): { ua: string; platform: string } {
  if (typeof navigator === "undefined") return { ua: "", platform: "" };
  // navigator.platform is deprecated but widely supported; userAgentData
  // is the modern replacement but not available in all WebViews yet.
  return { ua: navigator.userAgent ?? "", platform: navigator.platform ?? "" };
}

export function isMacPlatform(): boolean {
  const { ua, platform } = readPlatformStrings();
  return platform.startsWith("Mac") || ua.includes("Macintosh");
}

export function formatThinkingShortcut(): string {
  return isMacPlatform() ? "Ctrl+Option+T" : "Ctrl+Alt+T";
}

/** Excel for Mac's taskpane: a WKWebView, whose user agent lacks Safari's "Safari/" token. */
export function isEmbeddedAppleWebView(): boolean {
  const { ua } = readPlatformStrings();
  return isMacPlatform() && !ua.includes("Safari/");
}

export function isWindowsPlatform(): boolean {
  const { ua, platform } = readPlatformStrings();
  return platform.startsWith("Win") || ua.includes("Windows");
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
