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

export function isWindowsPlatform(): boolean {
  const { ua, platform } = readPlatformStrings();
  return platform.startsWith("Win") || ua.includes("Windows");
}
