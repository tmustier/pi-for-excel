/** Theme choice is synchronous so it applies before the first taskpane paint. */
import {
  createSpreadsheetHost,
  detectSpreadsheetHost,
  getCurrentSpreadsheetHost,
} from "../host/index.js";

export type ThemeMode = "office" | "light" | "dark";
const THEME_KEY = "pi.ui.theme";
const THEME_CHANGED = "pi:theme-changed";

export function getThemeMode(): ThemeMode {
  try {
    const stored = localStorage.getItem(THEME_KEY);
    if (stored === "light" || stored === "dark") return stored;
  } catch { /* Storage unavailable: follow Office. */ }
  return "office";
}

export function setThemeMode(mode: ThemeMode): void {
  localStorage.setItem(THEME_KEY, mode);
  document.dispatchEvent(new Event(THEME_CHANGED));
}

export function installThemeModeSync(): () => void {
  if (typeof document === "undefined" || typeof window === "undefined") return () => {};
  const apply = (): void => {
    const mode = getThemeMode();
    const currentHost = getCurrentSpreadsheetHost();
    const host = currentHost.kind !== "office" && detectSpreadsheetHost() === "office"
      ? createSpreadsheetHost("office") : currentHost;
    // Office on the web can report no usable theme. Do not substitute the OS
    // preference: the workbook may still be light. Explicit Dark always works.
    const dark = mode === "dark" || (mode === "office" && host.resolveThemeDark() === true);
    document.documentElement.classList.toggle("dark", dark);
    document.documentElement.style.colorScheme = dark ? "dark" : "light";
  };
  apply();
  // Excel has no OfficeThemeChanged event (that event is Outlook-only).
  // Re-read the host's small theme snapshot to follow changes and late Office boot.
  const timer = window.setInterval(apply, 1000);
  document.addEventListener(THEME_CHANGED, apply);
  window.addEventListener("storage", apply);
  return () => {
    window.clearInterval(timer);
    document.removeEventListener(THEME_CHANGED, apply);
    window.removeEventListener("storage", apply);
  };
}
