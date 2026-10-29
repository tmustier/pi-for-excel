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
  } catch {}
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
    const dark = mode === "dark" || (mode === "office" && host.resolveThemeDark() === true);
    document.documentElement.classList.toggle("dark", dark);
    document.documentElement.style.colorScheme = dark ? "dark" : "light";
  };
  apply();
  const timer = window.setInterval(apply, 1000);
  document.addEventListener(THEME_CHANGED, apply);
  window.addEventListener("storage", apply);
  return () => {
    window.clearInterval(timer);
    document.removeEventListener(THEME_CHANGED, apply);
    window.removeEventListener("storage", apply);
  };
}
