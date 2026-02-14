/**
 * Keyboard shortcuts overlay.
 *
 * Shortcuts are grouped into logical sections and key notation adapts
 * to the current platform (macOS symbols vs Windows/Linux labels).
 */

import { closeOverlayById, createOverlayDialog } from "../../ui/overlay-dialog.js";
import { SHORTCUTS_OVERLAY_ID } from "../../ui/overlay-ids.js";

// ---------------------------------------------------------------------------
// Platform detection
// ---------------------------------------------------------------------------

function isMacPlatform(): boolean {
  if (typeof navigator === "undefined") return false;

  // navigator.platform is deprecated but widely supported; userAgentData
  // is the modern replacement but not available in all WebViews yet.
  const ua = navigator.userAgent ?? "";
  const platform = navigator.platform ?? "";

  return platform.startsWith("Mac") || ua.includes("Macintosh");
}

// ---------------------------------------------------------------------------
// Shortcut data
// ---------------------------------------------------------------------------

interface ShortcutEntry {
  mac: string;
  win: string;
  description: string;
}

interface ShortcutGroup {
  title: string;
  shortcuts: ShortcutEntry[];
}

/** Key that displays identically on both platforms. */
function same(key: string, description: string): ShortcutEntry {
  return { mac: key, win: key, description };
}

const SHORTCUT_GROUPS: readonly ShortcutGroup[] = [
  {
    title: "Chat",
    shortcuts: [
      same("Enter", "Send message"),
      same("Enter (while streaming)", "Interrupt and redirect"),
      { mac: "⌥ Enter", win: "Alt+Enter", description: "Queue follow-up" },
      { mac: "⇧ Tab", win: "Shift+Tab", description: "Cycle thinking level" },
    ],
  },
  {
    title: "Tabs",
    shortcuts: [
      { mac: "⌘ T", win: "Ctrl+T", description: "New tab" },
      { mac: "⌘ W", win: "Ctrl+W", description: "Close tab" },
      { mac: "⌘ ⇧ T", win: "Ctrl+Shift+T", description: "Reopen closed tab" },
      same("← →", "Switch tabs (exit input first)"),
      { mac: "⌘ ⇧ [  /  ⌘ ⇧ ]", win: "Ctrl+PgUp / PgDn", description: "Previous / next tab" },
    ],
  },
  {
    title: "Navigation",
    shortcuts: [
      same("/", "Open command menu"),
      same("↑ ↓", "Navigate menu items"),
      same("F2", "Focus chat input"),
      same("F6", "Toggle focus: sheet ↔ sidebar"),
      { mac: "⇧ F6", win: "Shift+F6", description: "Toggle focus (reverse)" },
    ],
  },
  {
    title: "System",
    shortcuts: [
      same("Esc", "Dismiss overlay / stop generation / exit input"),
    ],
  },
];

// ---------------------------------------------------------------------------
// Overlay
// ---------------------------------------------------------------------------

export function showShortcutsDialog(): void {
  if (closeOverlayById(SHORTCUTS_OVERLAY_ID)) {
    return;
  }

  const mac = isMacPlatform();
  const dialog = createOverlayDialog({
    overlayId: SHORTCUTS_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-shortcuts-dialog",
  });

  const title = document.createElement("h2");
  title.className = "pi-overlay-title";
  title.textContent = "Keyboard Shortcuts";

  const list = document.createElement("div");
  list.className = "pi-shortcuts-list";

  for (const group of SHORTCUT_GROUPS) {
    const section = document.createElement("div");
    section.className = "pi-shortcuts-section";

    const header = document.createElement("div");
    header.className = "pi-shortcuts-section-header";
    header.textContent = group.title;
    section.appendChild(header);

    for (const shortcut of group.shortcuts) {
      const row = document.createElement("div");
      row.className = "pi-shortcuts-row";

      const keyEl = document.createElement("kbd");
      keyEl.className = "pi-shortcuts-key";
      keyEl.textContent = mac ? shortcut.mac : shortcut.win;

      const descEl = document.createElement("span");
      descEl.className = "pi-shortcuts-desc";
      descEl.textContent = shortcut.description;

      row.append(keyEl, descEl);
      section.appendChild(row);
    }

    list.appendChild(section);
  }

  const closeButton = document.createElement("button");
  closeButton.type = "button";
  closeButton.className = "pi-overlay-btn pi-overlay-btn--ghost pi-overlay-btn--full";
  closeButton.textContent = "Close";

  dialog.card.append(title, list, closeButton);
  closeButton.addEventListener("click", dialog.close);

  dialog.mount();
}
