/**
 * Keyboard shortcuts page.
 *
 * Shortcuts are grouped into logical sections and key notation adapts
 * to the current platform (macOS symbols vs Windows/Linux labels).
 */

import { t } from "../../../language/index.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";

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

function getShortcutGroups(): readonly ShortcutGroup[] {
  return [
    {
      title: t("shortcuts.section.chat"),
      shortcuts: [
        same("Enter", t("shortcuts.desc.send")),
        same(t("shortcuts.keys.enterStreaming"), t("shortcuts.desc.interrupt")),
        { mac: "⌥ Enter", win: "Alt+Enter", description: t("shortcuts.desc.queue") },
        { mac: "⌥ ↑", win: "Alt+↑", description: t("shortcuts.desc.restore_queue") },
        { mac: "⇧ Tab", win: "Shift+Tab", description: t("shortcuts.desc.cycle_thinking") },
      ],
    },
    {
      title: t("shortcuts.section.tabs"),
      shortcuts: [
        { mac: "⌘ T", win: "Ctrl+T", description: t("shortcuts.desc.new_tab") },
        { mac: "⌘ W", win: "Ctrl+W", description: t("shortcuts.desc.close_tab") },
        { mac: "⌘ ⇧ T", win: "Ctrl+Shift+T", description: t("shortcuts.desc.reopen_tab") },
        same("← →", t("shortcuts.desc.switch_tabs")),
        { mac: "⌘ ⇧ [  /  ⌘ ⇧ ]", win: "Ctrl+PgUp / PgDn", description: t("shortcuts.desc.prev_next_tab") },
      ],
    },
    {
      title: t("shortcuts.section.navigation"),
      shortcuts: [
        same("/", t("shortcuts.desc.command_menu")),
        same("↑ ↓", t("shortcuts.desc.navigate_menu")),
        same("F2", t("shortcuts.desc.focus_input")),
        same("F6", t("shortcuts.desc.toggle_focus")),
        { mac: "⇧ F6", win: "Shift+F6", description: t("shortcuts.desc.toggle_focus_reverse") },
      ],
    },
    {
      title: t("shortcuts.section.system"),
      shortcuts: [
        same("Esc", t("shortcuts.desc.dismiss")),
      ],
    },
  ];
}

// ---------------------------------------------------------------------------
// Page
// ---------------------------------------------------------------------------

export function createShortcutsPage(): SettingsShellPage {
  return {
    id: "shortcuts",
    parentId: "root",
    title: () => t("shortcuts.title"),
    subtitle: () => t("shortcuts.subtitle"),
    render: (ctx) => {
      const mac = isMacPlatform();

      const list = document.createElement("div");
      list.className = "pi-shortcuts-list";

      for (const group of getShortcutGroups()) {
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

      ctx.body.appendChild(list);
    },
  };
}
