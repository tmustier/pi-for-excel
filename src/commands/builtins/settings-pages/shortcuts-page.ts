/**
 * Keyboard shortcuts page.
 *
 * Shortcuts are grouped into logical sections and key notation adapts
 * to the current platform (macOS symbols vs Windows/Linux labels).
 */

import { t } from "../../../language/index.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import { isMacPlatform, isOfficeOnline } from "../../../utils/platform.js";

// ---------------------------------------------------------------------------
// Shortcut data
// ---------------------------------------------------------------------------

interface ShortcutEntry {
  mac: string;
  win: string;
  description: string;
  /** The browser takes this chord in Excel on the web, so it is not listed there. */
  browserReserved?: boolean;
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
        { mac: "⌃ ⌥ T", win: "Ctrl+Alt+T", description: t("shortcuts.desc.cycle_thinking") },
        { mac: "⌘ O", win: "Ctrl+O", description: t("shortcuts.desc.toggle_details") },
      ],
    },
    {
      title: t("shortcuts.section.tabs"),
      shortcuts: [
        { mac: "⌘ T", win: "Ctrl+T", description: t("shortcuts.desc.new_tab"), browserReserved: true },
        { mac: "⌘ W", win: "Ctrl+W", description: t("shortcuts.desc.close_tab"), browserReserved: true },
        { mac: "⌘ ⇧ T", win: "Ctrl+Shift+T", description: t("shortcuts.desc.reopen_tab"), browserReserved: true },
        same("← →", t("shortcuts.desc.switch_tabs")),
        { mac: "⌘ ⇧ [  /  ⌘ ⇧ ]", win: "Ctrl+PgUp / PgDn", description: t("shortcuts.desc.prev_next_tab"), browserReserved: true },
        { mac: "⇧ F10", win: "Shift+F10", description: t("shortcuts.desc.tab_menu") },
      ],
    },
    {
      title: t("shortcuts.section.navigation"),
      shortcuts: [
        same("/", t("shortcuts.desc.command_menu")),
        same("↑ ↓", t("shortcuts.desc.navigate_menu")),
        same("F2", t("shortcuts.desc.focus_input")),
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
      const web = isOfficeOnline();

      const list = document.createElement("div");
      list.className = "pi-shortcuts-list";

      if (web) {
        const note = document.createElement("p");
        note.className = "pi-shortcuts-note";
        note.textContent = t("shortcuts.webNote");
        list.appendChild(note);
      }

      for (const group of getShortcutGroups()) {
        const section = document.createElement("div");
        section.className = "pi-shortcuts-section";

        const header = document.createElement("div");
        header.className = "pi-shortcuts-section-header";
        header.textContent = group.title;
        section.appendChild(header);

        for (const shortcut of group.shortcuts.filter((entry) => !(web && entry.browserReserved))) {
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
