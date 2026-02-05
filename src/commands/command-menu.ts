/**
 * Slash command popup menu.
 *
 * Shows above the textarea when user types `/`.
 * Filters as user types, arrow keys navigate, Enter selects, Esc dismisses.
 */

import { commandRegistry, type SlashCommand } from "./types.js";

const SOURCE_BADGES: Record<string, string> = {
  builtin: "",
  extension: "ext",
  skill: "skill",
  prompt: "prompt",
};

let menuEl: HTMLElement | null = null;
let selectedIndex = 0;
let filteredCommands: SlashCommand[] = [];

export function showCommandMenu(filter: string, anchor: HTMLElement): void {
  filteredCommands = commandRegistry.list(filter);
  if (filteredCommands.length === 0) {
    hideCommandMenu();
    return;
  }
  selectedIndex = Math.min(selectedIndex, filteredCommands.length - 1);

  if (!menuEl) {
    menuEl = document.createElement("div");
    menuEl.id = "pi-command-menu";
    menuEl.className = "pi-cmd-menu";
    document.body.appendChild(menuEl);
  }

  // Position above the anchor (the input card)
  const anchorRect = anchor.getBoundingClientRect();
  menuEl.style.bottom = `${window.innerHeight - anchorRect.top + 4}px`;
  menuEl.style.left = `${anchorRect.left}px`;
  menuEl.style.width = `${anchorRect.width}px`;

  renderMenu();
  menuEl.style.display = "block";
}

export function hideCommandMenu(): void {
  if (menuEl) {
    menuEl.style.display = "none";
  }
  selectedIndex = 0;
}

export function isCommandMenuVisible(): boolean {
  return menuEl !== null && menuEl.style.display !== "none";
}

export function handleCommandMenuKey(e: KeyboardEvent): boolean {
  if (!isCommandMenuVisible()) return false;

  if (e.key === "ArrowUp") {
    e.preventDefault();
    selectedIndex = Math.max(0, selectedIndex - 1);
    renderMenu();
    return true;
  }
  if (e.key === "ArrowDown") {
    e.preventDefault();
    selectedIndex = Math.min(filteredCommands.length - 1, selectedIndex + 1);
    renderMenu();
    return true;
  }
  if (e.key === "Enter") {
    e.preventDefault();
    const cmd = filteredCommands[selectedIndex];
    if (cmd) {
      hideCommandMenu();
      // Clear the textarea
      const textarea = document.querySelector("message-editor textarea") as HTMLTextAreaElement;
      if (textarea) {
        textarea.value = "";
        textarea.dispatchEvent(new Event("input", { bubbles: true }));
      }
      cmd.execute("");
    }
    return true;
  }
  if (e.key === "Escape") {
    e.preventDefault();
    hideCommandMenu();
    return true;
  }
  if (e.key === "Tab") {
    // Tab-complete the command name
    e.preventDefault();
    const cmd = filteredCommands[selectedIndex];
    if (cmd) {
      const textarea = document.querySelector("message-editor textarea") as HTMLTextAreaElement;
      if (textarea) {
        textarea.value = `/${cmd.name} `;
        textarea.dispatchEvent(new Event("input", { bubbles: true }));
        // Move cursor to end
        textarea.selectionStart = textarea.selectionEnd = textarea.value.length;
      }
    }
    return true;
  }
  return false;
}

function renderMenu(): void {
  if (!menuEl) return;

  menuEl.innerHTML = filteredCommands.map((cmd, i) => {
    const badge = SOURCE_BADGES[cmd.source];
    const isSelected = i === selectedIndex;
    return `
      <div class="pi-cmd-item ${isSelected ? "selected" : ""}" data-index="${i}">
        <span class="pi-cmd-name">/${cmd.name}</span>
        ${badge ? `<span class="pi-cmd-badge">${badge}</span>` : ""}
        <span class="pi-cmd-desc">${cmd.description}</span>
      </div>
    `;
  }).join("");

  // Click handler
  menuEl.querySelectorAll(".pi-cmd-item").forEach((item) => {
    item.addEventListener("click", () => {
      const idx = parseInt((item as HTMLElement).dataset.index || "0");
      const cmd = filteredCommands[idx];
      if (cmd) {
        hideCommandMenu();
        const textarea = document.querySelector("message-editor textarea") as HTMLTextAreaElement;
        if (textarea) {
          textarea.value = "";
          textarea.dispatchEvent(new Event("input", { bubbles: true }));
        }
        cmd.execute("");
      }
    });
    item.addEventListener("mouseenter", () => {
      selectedIndex = parseInt((item as HTMLElement).dataset.index || "0");
      renderMenu();
    });
  });

  // Scroll selected into view
  const selectedEl = menuEl.querySelector(".selected");
  if (selectedEl) selectedEl.scrollIntoView({ block: "nearest" });
}

/**
 * Wire the command menu to a textarea.
 * Call once after the textarea is available.
 */
export function wireCommandMenu(textarea: HTMLTextAreaElement): void {
  const getAnchor = () => textarea.closest(".bg-card") as HTMLElement || textarea;

  textarea.addEventListener("input", () => {
    const val = textarea.value;
    if (val.startsWith("/") && !val.includes("\n")) {
      const filter = val.slice(1); // strip the `/`
      showCommandMenu(filter, getAnchor());
    } else {
      hideCommandMenu();
    }
  });
}
