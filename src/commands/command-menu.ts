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
  integration: "integration",
  prompt: "prompt",
};

let menuEl: HTMLDivElement | null = null;
let selectedIndex = 0;
let filteredCommands: SlashCommand[] = [];

function getCommandTextarea(): HTMLTextAreaElement | null {
  return document.querySelector("pi-input textarea");
}

function clearCommandTextarea(): void {
  const textarea = getCommandTextarea();
  if (!textarea) {
    return;
  }

  textarea.value = "";
  textarea.dispatchEvent(new Event("input", { bubbles: true }));
}

function runCommand(command: SlashCommand): void {
  hideCommandMenu();
  clearCommandTextarea();

  document.dispatchEvent(
    new CustomEvent("pi:command-run", { detail: { name: command.name, args: "" } }),
  );
}

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
    menuEl.setAttribute("role", "listbox");
    menuEl.setAttribute("aria-label", "Slash commands");
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
    const command = filteredCommands[selectedIndex];
    if (command) {
      runCommand(command);
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
    const command = filteredCommands[selectedIndex];
    const textarea = getCommandTextarea();
    if (command && textarea) {
      textarea.value = `/${command.name} `;
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
      textarea.selectionStart = textarea.value.length;
      textarea.selectionEnd = textarea.value.length;
    }
    return true;
  }

  return false;
}

function renderMenu(): void {
  if (!menuEl) {
    return;
  }

  menuEl.replaceChildren();

  for (const [index, command] of filteredCommands.entries()) {
    const badge = SOURCE_BADGES[command.source];
    const isSelected = index === selectedIndex;

    const item = document.createElement("button");
    item.type = "button";
    item.className = `pi-cmd-item${isSelected ? " selected" : ""}`;
    item.dataset.index = String(index);
    item.setAttribute("role", "option");
    item.setAttribute("aria-selected", isSelected ? "true" : "false");

    const name = document.createElement("span");
    name.className = "pi-cmd-name";
    name.textContent = `/${command.name}`;
    item.appendChild(name);

    if (badge.length > 0) {
      const badgeEl = document.createElement("span");
      badgeEl.className = "pi-cmd-badge";
      badgeEl.textContent = badge;
      item.appendChild(badgeEl);
    }

    const description = document.createElement("span");
    description.className = "pi-cmd-desc";
    description.textContent = command.description;
    item.appendChild(description);

    item.addEventListener("click", () => {
      runCommand(command);
    });

    item.addEventListener("mouseenter", () => {
      selectedIndex = index;
      renderMenu();
    });

    menuEl.appendChild(item);

    if (isSelected) {
      item.scrollIntoView({ block: "nearest" });
    }
  }
}

/**
 * Wire the command menu to a textarea.
 * Call once after the textarea is available.
 */
export function wireCommandMenu(textarea: HTMLTextAreaElement): void {
  const getAnchor = (): HTMLElement =>
    textarea.closest<HTMLElement>(".pi-input-card")
    ?? textarea.closest<HTMLElement>(".bg-card")
    ?? textarea;

  textarea.addEventListener("input", () => {
    const value = textarea.value;
    if (value.startsWith("/") && !value.includes("\n")) {
      const filter = value.slice(1); // strip the `/`
      showCommandMenu(filter, getAnchor());
      return;
    }

    hideCommandMenu();
  });
}
