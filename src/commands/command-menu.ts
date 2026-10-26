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

function getComposer(): (HTMLElement & { submit?: () => void }) | null {
  return document.querySelector("pi-input");
}

function getCommandTextarea(): HTMLTextAreaElement | null {
  return document.querySelector("pi-input textarea");
}

/**
 * Choosing a command puts its name in the draft (keeping any arguments typed
 * after the name) and submits it the way Enter does, so a command refused
 * while Pi is busy leaves the draft in place.
 */
function runCommand(command: SlashCommand): void {
  hideCommandMenu();
  const textarea = getCommandTextarea();
  const composer = getComposer();
  if (!textarea || !composer?.submit) return;

  const value = textarea.value;
  const spaceIdx = value.indexOf(" ");
  const args = spaceIdx > 0 ? value.slice(spaceIdx + 1).trim() : "";
  textarea.value = args ? `/${command.name} ${args}` : `/${command.name}`;
  composer.submit();
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

/**
 * The menu owns arrows, Enter, Tab and Escape only while focus is in the
 * composer it belongs to; a handled key goes nowhere else (one Escape closes
 * the menu without also stopping the turn).
 */
export function handleCommandMenuKey(e: KeyboardEvent): boolean {
  if (!isCommandMenuVisible()) return false;
  const textarea = getCommandTextarea();
  if (!textarea || e.target !== textarea) {
    hideCommandMenu();
    return false;
  }
  if (e.isComposing || e.keyCode === 229) return false;

  const handled = handleMenuKey(e);
  if (handled) {
    e.stopPropagation();
    e.stopImmediatePropagation();
  }
  return handled;
}

function handleMenuKey(e: KeyboardEvent): boolean {

  const hasModifier = e.altKey || e.ctrlKey || e.metaKey || e.shiftKey;

  if (e.key === "ArrowUp" && !hasModifier) {
    e.preventDefault();
    selectIndex(Math.max(0, selectedIndex - 1));
    return true;
  }

  if (e.key === "ArrowDown" && !hasModifier) {
    e.preventDefault();
    selectIndex(Math.min(filteredCommands.length - 1, selectedIndex + 1));
    return true;
  }

  if (e.key === "Enter" && !e.shiftKey) {
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
    const badge = SOURCE_BADGES[command.source] ?? "";
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

    // Keep focus (and the menu) in the composer while clicking a row.
    item.addEventListener("mousedown", (event) => {
      event.preventDefault();
    });
    item.addEventListener("click", () => {
      runCommand(command);
    });

    // Only a moving pointer moves the highlight (a menu that opens under a
    // resting pointer keeps its first row), and it never re-renders rows,
    // which would swallow the click.
    item.addEventListener("mousemove", () => {
      if (selectedIndex !== index) selectIndex(index);
    });

    menuEl.appendChild(item);

    if (isSelected) {
      item.scrollIntoView({ block: "nearest" });
    }
  }
}

function selectIndex(index: number): void {
  selectedIndex = index;
  if (!menuEl) return;
  for (const item of menuEl.querySelectorAll<HTMLElement>(".pi-cmd-item")) {
    const isSelected = item.dataset.index === String(index);
    item.classList.toggle("selected", isSelected);
    item.setAttribute("aria-selected", isSelected ? "true" : "false");
    if (isSelected) item.scrollIntoView({ block: "nearest" });
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

  // Clicking elsewhere, switching tabs or opening an overlay closes the menu.
  textarea.addEventListener("blur", () => {
    hideCommandMenu();
  });
}
