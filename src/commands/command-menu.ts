import { t } from "../language/index.js";
import { commandRegistry, type SlashCommand } from "./types.js";
import type { PiInput } from "../ui/pi-input.js";

const SOURCE_BADGES: Record<SlashCommand["source"], string> = {
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

function runCommand(command: SlashCommand): void {
  hideCommandMenu();
  const textarea = getCommandTextarea();
  const composer = document.querySelector<PiInput>("pi-input");
  if (!textarea || !composer) return;

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
    menuEl.setAttribute("aria-label", t("command-menu.aria"));
    document.body.appendChild(menuEl);
  }


  const anchorRect = anchor.getBoundingClientRect();
  menuEl.style.bottom = `${window.innerHeight - anchorRect.top + 4}px`;
  menuEl.style.left = `${anchorRect.left}px`;
  menuEl.style.width = `${anchorRect.width}px`;

  renderMenu();
  menuEl.style.display = "block";
  const menuOffset = Math.round(window.innerHeight - menuEl.getBoundingClientRect().top);
  document.documentElement.style.setProperty("--pi-command-menu-offset", `${menuOffset}px`);
}

export function hideCommandMenu(): void {
  if (menuEl) {
    menuEl.style.display = "none";
  }
  document.documentElement.style.removeProperty("--pi-command-menu-offset");
  selectedIndex = 0;
  getCommandTextarea()?.removeAttribute("aria-activedescendant");
}

export function isCommandMenuVisible(): boolean {
  return menuEl !== null && menuEl.style.display !== "none";
}

export function handleCommandMenuKey(e: KeyboardEvent): boolean {
  if (!isCommandMenuVisible()) return false;
  const textarea = getCommandTextarea();
  if (!textarea || e.target !== textarea) {
    hideCommandMenu();
    return false;
  }
  if (e.isComposing || e.keyCode === 229) return false;

  const hasModifier = e.altKey || e.ctrlKey || e.metaKey || e.shiftKey;
  const command = filteredCommands[selectedIndex];
  if (e.key === "ArrowUp" && !hasModifier) {
    selectIndex(Math.max(0, selectedIndex - 1));
  } else if (e.key === "ArrowDown" && !hasModifier) {
    selectIndex(Math.min(filteredCommands.length - 1, selectedIndex + 1));
  } else if (e.key === "Enter" && !e.shiftKey) {
    if (command) runCommand(command);
  } else if (e.key === "Escape") {
    hideCommandMenu();
  } else if (e.key === "Tab") {
    if (command) {
      textarea.value = `/${command.name} `;
      textarea.dispatchEvent(new Event("input", { bubbles: true }));
      textarea.selectionStart = textarea.value.length;
      textarea.selectionEnd = textarea.value.length;
    }
  } else {
    return false;
  }

  e.preventDefault();
  e.stopPropagation();
  e.stopImmediatePropagation();
  return true;
}

function renderMenu(): void {
  if (!menuEl) {
    return;
  }

  menuEl.replaceChildren();

  let previousGroup: SlashCommand["menuGroup"];
  for (const [index, command] of filteredCommands.entries()) {
    const group = command.menuGroup ?? "setup";
    if (group !== previousGroup) {
      const heading = document.createElement("div");
      heading.className = "pi-cmd-group";
      heading.setAttribute("role", "presentation");
      heading.textContent = group === "chat" ? t("command-menu.group.chat")
        : group === "workbook" ? t("command-menu.group.workbook") : t("command-menu.group.setup");
      menuEl.appendChild(heading);
      previousGroup = group;
    }
    const badge = SOURCE_BADGES[command.source];
    const isSelected = index === selectedIndex;

    const item = document.createElement("button");
    item.type = "button";
    item.className = `pi-cmd-item${isSelected ? " selected" : ""}`;
    item.dataset.index = String(index);
    item.id = `pi-cmd-option-${index}`;
    item.tabIndex = -1;
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
    description.textContent = command.name === "extensions"
      ? t("command-menu.extensionsAliases") : command.description;
    description.title = description.textContent;
    item.appendChild(description);

    item.addEventListener("mousedown", (event) => {
      event.preventDefault();
    });
    item.addEventListener("click", () => {
      runCommand(command);
    });

    item.addEventListener("mousemove", () => {
      if (selectedIndex !== index) selectIndex(index);
    });

    menuEl.appendChild(item);

    if (isSelected) {
      item.scrollIntoView({ block: "nearest" });
      getCommandTextarea()?.setAttribute("aria-activedescendant", item.id);
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

export function wireCommandMenu(composer: HTMLElement): void {
  composer.addEventListener("focusin", (event) => {
    if (event.target instanceof HTMLTextAreaElement) event.target.setAttribute("aria-controls", "pi-command-menu");
  });

  document.addEventListener("pointerdown", (event) => {
    if (!isCommandMenuVisible()) return;
    const target = event.target;
    if (target instanceof Node && (menuEl?.contains(target) || composer.contains(target))) return;
    hideCommandMenu();
  }, true);

  composer.addEventListener("input", (event) => {
    const textarea = event.target;
    if (!(textarea instanceof HTMLTextAreaElement)) return;
    const anchor = textarea.closest<HTMLElement>(".pi-input-card")
      ?? textarea.closest<HTMLElement>(".bg-card")
      ?? textarea;
    const value = textarea.value;
    if (value.startsWith("/") && !value.includes("\n")) {
      const filter = value.slice(1);
      showCommandMenu(filter, anchor);
      return;
    }

    hideCommandMenu();
  });

  composer.addEventListener("focusout", (event) => {
    if (event.target instanceof HTMLTextAreaElement) hideCommandMenu();
  });
}
