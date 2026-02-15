/**
 * Status-bar popovers (thinking selector + context quick actions).
 */

import type { ThinkingLevel } from "@mariozechner/pi-agent-core";

export type StatusCommandName = "compact" | "new";

type StatusPopoverKind = "thinking" | "context" | "proxy";

interface ActivePopoverState {
  kind: StatusPopoverKind;
  anchor: Element;
  popover: HTMLDivElement;
  cleanup: () => void;
}

interface ThinkingPopoverOptions {
  anchor: Element;
  description: string;
  levels: readonly ThinkingLevel[];
  activeLevel: ThinkingLevel;
  onSelectLevel: (level: ThinkingLevel) => void;
}

interface ContextPopoverOptions {
  anchor: Element;
  description: string;
  onRunCommand: (command: StatusCommandName) => void;
}

const THINKING_LEVEL_LABELS: Record<ThinkingLevel, string> = {
  off: "Off",
  minimal: "Minimal",
  low: "Low",
  medium: "Medium",
  high: "High",
  xhigh: "Max",
};

const THINKING_LEVEL_HINTS: Record<ThinkingLevel, string> = {
  off: "Fastest — no reasoning step",
  minimal: "Quick — light reasoning",
  low: "Fast — moderate reasoning",
  medium: "Balanced — solid reasoning",
  high: "Slow — thorough reasoning",
  xhigh: "Slowest — deepest reasoning",
};

let activePopover: ActivePopoverState | null = null;

function clamp(value: number, min: number, max: number): number {
  if (max < min) return min;
  return Math.min(Math.max(value, min), max);
}

function normalizeDescription(text: string): string {
  const compact = text.replace(/\s+/g, " ").trim();
  return compact.length > 0 ? compact : "No details available.";
}

function createPopoverBase(kind: StatusPopoverKind): HTMLDivElement {
  const popover = document.createElement("div");
  popover.className = `pi-status-popover pi-status-popover--${kind}`;
  popover.setAttribute("role", "dialog");
  popover.setAttribute("aria-live", "polite");
  return popover;
}

function positionPopover(popover: HTMLDivElement, anchor: Element): void {
  const anchorRect = anchor.getBoundingClientRect();
  const viewportWidth = document.documentElement.clientWidth;
  const viewportHeight = document.documentElement.clientHeight;

  popover.style.left = "0px";
  popover.style.top = "0px";
  popover.style.visibility = "hidden";

  const popoverWidth = popover.offsetWidth;
  const popoverHeight = popover.offsetHeight;

  const maxLeft = Math.max(8, viewportWidth - popoverWidth - 8);
  const left = clamp(anchorRect.right - popoverWidth, 8, maxLeft);

  const preferredTop = anchorRect.top - popoverHeight - 8;
  const fallbackTop = anchorRect.bottom + 8;
  const maxTop = Math.max(8, viewportHeight - popoverHeight - 8);
  const top = preferredTop >= 8
    ? preferredTop
    : clamp(fallbackTop, 8, maxTop);

  popover.style.left = `${left}px`;
  popover.style.top = `${top}px`;
  popover.style.visibility = "visible";
}

function closeIfOpen(state: ActivePopoverState | null): void {
  if (!state) return;
  state.cleanup();
  state.popover.remove();
}

export function closeStatusPopover(): void {
  const state = activePopover;
  activePopover = null;
  closeIfOpen(state);
}

function mountPopover(kind: StatusPopoverKind, anchor: Element, popover: HTMLDivElement): void {
  closeStatusPopover();

  document.body.appendChild(popover);
  positionPopover(popover, anchor);

  const reposition = () => {
    if (!activePopover || activePopover.popover !== popover) return;
    positionPopover(popover, anchor);
  };

  const onMouseDown = (event: MouseEvent) => {
    const target = event.target;
    if (!(target instanceof Node)) return;

    if (popover.contains(target)) return;
    if (anchor instanceof Node && anchor.contains(target)) return;

    closeStatusPopover();
  };

  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    closeStatusPopover();
  };

  document.addEventListener("mousedown", onMouseDown, true);
  window.addEventListener("keydown", onKeyDown, true);
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);

  activePopover = {
    kind,
    anchor,
    popover,
    cleanup: () => {
      document.removeEventListener("mousedown", onMouseDown, true);
      window.removeEventListener("keydown", onKeyDown, true);
      window.removeEventListener("resize", reposition);
      window.removeEventListener("scroll", reposition, true);
    },
  };
}

function shouldToggle(kind: StatusPopoverKind, anchor: Element): boolean {
  return Boolean(
    activePopover
    && activePopover.kind === kind
    && activePopover.anchor === anchor,
  );
}

function createDescriptionBlock(text: string): HTMLParagraphElement {
  const description = document.createElement("p");
  description.className = "pi-status-popover__description";
  description.textContent = normalizeDescription(text);
  return description;
}

export function toggleThinkingPopover(opts: ThinkingPopoverOptions): void {
  if (shouldToggle("thinking", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const popover = createPopoverBase("thinking");

  const title = document.createElement("h3");
  title.className = "pi-status-popover__title";
  title.textContent = "Thinking level";

  const description = createDescriptionBlock(opts.description);

  const list = document.createElement("div");
  list.className = "pi-status-popover__list";

  for (const level of opts.levels) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "pi-status-popover__item";
    if (level === opts.activeLevel) {
      button.classList.add("is-active");
    }

    const body = document.createElement("span");
    body.className = "pi-status-popover__item-body";

    const label = document.createElement("span");
    label.className = "pi-status-popover__item-label";
    label.textContent = THINKING_LEVEL_LABELS[level];

    const hint = document.createElement("span");
    hint.className = "pi-status-popover__item-hint";
    hint.textContent = THINKING_LEVEL_HINTS[level];

    body.append(label, hint);

    const marker = document.createElement("span");
    marker.className = "pi-status-popover__item-marker";
    marker.textContent = level === opts.activeLevel ? "✓" : "";

    button.append(body, marker);

    button.addEventListener("click", () => {
      opts.onSelectLevel(level);
      closeStatusPopover();
    });

    list.appendChild(button);
  }

  popover.append(title, description, list);
  mountPopover("thinking", opts.anchor, popover);
}

function createCommandButton(args: {
  command: StatusCommandName;
  title: string;
  description: string;
  onRun: (command: StatusCommandName) => void;
}): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.className = "pi-status-popover__command";

  const command = document.createElement("span");
  command.className = "pi-status-popover__command-name";
  command.textContent = `/${args.command}`;

  const description = document.createElement("span");
  description.className = "pi-status-popover__command-desc";
  description.textContent = args.description;

  button.append(command, description);
  button.setAttribute("aria-label", `${args.title}: ${args.description}`);

  button.addEventListener("click", () => {
    args.onRun(args.command);
    closeStatusPopover();
  });

  return button;
}

export function toggleContextPopover(opts: ContextPopoverOptions): void {
  if (shouldToggle("context", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const popover = createPopoverBase("context");

  const title = document.createElement("h3");
  title.className = "pi-status-popover__title";
  title.textContent = "Context usage";

  const description = createDescriptionBlock(opts.description);

  const actions = document.createElement("div");
  actions.className = "pi-status-popover__commands";

  actions.append(
    createCommandButton({
      command: "compact",
      title: "Compact conversation",
      description: "Summarize earlier messages to free context budget.",
      onRun: opts.onRunCommand,
    }),
    createCommandButton({
      command: "new",
      title: "Start new chat",
      description: "Open a fresh tab with empty context.",
      onRun: opts.onRunCommand,
    }),
  );

  popover.append(title, description, actions);
  mountPopover("context", opts.anchor, popover);
}

// ── Proxy helper popover ──────────────────────────────────────────────

const PROXY_COMMAND = "npx pi-for-excel-proxy";
const INSTALL_GUIDE_URL = "https://pi.dev/excel#connect";

export interface ProxyPopoverOptions {
  anchor: Element;
  onDismiss: () => void;
}

export function toggleProxyPopover(opts: ProxyPopoverOptions): void {
  if (shouldToggle("proxy", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const popover = createPopoverBase("proxy");
  popover.classList.add("pi-proxy-popover");

  const title = document.createElement("div");
  title.className = "pi-proxy-popover__title";
  title.textContent = "Local helper not running";

  const body = document.createElement("div");
  body.className = "pi-proxy-popover__text";
  body.textContent =
    "Some features (web search, sign-in, external services) need a small local helper. Start it with:";

  const cmdRow = document.createElement("div");
  cmdRow.className = "pi-proxy-popover__cmd";

  const code = document.createElement("code");
  code.textContent = PROXY_COMMAND;

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "pi-proxy-popover__copy";
  copyBtn.title = "Copy";
  copyBtn.textContent = "⧉";
  copyBtn.addEventListener("click", () => {
    if (!navigator.clipboard?.writeText) {
      // Fallback: select the text so the user can Cmd+C / Ctrl+C
      const range = document.createRange();
      range.selectNodeContents(code);
      const sel = window.getSelection();
      sel?.removeAllRanges();
      sel?.addRange(range);
      return;
    }
    void navigator.clipboard.writeText(PROXY_COMMAND).then(
      () => {
        copyBtn.textContent = "✓";
        setTimeout(() => { copyBtn.textContent = "⧉"; }, 1500);
      },
      () => {
        // Permission denied — select text as fallback
        const range = document.createRange();
        range.selectNodeContents(code);
        const sel = window.getSelection();
        sel?.removeAllRanges();
        sel?.addRange(range);
      },
    );
  });

  cmdRow.append(code, copyBtn);

  const hint = document.createElement("div");
  hint.className = "pi-proxy-popover__text";
  hint.textContent = "Keep it running while using Pi.";

  const footer = document.createElement("div");
  footer.className = "pi-proxy-popover__footer";

  const guideLink = document.createElement("a");
  guideLink.className = "pi-proxy-popover__link";
  guideLink.href = INSTALL_GUIDE_URL;
  guideLink.target = "_blank";
  guideLink.rel = "noopener noreferrer";
  guideLink.textContent = "No Node.js? See install guide →";

  const dismissBtn = document.createElement("button");
  dismissBtn.type = "button";
  dismissBtn.className = "pi-proxy-popover__dismiss";
  dismissBtn.textContent = "Don't show again";
  dismissBtn.addEventListener("click", () => {
    opts.onDismiss();
    closeStatusPopover();
  });

  footer.append(guideLink, dismissBtn);

  popover.append(title, body, cmdRow, hint, footer);
  mountPopover("proxy", opts.anchor, popover);
}
