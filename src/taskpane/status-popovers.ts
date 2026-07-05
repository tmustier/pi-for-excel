/**
 * Status-bar popovers (thinking selector + context quick actions).
 */

import type { ThinkingLevel } from "@earendil-works/pi-agent-core";
import { t } from "../language/index.js";

import type { StatusContextWarningSeverity } from "./status-context.js";

export type StatusCommandName = "compact" | "new";

type StatusPopoverKind = "thinking" | "context";

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
  tokenDetail?: string;
  warning?: { text: string; severity: StatusContextWarningSeverity };
  onRunCommand: (command: StatusCommandName) => void;
}

function getThinkingLevelLabels(): Record<ThinkingLevel, string> {
  return {
    off: t("status.thinking.off"),
    minimal: t("status.thinking.min"),
    low: t("status.thinking.low"),
    medium: t("status.thinking.medium"),
    high: t("status.thinking.high"),
    xhigh: t("status.thinking.max"),
  };
}

function getThinkingLevelHints(): Record<ThinkingLevel, string> {
  return {
    off: t("status.thinking.offHint"),
    minimal: t("status.thinking.minimalHint"),
    low: t("status.thinking.lowHint"),
    medium: t("status.thinking.mediumHint"),
    high: t("status.thinking.highHint"),
    xhigh: t("status.thinking.xhighHint"),
  };
}

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
  title.textContent = t("status-popovers.thinkingLevel");

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
    label.textContent = getThinkingLevelLabels()[level];

    const hint = document.createElement("span");
    hint.className = "pi-status-popover__item-hint";
    hint.textContent = getThinkingLevelHints()[level];

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
  title.textContent = t("status-popovers.contextUsage");

  const description = createDescriptionBlock(opts.description);

  if (opts.tokenDetail) {
    const detail = document.createElement("p");
    detail.className = "pi-status-popover__token-detail";
    detail.textContent = opts.tokenDetail;
    popover.append(title, description, detail);
  } else {
    popover.append(title, description);
  }

  if (opts.warning) {
    const warn = document.createElement("p");
    warn.className = `pi-status-popover__warning pi-status-popover__warning--${opts.warning.severity}`;
    warn.textContent = opts.warning.text;
    popover.appendChild(warn);
  }

  const actions = document.createElement("div");
  actions.className = "pi-status-popover__commands";

  actions.append(
    createCommandButton({
      command: "compact",
      title: t("status-popovers.compactTitle"),
      description: t("status-popovers.compactDesc"),
      onRun: opts.onRunCommand,
    }),
    createCommandButton({
      command: "new",
      title: t("status-popovers.newTitle"),
      description: t("status-popovers.newDesc"),
      onRun: opts.onRunCommand,
    }),
  );

  popover.appendChild(actions);
  mountPopover("context", opts.anchor, popover);
}

