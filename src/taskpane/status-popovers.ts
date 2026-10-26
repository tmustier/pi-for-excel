/**
 * Status-bar popovers (thinking selector + context quick actions).
 */

import type { ThinkingLevel } from "@earendil-works/pi-agent-core";
import type { ExecutionMode } from "../execution/mode.js";
import { t } from "../language/index.js";
import { focusInitialMenuItem, handleMenuKeydown } from "../ui/menu-keyboard.js";

import type { StatusContextWarningSeverity } from "./status-context.js";
import { getThinkingLevelHint, getThinkingLevelLabel } from "./thinking-display.js";

export type StatusCommandName = "compact" | "new";

type StatusPopoverKind = "thinking" | "context" | "mode";

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

interface ModePopoverOptions {
  anchor: Element;
  activeMode: ExecutionMode;
  onSelectMode: (mode: ExecutionMode) => void;
}

/** Status-bar controls are re-rendered often; find the live trigger by kind. */
const ANCHOR_SELECTORS: Record<StatusPopoverKind, string> = {
  thinking: ".pi-status-thinking",
  context: ".pi-status-ctx--trigger",
  mode: ".pi-status-mode",
};

let activePopover: ActivePopoverState | null = null;
let popoverSequence = 0;

function clamp(value: number, min: number, max: number): number {
  if (max < min) return min;
  return Math.min(Math.max(value, min), max);
}

interface StatusPopoverLayoutInput {
  anchor: Pick<DOMRect, "bottom" | "right" | "top">;
  viewportWidth: number;
  viewportHeight: number;
  popoverWidth: number;
  popoverHeight: number;
}

export interface StatusPopoverLayout {
  left: number;
  top: number;
  maxHeight: number;
}

export function resolveStatusPopoverLayout(input: StatusPopoverLayoutInput): StatusPopoverLayout {
  const inset = 8;
  const maxHeight = Math.max(0, input.viewportHeight - inset * 2);
  const boundedHeight = Math.min(input.popoverHeight, maxHeight);
  const maxLeft = Math.max(inset, input.viewportWidth - input.popoverWidth - inset);
  const left = clamp(input.anchor.right - input.popoverWidth, inset, maxLeft);

  const preferredTop = input.anchor.top - boundedHeight - inset;
  const fallbackTop = input.anchor.bottom + inset;
  const maxTop = Math.max(inset, input.viewportHeight - boundedHeight - inset);
  const top = preferredTop >= inset
    ? preferredTop
    : clamp(fallbackTop, inset, maxTop);

  return { left, top, maxHeight };
}

function normalizeDescription(text: string): string {
  const compact = text.replace(/\s+/g, " ").trim();
  return compact.length > 0 ? compact : t("status-popovers.noDetails");
}

function createPopoverBase(kind: StatusPopoverKind, titleText: string): { popover: HTMLDivElement; title: HTMLHeadingElement } {
  popoverSequence += 1;
  const popover = document.createElement("div");
  popover.className = `pi-status-popover pi-status-popover--${kind}`;
  popover.setAttribute("role", "dialog");

  const title = document.createElement("h3");
  title.className = "pi-status-popover__title";
  title.id = `pi-status-popover-title-${popoverSequence}`;
  title.textContent = titleText;
  popover.setAttribute("aria-labelledby", title.id);
  return { popover, title };
}

function createMenuList(labelledBy: string): HTMLDivElement {
  const list = document.createElement("div");
  list.className = "pi-status-popover__list";
  list.setAttribute("role", "menu");
  list.setAttribute("aria-labelledby", labelledBy);
  return list;
}

function createChoiceItem(args: {
  label: string;
  hint: string;
  checked: boolean;
  onSelect: () => void;
}): HTMLButtonElement {
  const button = document.createElement("button");
  button.type = "button";
  button.tabIndex = -1;
  button.className = "pi-status-popover__item";
  button.setAttribute("role", "menuitemradio");
  button.setAttribute("aria-checked", args.checked ? "true" : "false");
  if (args.checked) {
    button.classList.add("is-active");
  }

  const body = document.createElement("span");
  body.className = "pi-status-popover__item-body";

  const label = document.createElement("span");
  label.className = "pi-status-popover__item-label";
  label.textContent = args.label;

  const hint = document.createElement("span");
  hint.className = "pi-status-popover__item-hint";
  hint.textContent = args.hint;

  body.append(label, hint);

  const marker = document.createElement("span");
  marker.className = "pi-status-popover__item-marker";
  marker.setAttribute("aria-hidden", "true");
  marker.textContent = args.checked ? "✓" : "";

  button.append(body, marker);
  button.addEventListener("click", () => {
    args.onSelect();
    closeStatusPopover({ restoreFocus: true });
  });
  return button;
}

function positionPopover(popover: HTMLDivElement, anchor: Element): void {
  const anchorRect = anchor.getBoundingClientRect();
  const viewportWidth = document.documentElement.clientWidth;
  const viewportHeight = document.documentElement.clientHeight;

  popover.style.left = "0px";
  popover.style.top = "0px";
  popover.style.visibility = "hidden";
  popover.style.maxHeight = `${Math.max(0, viewportHeight - 16)}px`;

  // Use fractional sizes: offsetWidth/offsetHeight round, which can push the
  // popover's edge past the inset by a sub-pixel amount.
  const popoverRect = popover.getBoundingClientRect();
  const layout = resolveStatusPopoverLayout({
    anchor: anchorRect,
    viewportWidth,
    viewportHeight,
    popoverWidth: popoverRect.width,
    popoverHeight: popoverRect.height,
  });

  popover.style.left = `${layout.left}px`;
  popover.style.top = `${layout.top}px`;
  popover.style.maxHeight = `${layout.maxHeight}px`;
  popover.style.visibility = "visible";
}

function closeIfOpen(state: ActivePopoverState | null): void {
  if (!state) return;
  state.cleanup();
  state.popover.remove();
}

/** Close the open status popover; with `restoreFocus`, focus returns to its trigger. */
export function closeStatusPopover(opts: { restoreFocus?: boolean } = {}): void {
  const state = activePopover;
  activePopover = null;
  closeIfOpen(state);
  if (!state || !opts.restoreFocus) return;

  const anchor = state.anchor.isConnected
    ? state.anchor
    : document.querySelector(ANCHOR_SELECTORS[state.kind]);
  if (anchor instanceof HTMLElement) {
    anchor.focus({ preventScroll: true });
  }
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

  // Escape closes from anywhere (the trigger may still hold focus on hosts that
  // do not focus buttons on click); the menu keys apply inside the popover.
  const onKeyDown = (event: KeyboardEvent) => {
    if (event.key !== "Escape") return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    closeStatusPopover({ restoreFocus: true });
  };

  const onPopoverKeyDown = (event: KeyboardEvent) => {
    if (event.key === "Escape") return;
    handleMenuKeydown(event, popover, () => closeStatusPopover({ restoreFocus: true }));
  };

  popover.addEventListener("keydown", onPopoverKeyDown);
  document.addEventListener("mousedown", onMouseDown, true);
  window.addEventListener("keydown", onKeyDown, true);
  window.addEventListener("resize", reposition);
  window.addEventListener("scroll", reposition, true);

  activePopover = {
    kind,
    anchor,
    popover,
    cleanup: () => {
      popover.removeEventListener("keydown", onPopoverKeyDown);
      document.removeEventListener("mousedown", onMouseDown, true);
      window.removeEventListener("keydown", onKeyDown, true);
      window.removeEventListener("resize", reposition);
      window.removeEventListener("scroll", reposition, true);
    },
  };

  focusInitialMenuItem(popover);
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

  const { popover, title } = createPopoverBase("thinking", t("status-popovers.thinkingLevel"));
  const description = createDescriptionBlock(opts.description);
  const list = createMenuList(title.id);

  for (const level of opts.levels) {
    list.appendChild(createChoiceItem({
      label: getThinkingLevelLabel(level),
      hint: getThinkingLevelHint(level),
      checked: level === opts.activeLevel,
      onSelect: () => opts.onSelectLevel(level),
    }));
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
  button.tabIndex = -1;
  button.className = "pi-status-popover__command";
  button.setAttribute("role", "menuitem");

  const command = document.createElement("span");
  command.className = "pi-status-popover__command-name";
  command.textContent = `/${args.command}`;

  const description = document.createElement("span");
  description.className = "pi-status-popover__command-desc";
  description.textContent = args.description;

  button.append(command, description);
  button.setAttribute("aria-label", `${args.title}: ${args.description}`);

  button.addEventListener("click", () => {
    closeStatusPopover({ restoreFocus: true });
    args.onRun(args.command);
  });

  return button;
}

export function toggleContextPopover(opts: ContextPopoverOptions): void {
  if (shouldToggle("context", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const { popover, title } = createPopoverBase("context", t("status-popovers.contextUsage"));

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

  const actions = createMenuList(title.id);
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


export function toggleModePopover(opts: ModePopoverOptions): void {
  if (shouldToggle("mode", opts.anchor)) {
    closeStatusPopover();
    return;
  }

  const { popover, title } = createPopoverBase("mode", t("status-popovers.mode.title"));
  const list = createMenuList(title.id);
  const modes: ReadonlyArray<{ mode: ExecutionMode; label: string; hint: string }> = [
    { mode: "yolo", label: t("status.mode.auto"), hint: t("status-popovers.mode.autoHint") },
    { mode: "safe", label: t("status.mode.confirm"), hint: t("status-popovers.mode.confirmHint") },
  ];

  for (const entry of modes) {
    list.appendChild(createChoiceItem({
      label: entry.label,
      hint: entry.hint,
      checked: entry.mode === opts.activeMode,
      onSelect: () => opts.onSelectMode(entry.mode),
    }));
  }

  popover.append(title, list);
  mountPopover("mode", opts.anchor, popover);
}
