/**
 * Safe UI projection primitives for sandboxed extensions.
 *
 * The sandbox can only send a constrained, structured UI tree.
 * Host runtime normalizes and renders that tree without using innerHTML.
 */

import { isRecord } from "../utils/type-guards.js";

export interface SandboxUiTextNode {
  kind: "text";
  text: string;
}

export interface SandboxUiElementNode {
  kind: "element";
  tag: string;
  className?: string;
  actionId?: string;
  children: SandboxUiNode[];
}

export type SandboxUiNode = SandboxUiTextNode | SandboxUiElementNode;

const ALLOWED_TAGS = new Set<string>([
  "div",
  "span",
  "p",
  "strong",
  "em",
  "code",
  "pre",
  "ul",
  "ol",
  "li",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "button",
]);

const CLASS_TOKEN_PATTERN = /^[A-Za-z0-9_-]{1,40}$/u;
const ACTION_ID_PATTERN = /^[A-Za-z0-9:_-]{1,64}$/u;
const MAX_TEXT_LENGTH = 8_000;
const MAX_NODES = 300;
const MAX_DEPTH = 12;
const MAX_CLASS_TOKENS = 8;

function normalizeText(value: unknown): string {
  if (typeof value !== "string") {
    return "";
  }

  if (value.length <= MAX_TEXT_LENGTH) {
    return value;
  }

  return value.slice(0, MAX_TEXT_LENGTH);
}

function normalizeTag(value: unknown): string {
  if (typeof value !== "string") {
    return "div";
  }

  const lowered = value.trim().toLowerCase();
  if (!ALLOWED_TAGS.has(lowered)) {
    return "div";
  }

  return lowered;
}

function normalizeClassName(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const tokens = value
    .split(/\s+/u)
    .map((token) => token.trim())
    .filter((token) => token.length > 0 && CLASS_TOKEN_PATTERN.test(token))
    .slice(0, MAX_CLASS_TOKENS);

  if (tokens.length === 0) {
    return undefined;
  }

  return tokens.join(" ");
}

function normalizeActionId(value: unknown): string | undefined {
  if (typeof value !== "string") {
    return undefined;
  }

  const trimmed = value.trim();
  if (!ACTION_ID_PATTERN.test(trimmed)) {
    return undefined;
  }

  return trimmed;
}

interface NormalizeState {
  remainingNodes: number;
}

function normalizeNode(raw: unknown, depth: number, state: NormalizeState): SandboxUiNode | null {
  if (state.remainingNodes <= 0) {
    return null;
  }

  if (!isRecord(raw)) {
    return null;
  }

  const kind = raw.kind;
  if (kind === "text") {
    state.remainingNodes -= 1;
    return {
      kind: "text",
      text: normalizeText(raw.text),
    };
  }

  if (kind !== "element") {
    return null;
  }

  state.remainingNodes -= 1;

  const tag = normalizeTag(raw.tag);
  const className = normalizeClassName(raw.className);
  const actionId = normalizeActionId(raw.actionId);

  const children: SandboxUiNode[] = [];
  if (depth < MAX_DEPTH && Array.isArray(raw.children)) {
    for (const child of raw.children) {
      const normalizedChild = normalizeNode(child, depth + 1, state);
      if (!normalizedChild) {
        continue;
      }

      children.push(normalizedChild);
      if (state.remainingNodes <= 0) {
        break;
      }
    }
  }

  return {
    kind: "element",
    tag,
    className,
    actionId,
    children,
  };
}

export function normalizeSandboxUiNode(raw: unknown): SandboxUiNode {
  const state: NormalizeState = {
    remainingNodes: MAX_NODES,
  };

  const normalized = normalizeNode(raw, 0, state);
  if (normalized) {
    return normalized;
  }

  return {
    kind: "text",
    text: "",
  };
}

export function collectSandboxUiActionIds(node: SandboxUiNode): string[] {
  const actionIds: string[] = [];

  const visit = (current: SandboxUiNode): void => {
    if (current.kind === "text") {
      return;
    }

    if (current.actionId) {
      actionIds.push(current.actionId);
    }

    for (const child of current.children) {
      visit(child);
    }
  };

  visit(node);
  return actionIds;
}

export function renderSandboxUiNode(
  node: SandboxUiNode,
  onAction: (actionId: string) => void,
): Node {
  if (node.kind === "text") {
    return document.createTextNode(node.text);
  }

  const element = document.createElement(node.tag);
  if (node.className) {
    element.className = node.className;
  }

  if (node.actionId) {
    const actionId = node.actionId;

    element.addEventListener("click", (event) => {
      event.preventDefault();
      onAction(actionId);
    });

    if (node.tag !== "button") {
      element.setAttribute("role", "button");
      element.tabIndex = 0;
      element.addEventListener("keydown", (event: KeyboardEvent) => {
        if (event.key !== "Enter" && event.key !== " ") {
          return;
        }

        event.preventDefault();
        onAction(actionId);
      });
    }
  }

  for (const child of node.children) {
    element.appendChild(renderSandboxUiNode(child, onAction));
  }

  return element;
}

export function renderSandboxUiTree(
  mount: HTMLElement,
  node: SandboxUiNode,
  onAction: (actionId: string) => void,
): void {
  mount.replaceChildren(renderSandboxUiNode(node, onAction));
}
