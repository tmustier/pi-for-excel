/**
 * Patch thinking block labels after completion.
 *
 * Upstream thinking-block currently leaves completed labels as "Thinking…".
 * We patch finished blocks to:
 * - "Thought for Xs" / "Thought for Xm Xs" when timing is available
 * - t("compat.thought") fallback for restored history with no timing data
 */

interface ThinkingBlockState {
  startedAtMs: number | null;
  wasStreaming: boolean;
  completedLabel: string | null;
}

import { t } from "../language/index.js";

const stateByBlock = new WeakMap<HTMLElement, ThinkingBlockState>();

let patchInstalled = false;
let observer: MutationObserver | null = null;
let rafId: number | null = null;

function isHTMLElement(value: unknown): value is HTMLElement {
  return value instanceof HTMLElement;
}

function formatDuration(elapsedMs: number): string {
  const totalSeconds = Math.max(1, Math.round(elapsedMs / 1000));
  const minutes = Math.floor(totalSeconds / 60);
  const seconds = totalSeconds % 60;

  if (minutes <= 0) {
    return `${totalSeconds}s`;
  }

  return `${minutes}m ${seconds}s`;
}

function getHeader(block: HTMLElement): HTMLElement | null {
  const header = block.querySelector(".thinking-header");
  return isHTMLElement(header) ? header : null;
}

function getLabelElement(header: HTMLElement): HTMLElement | null {
  const directSpanChildren = Array.from(header.children).filter(
    (child): child is HTMLElement => isHTMLElement(child) && child.tagName === "SPAN",
  );

  if (directSpanChildren.length === 0) {
    return null;
  }

  // In upstream component, label text is the last direct span child.
  return directSpanChildren[directSpanChildren.length - 1] ?? null;
}

function hasStreamingShimmer(header: HTMLElement): boolean {
  return header.querySelector(".animate-shimmer") !== null;
}

function looksLikeThinkingLabel(text: string): boolean {
  return /^thinking(?:\u2026|\.\.\.)?$/i.test(text.trim());
}

function ensureState(block: HTMLElement): ThinkingBlockState {
  const existing = stateByBlock.get(block);
  if (existing) {
    return existing;
  }

  const next: ThinkingBlockState = {
    startedAtMs: null,
    wasStreaming: false,
    completedLabel: null,
  };
  stateByBlock.set(block, next);
  return next;
}

function syncThinkingBlockLabel(block: HTMLElement): void {
  const header = getHeader(block);
  if (!header) {
    return;
  }

  const labelEl = getLabelElement(header);
  if (!labelEl) {
    return;
  }

  const state = ensureState(block);

  if (hasStreamingShimmer(header)) {
    if (!state.wasStreaming) {
      state.startedAtMs = Date.now();
      state.completedLabel = null;
    }
    state.wasStreaming = true;
    return;
  }

  if (state.wasStreaming) {
    state.wasStreaming = false;
  }

  let nextLabel = state.completedLabel;

  if (nextLabel === null) {
    if (state.startedAtMs !== null) {
      nextLabel = `Thought for ${formatDuration(Date.now() - state.startedAtMs)}`;
      state.completedLabel = nextLabel;
    } else {
      const currentText = labelEl.textContent?.trim() ?? "";
      if (looksLikeThinkingLabel(currentText)) {
        nextLabel = t("compat.thought");
        state.completedLabel = nextLabel;
      }
    }
  }

  if (nextLabel !== null && labelEl.textContent !== nextLabel) {
    labelEl.textContent = nextLabel;
  }
}

function collectThinkingBlocks(root: ParentNode): HTMLElement[] {
  const blocks: HTMLElement[] = [];

  if (isHTMLElement(root) && root.matches("thinking-block")) {
    blocks.push(root);
  }

  const descendants = root.querySelectorAll("thinking-block");
  for (const node of descendants) {
    if (isHTMLElement(node)) {
      blocks.push(node);
    }
  }

  return blocks;
}

function runPass(root: ParentNode): void {
  const blocks = collectThinkingBlocks(root);
  for (const block of blocks) {
    syncThinkingBlockLabel(block);
  }
}

function schedulePass(root: ParentNode): void {
  if (rafId !== null) {
    return;
  }

  rafId = requestAnimationFrame(() => {
    rafId = null;
    runPass(root);
  });
}

function install(root: ParentNode): void {
  runPass(root);

  observer = new MutationObserver(() => {
    schedulePass(root);
  });

  observer.observe(root, {
    subtree: true,
    childList: true,
    attributes: true,
    attributeFilter: ["class"],
  });
}

/**
 * Install the thinking-duration label patch once.
 */
export function installThinkingDurationPatch(): void {
  if (patchInstalled) {
    return;
  }
  patchInstalled = true;

  if (document.body) {
    install(document.body);
    return;
  }

  window.addEventListener(
    "DOMContentLoaded",
    () => {
      const root = document.body ?? document.documentElement;
      install(root);
    },
    { once: true },
  );
}

/** For tests/debug only. */
export function __resetThinkingDurationPatchForTests(): void {
  observer?.disconnect();
  observer = null;

  if (rafId !== null) {
    cancelAnimationFrame(rafId);
    rafId = null;
  }

  patchInstalled = false;
}
