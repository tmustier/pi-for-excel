/**
 * Tool card grouping — wraps consecutive same-tool calls in a single
 * collapsible container. Whether a group is open comes from
 * transcript-state.ts: groups of 3+ start collapsed unless details are
 * expanded (Ctrl/Cmd+O), and only a group the user toggles keeps its own
 * state, keyed by its first call.
 */

import { t } from "../language/index.js";
import {
  isToolGroupExpanded,
  setDisclosureExpanded,
  subscribeDisclosure,
} from "./messages/transcript-state.js";

/* ── Group header summary ──────────────────────────────── */

/** Tool name → locale key for the group header ("3 reads"). */
const TOOL_GROUP_LABEL_KEYS: Record<string, string> = {
  fill_formula: "tool-group.fill_formula",
  write_cells: "tool-group.write_cells",
  read_range: "tool-group.read_range",
  format_cells: "tool-group.format_cells",
  conditional_format: "tool-group.conditional_format",
  view_settings: "tool-group.view_settings",
  charts: "tool-group.charts",
  execute_office_js: "tool-group.script",
  execute_wps_js: "tool-group.script",
};

function describeGroup(toolName: string, count: number): string {
  const key = TOOL_GROUP_LABEL_KEYS[toolName];
  return key ? t(key, { count }) : t("tool-group.other", { count, tool: toolName });
}

function buildGroupHeader(toolName: string, count: number, groupKey: string, expanded: boolean): HTMLButtonElement {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "pi-tool-group__header";
  btn.setAttribute("aria-expanded", expanded ? "true" : "false");

  const chevron = document.createElement("span");
  chevron.className = "pi-tool-group__chevron";
  chevron.textContent = "▸";

  const label = document.createElement("span");
  label.className = "pi-tool-group__label";
  label.textContent = describeGroup(toolName, count);

  btn.append(chevron, label);

  // Regrouping rebuilds the header from transcript state.
  btn.addEventListener("click", () => setDisclosureExpanded(groupKey, !expanded));

  return btn;
}

/* ── Main entry point ──────────────────────────────────── */

/**
 * Initialise tool-card grouping on the given root element.
 * Returns a cleanup function that disconnects the observer and removes
 * all grouping artefacts.
 */
export function initToolGrouping(root: HTMLElement): () => void {
  let rafId = 0;

  /* ── Unwrap existing groups ────────────────────────────── */

  function unwrapAll() {
    for (const wrapper of root.querySelectorAll(".pi-tool-group")) {
      const parent = wrapper.parentNode;
      if (!parent) continue;

      // Remove injected header before unwrapping.
      wrapper.querySelector(".pi-tool-group__header")?.remove();
      while (wrapper.firstChild) parent.insertBefore(wrapper.firstChild, wrapper);
      parent.removeChild(wrapper);
    }
  }

  /* ── Grouping pass ────────────────────────────────────── */

  function applyGrouping() {
    // Disconnect observer during DOM manipulation to avoid re-entrancy.
    observer.disconnect();

    // Flatten — move all tool-messages back to root.
    unwrapAll();

    // Clean up classes on all tool-messages.
    const toolMessages: Element[] = [];
    for (const el of root.querySelectorAll("tool-message")) {
      el.classList.remove("pi-group-member");
      toolMessages.push(el);
    }

    // Identify runs of 2+ consecutive same-name completed tools.
    const runs: { elements: Element[]; toolName: string }[] = [];
    let currentRun: Element[] = [];
    let currentToolName = "";

    for (const el of toolMessages) {
      const card = el.querySelector(".pi-tool-card");
      if (!card) {
        if (currentRun.length >= 2) runs.push({ elements: currentRun, toolName: currentToolName });
        currentRun = [];
        currentToolName = "";
        continue;
      }

      const toolName = card.getAttribute("data-tool-name") ?? "";
      const cardState = card.getAttribute("data-state");

      if (cardState !== "complete" || !toolName) {
        if (currentRun.length >= 2) runs.push({ elements: currentRun, toolName: currentToolName });
        currentRun = [];
        currentToolName = "";
        continue;
      }

      if (currentRun.length > 0) {
        const prev = currentRun[currentRun.length - 1];
        if (!prev) {
          currentRun = [el];
          currentToolName = toolName;
          continue;
        }

        const prevCard = prev.querySelector(".pi-tool-card");
        const prevName = prevCard?.getAttribute("data-tool-name");

        if (prevName === toolName && areConsecutiveSiblings(prev, el)) {
          currentRun.push(el);
        } else {
          if (currentRun.length >= 2) runs.push({ elements: currentRun, toolName: currentToolName });
          currentRun = [el];
          currentToolName = toolName;
        }
      } else {
        currentRun.push(el);
        currentToolName = toolName;
      }
    }
    if (currentRun.length >= 2) runs.push({ elements: currentRun, toolName: currentToolName });

    // Wrap each run in a container element with a collapsible header.
    for (const run of runs) {
      const leader = run.elements[0];
      if (!leader) {
        continue;
      }

      const members = run.elements.slice(1);
      const count = run.elements.length;

      const groupKey = leader.getAttribute("data-group-key") ?? "";
      const expanded = isToolGroupExpanded(groupKey, count);

      const wrapper = document.createElement("div");
      wrapper.className = "pi-tool-group" + (expanded ? "" : " pi-tool-group--collapsed");

      // Inject group header with summary + expand/collapse toggle.
      const header = buildGroupHeader(run.toolName, count, groupKey, expanded);
      wrapper.appendChild(header);

      if (leader.parentNode) leader.parentNode.insertBefore(wrapper, leader);
      for (const el of run.elements) wrapper.appendChild(el);

      for (const m of members) m.classList.add("pi-group-member");
    }

    // Reconnect observer after all DOM work is done.
    observer.observe(root, { childList: true, subtree: true });
  }

  /**
   * Check whether two elements are consecutive siblings (no intervening
   * element siblings — only whitespace text nodes allowed).
   */
  function areConsecutiveSiblings(a: Element, b: Element): boolean {
    let node: Node | null = a.nextSibling;
    while (node) {
      if (node === b) return true;
      if (node.nodeType === Node.TEXT_NODE) {
        node = node.nextSibling;
        continue;
      }
      if (node.nodeType === Node.ELEMENT_NODE) return false;
      node = node.nextSibling;
    }
    return false;
  }

  /* ── Observer ──────────────────────────────────────────── */

  function scheduleGrouping() {
    if (rafId) return;
    rafId = requestAnimationFrame(() => {
      rafId = 0;
      applyGrouping();
    });
  }

  const observer = new MutationObserver(scheduleGrouping);
  observer.observe(root, { childList: true, subtree: true });
  const unsubscribeDisclosure = subscribeDisclosure(scheduleGrouping);

  // Initial pass.
  applyGrouping();

  /* ── Cleanup ──────────────────────────────────────────── */

  return () => {
    observer.disconnect();
    unsubscribeDisclosure();
    if (rafId) cancelAnimationFrame(rafId);

    unwrapAll();
    for (const el of root.querySelectorAll("tool-message")) {
      el.classList.remove("pi-group-member");
    }
  };
}
