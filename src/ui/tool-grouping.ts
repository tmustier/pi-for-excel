/**
 * Tool card grouping — wraps consecutive same-tool calls in a single
 * collapsible container. Groups of 3+ start collapsed; groups of 2 start
 * expanded (grouping with stripped card chrome only).
 */

/* ── Group header summary ──────────────────────────────── */

/** Map tool-name → human-readable plural label for the group header. */
const TOOL_GROUP_LABELS: Record<string, string> = {
  fill_formula: "fill operations",
  write_cells: "edits",
  read_range: "reads",
  format_cells: "format operations",
  conditional_format: "conditional formats",
  view_settings: "view changes",
  chart: "chart operations",
  execute_office_js: "script executions",
};

function describeGroup(toolName: string, count: number): string {
  const label = TOOL_GROUP_LABELS[toolName] ?? `${toolName} calls`;
  return `${count} ${label}`;
}

function buildGroupHeader(toolName: string, count: number, collapsed: boolean): HTMLButtonElement {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "pi-tool-group__header";
  btn.setAttribute("aria-expanded", collapsed ? "false" : "true");

  const chevron = document.createElement("span");
  chevron.className = "pi-tool-group__chevron";
  chevron.textContent = "▸";

  const label = document.createElement("span");
  label.className = "pi-tool-group__label";
  label.textContent = describeGroup(toolName, count);

  btn.append(chevron, label);

  btn.addEventListener("click", () => {
    const wrapper = btn.parentElement;
    if (!wrapper) return;
    const isCollapsed = wrapper.classList.toggle("pi-tool-group--collapsed");
    btn.setAttribute("aria-expanded", isCollapsed ? "false" : "true");
  });

  return btn;
}

/* ── Collapse threshold ──────────────────────────────────── */

/** Groups with this many or more items start collapsed. */
const COLLAPSE_THRESHOLD = 3;

/* ── Main entry point ──────────────────────────────────── */

/**
 * Initialise tool-card grouping on the given root element.
 * Returns a cleanup function that disconnects the observer and removes
 * all grouping artefacts.
 */
export function initToolGrouping(root: HTMLElement): () => void {
  let rafId = 0;

  /**
   * Preserve user toggle state across regrouping passes. Keyed on the
   * first tool-message element of each group — if the user explicitly
   * expanded or collapsed a group, we restore that state when the group
   * is rebuilt rather than using the default.
   */
  const userToggleState = new WeakMap<Element, boolean>();

  /* ── Unwrap existing groups ────────────────────────────── */

  function unwrapAll() {
    for (const wrapper of root.querySelectorAll(".pi-tool-group")) {
      const parent = wrapper.parentNode;
      if (!parent) continue;

      // Snapshot user toggle state before unwrapping.
      const firstMessage = wrapper.querySelector("tool-message");
      if (firstMessage) {
        userToggleState.set(
          firstMessage,
          wrapper.classList.contains("pi-tool-group--collapsed"),
        );
      }

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
      const members = run.elements.slice(1);
      const count = run.elements.length;

      // Restore user toggle state if available, otherwise use default.
      const savedState = userToggleState.get(leader);
      const collapsed = savedState ?? count >= COLLAPSE_THRESHOLD;

      const wrapper = document.createElement("div");
      wrapper.className = "pi-tool-group" + (collapsed ? " pi-tool-group--collapsed" : "");

      // Inject group header with summary + expand/collapse toggle.
      const header = buildGroupHeader(run.toolName, count, collapsed);
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

  // Initial pass.
  applyGrouping();

  /* ── Cleanup ──────────────────────────────────────────── */

  return () => {
    observer.disconnect();
    if (rafId) cancelAnimationFrame(rafId);

    unwrapAll();
    for (const el of root.querySelectorAll("tool-message")) {
      el.classList.remove("pi-group-member");
    }
  };
}
