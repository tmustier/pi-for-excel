/**
 * Small DOM-only queue display for queued steering / follow-up messages.
 *
 * This intentionally stays as plain DOM manipulation (not Lit) for now.
 */

import type { Agent } from "@mariozechner/pi-agent-core";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { extractTextFromContent } from "../utils/content.js";

export type QueuedMessageType = "steer" | "follow-up";

export function createQueueDisplay(opts: {
  agent: Agent;
  sidebar: PiSidebar;
}): {
  add: (type: QueuedMessageType, text: string) => void;
  clear: () => void;
} {
  const { agent, sidebar } = opts;

  type QueuedItem = { type: QueuedMessageType; text: string };
  const queued: QueuedItem[] = [];

  function updateQueueDisplay() {
    let container = document.getElementById("pi-queue-display");
    if (queued.length === 0) {
      container?.remove();
      return;
    }
    if (!container) {
      container = document.createElement("div");
      container.id = "pi-queue-display";
      container.className = "pi-queue";
      document.body.appendChild(container);
    }

    // Position above the working indicator (or input area if indicator hidden)
    const workingEl = sidebar.querySelector("pi-working-indicator") as HTMLElement | null;
    const inputArea = sidebar.querySelector(".pi-input-area") as HTMLElement | null;
    const anchorEl = workingEl && workingEl.offsetHeight > 0 ? workingEl : inputArea;
    const anchorTop = anchorEl ? anchorEl.getBoundingClientRect().top : window.innerHeight - 80;
    container.style.bottom = `${window.innerHeight - anchorTop}px`;

    container.innerHTML = queued
      .map(({ type, text }) => {
        const label = type === "steer" ? "Steering" : "Follow-up";
        const cls = type === "steer" ? "pi-queue__label--steer" : "pi-queue__label--followup";
        const truncated = text.length > 50 ? text.slice(0, 47) + "…" : text;
        return `<div class="pi-queue__item">
        <span class="pi-queue__label ${cls}">${label}</span>
        <span class="pi-queue__text">${truncated}</span>
      </div>`;
      })
      .join("");
  }

  function add(type: QueuedMessageType, text: string) {
    queued.push({ type, text });
    updateQueueDisplay();
  }

  function clear() {
    queued.length = 0;
    updateQueueDisplay();
  }

  agent.subscribe((ev) => {
    if (queued.length === 0) return;

    if (ev.type === "message_start" && ev.message.role === "user") {
      const msgText = extractTextFromContent(ev.message.content);
      const idx = queued.findIndex((q) => q.text === msgText);
      if (idx !== -1) {
        queued.splice(idx, 1);
        updateQueueDisplay();
      }
    }

    if (ev.type === "agent_end" && queued.length > 0) clear();
  });

  return { add, clear };
}
