/**
 * Builtin command overlays (provider picker, resume, shortcuts).
 */

import type { Agent } from "@mariozechner/pi-agent-core";
import { getAppStorage } from "@mariozechner/pi-web-ui";

import { showToast } from "../../ui/toast.js";
import type { PiSidebar } from "../../ui/pi-sidebar.js";

export async function showProviderPicker(): Promise<void> {
  const existing = document.getElementById("pi-login-overlay");
  if (existing) {
    existing.remove();
    return;
  }

  const { ALL_PROVIDERS, buildProviderRow } = await import("../../ui/provider-login.js");
  const storage = getAppStorage();
  const configuredKeys = await storage.providerKeys.list();
  const configuredSet = new Set(configuredKeys);

  const overlay = document.createElement("div");
  overlay.id = "pi-login-overlay";
  overlay.className = "pi-welcome-overlay";

  overlay.innerHTML = `
    <div class="pi-welcome-card" style="text-align: left; max-width: 340px;">
      <h2 style="font-size: 16px; font-weight: 600; margin: 0 0 4px; font-family: var(--font-sans);">Providers</h2>
      <p style="font-size: 12px; color: var(--muted-foreground); margin: 0 0 12px; font-family: var(--font-sans);">Connect providers to use their models.</p>
      <div class="pi-login-providers" style="display: flex; flex-direction: column; gap: 4px;"></div>
    </div>
  `;

  const list = overlay.querySelector<HTMLDivElement>(".pi-login-providers");
  if (!list) {
    throw new Error("Provider list container not found");
  }

  const expandedRef = { current: null as HTMLElement | null };

  for (const provider of ALL_PROVIDERS) {
    const isActive = configuredSet.has(provider.id);
    const row = buildProviderRow(provider, {
      isActive,
      expandedRef,
      onConnected: (_row: HTMLElement, _id: string, label: string) => {
        document.dispatchEvent(new CustomEvent("pi:providers-changed"));
        showToast(`${label} connected`);
      },
    });
    list.appendChild(row);
  }

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });

  document.body.appendChild(overlay);
}

export async function showResumeDialog(agent: Agent): Promise<void> {
  const storage = getAppStorage();
  const sessions = await storage.sessions.getAllMetadata();

  if (sessions.length === 0) {
    showToast("No previous sessions");
    return;
  }

  const existing = document.getElementById("pi-resume-overlay");
  if (existing) {
    existing.remove();
    return;
  }

  const overlay = document.createElement("div");
  overlay.id = "pi-resume-overlay";
  overlay.className = "pi-welcome-overlay";

  const formatDate = (iso: string) => {
    const d = new Date(iso);
    const now = new Date();
    const diff = now.getTime() - d.getTime();
    if (diff < 60000) return "just now";
    if (diff < 3600000) return `${Math.round(diff / 60000)}m ago`;
    if (diff < 86400000) return `${Math.round(diff / 3600000)}h ago`;
    if (diff < 604800000) return `${Math.round(diff / 86400000)}d ago`;
    return d.toLocaleDateString();
  };

  overlay.innerHTML = `
    <div class="pi-welcome-card" style="text-align: left; max-height: 80vh; overflow: hidden; display: flex; flex-direction: column;">
      <h2 style="font-size: 16px; font-weight: 600; margin: 0 0 12px; font-family: var(--font-sans); flex-shrink: 0;">Resume Session</h2>
      <div class="pi-resume-list" style="overflow-y: auto; display: flex; flex-direction: column; gap: 4px;">
        ${sessions
          .slice(0, 20)
          .map(
            (s) => `
          <button class="pi-welcome-provider pi-resume-item" data-id="${s.id}" style="display: flex; flex-direction: column; align-items: flex-start; gap: 2px;">
            <span style="font-size: 13px; font-weight: 500;">${s.title || "Untitled"}</span>
            <span style="font-size: 11px; color: var(--muted-foreground);">${s.messageCount || 0} messages · ${formatDate(s.lastModified)}</span>
          </button>
        `,
          )
          .join("")}
      </div>
    </div>
  `;

  overlay.addEventListener("click", async (e) => {
    if (e.target === overlay) {
      overlay.remove();
      return;
    }

    const item = (e.target as HTMLElement).closest(
      ".pi-resume-item",
    ) as HTMLElement | null;
    if (!item) return;

    const id = item.dataset.id;
    if (!id) return;

    const sessionData = await storage.sessions.loadSession(id);
    if (!sessionData) {
      showToast("Session not found");
      overlay.remove();
      return;
    }

    // Restore messages and model
    agent.replaceMessages(sessionData.messages || []);
    if (sessionData.model) {
      agent.setModel(sessionData.model);
    }
    if (sessionData.thinkingLevel) {
      agent.setThinkingLevel(sessionData.thinkingLevel);
    }

    // Notify session tracker of the resumed session
    document.dispatchEvent(
      new CustomEvent("pi:session-resumed", {
        detail: {
          id: sessionData.id,
          title: sessionData.title,
          createdAt: sessionData.createdAt,
        },
      }),
    );

    // Force UI to re-render + hide empty state
    const iface = document.querySelector<PiSidebar>("pi-sidebar");
    iface?.requestUpdate();
    document.dispatchEvent(new CustomEvent("pi:model-changed"));

    overlay.remove();
    showToast(`Resumed: ${sessionData.title || "Untitled"}`);
  });

  document.body.appendChild(overlay);
}

export function showShortcutsDialog(): void {
  const shortcuts = [
    ["Enter", "Send message"],
    ["Shift+Tab", "Cycle thinking level"],
    ["Esc", "Abort agent / dismiss menu"],
    ["Enter (streaming)", "Steer — redirect agent"],
    ["⌥Enter", "Queue follow-up message"],
    ["/", "Open command menu"],
    ["↑↓", "Navigate command menu"],
    ["F6", "Focus: Sheet ↔ Sidebar"],
    ["⇧F6", "Focus: reverse direction"],
  ];

  const existing = document.getElementById("pi-shortcuts-overlay");
  if (existing) {
    existing.remove();
    return;
  }

  const overlay = document.createElement("div");
  overlay.id = "pi-shortcuts-overlay";
  overlay.className = "pi-welcome-overlay";
  overlay.innerHTML = `
    <div class="pi-welcome-card" style="text-align: left;">
      <h2 style="font-size: 16px; font-weight: 600; margin: 0 0 12px; font-family: var(--font-sans);">Keyboard Shortcuts</h2>
      <div style="display: flex; flex-direction: column; gap: 6px;">
        ${shortcuts
          .map(
            ([key, desc]) => `
          <div style="display: flex; justify-content: space-between; align-items: center; gap: 12px;">
            <kbd style="font-family: var(--font-mono); font-size: 11px; padding: 2px 6px; background: oklch(0 0 0 / 0.05); border-radius: 4px; white-space: nowrap;">${key}</kbd>
            <span style="font-size: 12.5px; color: var(--muted-foreground); font-family: var(--font-sans);">${desc}</span>
          </div>
        `,
          )
          .join("")}
      </div>
      <button onclick="this.closest('.pi-welcome-overlay').remove()" style="margin-top: 16px; width: 100%; padding: 8px; border-radius: 8px; border: 1px solid oklch(0 0 0 / 0.08); background: oklch(0 0 0 / 0.03); cursor: pointer; font-family: var(--font-sans); font-size: 13px;">Close</button>
    </div>
  `;

  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) overlay.remove();
  });

  document.body.appendChild(overlay);
}
