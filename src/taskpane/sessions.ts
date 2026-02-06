/**
 * Session persistence wiring for the taskpane.
 *
 * Owns:
 * - auto-saving agent state to IndexedDB
 * - restoring latest session on startup
 * - keeping internal session id/title in sync with /new, /name, /resume events
 */

import type { Agent } from "@mariozechner/pi-agent-core";
import type { SessionsStore } from "@mariozechner/pi-web-ui";

import type { PiSidebar } from "../ui/pi-sidebar.js";
import { extractTextFromContent } from "../utils/content.js";

export async function setupSessionPersistence(opts: {
  agent: Agent;
  sidebar: PiSidebar;
  sessions: SessionsStore;
}): Promise<void> {
  const { agent, sidebar, sessions } = opts;

  let sessionId: string = crypto.randomUUID();
  let sessionTitle = "";
  let sessionCreatedAt = new Date().toISOString();
  let firstAssistantSeen = false;

  async function saveSession() {
    if (!firstAssistantSeen) return;

    try {
      const now = new Date().toISOString();
      const messages = agent.state.messages;

      if (!sessionTitle && messages.length > 0) {
        const firstUser = messages.find((m) => m.role === "user");
        if (firstUser) {
          const text = extractTextFromContent(firstUser.content);
          sessionTitle = text.slice(0, 80) || "Untitled";
        }
      }

      let preview = "";
      for (const m of messages) {
        if (m.role !== "user" && m.role !== "assistant") continue;
        const text = extractTextFromContent(m.content);
        preview += text + "\n";
        if (preview.length > 2048) {
          preview = preview.slice(0, 2048);
          break;
        }
      }

      let inputTokens = 0;
      let outputTokens = 0;
      let cacheReadTokens = 0;
      let cacheWriteTokens = 0;
      let totalTokens = 0;

      let costInput = 0;
      let costOutput = 0;
      let costCacheRead = 0;
      let costCacheWrite = 0;
      let costTotal = 0;

      for (const m of messages) {
        if (m.role !== "assistant") continue;
        const u = m.usage;
        inputTokens += u.input;
        outputTokens += u.output;
        cacheReadTokens += u.cacheRead;
        cacheWriteTokens += u.cacheWrite;
        totalTokens += u.totalTokens;

        costInput += u.cost.input;
        costOutput += u.cost.output;
        costCacheRead += u.cost.cacheRead;
        costCacheWrite += u.cost.cacheWrite;
        costTotal += u.cost.total;
      }

      await sessions.saveSession(
        sessionId,
        agent.state,
        {
          id: sessionId,
          title: sessionTitle,
          createdAt: sessionCreatedAt,
          lastModified: now,
          messageCount: messages.length,
          usage: {
            input: inputTokens,
            output: outputTokens,
            cacheRead: cacheReadTokens,
            cacheWrite: cacheWriteTokens,
            totalTokens,
            cost: {
              input: costInput,
              output: costOutput,
              cacheRead: costCacheRead,
              cacheWrite: costCacheWrite,
              total: costTotal,
            },
          },
          thinkingLevel: agent.state.thinkingLevel || "off",
          preview,
        },
        sessionTitle,
      );
    } catch (err) {
      console.warn("[pi] Session save failed:", err);
    }
  }

  function startNewSession() {
    sessionId = crypto.randomUUID();
    sessionTitle = "";
    sessionCreatedAt = new Date().toISOString();
    firstAssistantSeen = false;
  }

  agent.subscribe((ev) => {
    if (ev.type === "message_end") {
      if (ev.message.role === "assistant") firstAssistantSeen = true;
      if (firstAssistantSeen) saveSession();
    }
  });

  // Auto-restore latest session
  try {
    const latestId = await sessions.getLatestSessionId();
    if (latestId) {
      const sessionData = await sessions.loadSession(latestId);
      if (sessionData && sessionData.messages.length > 0) {
        sessionId = sessionData.id;
        sessionTitle = sessionData.title || "";
        sessionCreatedAt = sessionData.createdAt;
        firstAssistantSeen = true;
        agent.replaceMessages(sessionData.messages);
        if (sessionData.model) {
          agent.setModel(sessionData.model);
        }
        if (sessionData.thinkingLevel) {
          agent.setThinkingLevel(sessionData.thinkingLevel);
        }
        // Force sidebar to pick up restored messages
        sidebar.syncFromAgent();
        console.log(`[pi] Restored session: ${sessionTitle || latestId}`);
      }
    }
  } catch (err) {
    console.warn("[pi] Session restore failed:", err);
  }

  document.addEventListener("pi:session-new", () => startNewSession());
  document.addEventListener(
    "pi:session-rename",
    ((e: CustomEvent) => {
      sessionTitle = e.detail?.title || sessionTitle;
      saveSession();
    }) as EventListener,
  );
  document.addEventListener(
    "pi:session-resumed",
    ((e: CustomEvent) => {
      sessionId = e.detail?.id || sessionId;
      sessionTitle = e.detail?.title || "";
      sessionCreatedAt = e.detail?.createdAt || new Date().toISOString();
      firstAssistantSeen = true;
    }) as EventListener,
  );
}
