/**
 * Auto-context injection.
 *
 * Adds (when available):
 * - workbook structure context (blueprint), only on first send and when invalidated
 * - workspace files summary
 * - selection context (read around current selection)
 * - change tracker summary (cells edited since last message)
 */

import type { AgentMessage } from "@mariozechner/pi-agent-core";

import type { ChangeTracker } from "../context/change-tracker.js";
import { getBlueprint, getBlueprintRevision } from "../context/blueprint.js";
import { readSelectionContext } from "../context/selection.js";
import { getFilesWorkspace } from "../files/workspace.js";
import { extractTextFromContent } from "../utils/content.js";
import { getWorkbookContext } from "../workbook/context.js";

import {
  decideWorkbookContextRefresh,
  type BlueprintRefreshReason,
} from "./context-refresh-decision.js";

const AUTO_CONTEXT_PREFIX = "[Auto-context]";
const WORKBOOK_CONTEXT_REFRESH_PREFIX = "[Workbook context refresh:";
const WORKSPACE_FILES_REFRESH_PREFIX = "[Workspace files refresh:";

function withTimeout<T>(promise: Promise<T>, timeoutMs: number): Promise<T | null> {
  let timeoutId: ReturnType<typeof setTimeout> | undefined;
  const timeoutPromise = new Promise<null>((resolve) => {
    timeoutId = setTimeout(() => resolve(null), timeoutMs);
  });

  const raced = Promise.race<T | null>([promise, timeoutPromise]);
  return raced.finally(() => {
    if (timeoutId !== undefined) clearTimeout(timeoutId);
  });
}

function buildWorkbookContextSection(blueprint: string, reason: BlueprintRefreshReason): string {
  let reasonText = "initial";
  if (reason === "workbook_switched") reasonText = "workbook switched";
  if (reason === "blueprint_invalidated") reasonText = "workbook structure changed";
  if (reason === "context_missing") reasonText = "context missing from history";

  return [
    `[Workbook context refresh: ${reasonText}]`,
    blueprint,
  ].join("\n\n");
}

function historyHasWorkbookContextRefresh(messages: readonly AgentMessage[]): boolean {
  for (const message of messages) {
    if (!(message.role === "user" || message.role === "user-with-attachments")) {
      continue;
    }

    const text = extractTextFromContent(message.content);
    if (!text.includes(AUTO_CONTEXT_PREFIX)) continue;
    if (text.includes(WORKBOOK_CONTEXT_REFRESH_PREFIX)) {
      return true;
    }
  }

  return false;
}

function historyHasWorkspaceFilesRefresh(messages: readonly AgentMessage[]): boolean {
  for (const message of messages) {
    if (!(message.role === "user" || message.role === "user-with-attachments")) {
      continue;
    }

    const text = extractTextFromContent(message.content);
    if (!text.includes(AUTO_CONTEXT_PREFIX)) continue;
    if (text.includes(WORKSPACE_FILES_REFRESH_PREFIX)) {
      return true;
    }
  }

  return false;
}

function buildWorkspaceFilesSection(summary: string, reason: "initial" | "files_changed"): string {
  const reasonText = reason === "files_changed" ? "files changed" : "initial";
  return [`[Workspace files refresh: ${reasonText}]`, summary].join("\n\n");
}

export function createContextInjector(changeTracker: ChangeTracker) {
  let lastInjectedWorkbookId: string | null | undefined;
  let lastInjectedBlueprintRevision = -1;
  let lastInjectedWorkspaceSignature: string | undefined;

  return async (messages: AgentMessage[], _signal?: AbortSignal): Promise<AgentMessage[]> => {
    const injections: string[] = [];

    // Workbook structure context: inject only when needed.
    try {
      const workbookCtx = await withTimeout(getWorkbookContext().catch(() => null), 1200);
      const workbookId = workbookCtx?.workbookId ?? null;
      const currentRevision = getBlueprintRevision(workbookId);
      const hasWorkbookContextMessage = historyHasWorkbookContextRefresh(messages);

      const decision = decideWorkbookContextRefresh({
        lastInjectedWorkbookId,
        lastInjectedBlueprintRevision,
        currentWorkbookId: workbookId,
        currentBlueprintRevision: currentRevision,
        hasWorkbookContextMessage,
      });

      if (decision.shouldBootstrap) {
        lastInjectedWorkbookId = workbookId;
        lastInjectedBlueprintRevision = currentRevision;
      }

      if (decision.refreshReason) {
        const blueprint = await withTimeout(getBlueprint(workbookId).catch(() => null), 2500);
        if (blueprint && blueprint.trim().length > 0) {
          injections.push(buildWorkbookContextSection(blueprint, decision.refreshReason));
          lastInjectedWorkbookId = workbookId;
          lastInjectedBlueprintRevision = getBlueprintRevision(workbookId);
        }
      }
    } catch {
      // ignore
    }

    try {
      const workspace = getFilesWorkspace();
      const snapshot = await withTimeout(workspace.getSnapshot().catch(() => null), 1200);

      if (snapshot) {
        const signature = `${snapshot.backend.kind}|${snapshot.signature}`;
        const hasWorkspaceMessage = historyHasWorkspaceFilesRefresh(messages);

        if (lastInjectedWorkspaceSignature === undefined && hasWorkspaceMessage) {
          lastInjectedWorkspaceSignature = signature;
        }

        const shouldInjectInitial =
          lastInjectedWorkspaceSignature === undefined &&
          !hasWorkspaceMessage &&
          snapshot.files.length > 0;
        const shouldInjectChanged =
          lastInjectedWorkspaceSignature !== undefined &&
          signature !== lastInjectedWorkspaceSignature;

        if (shouldInjectInitial || shouldInjectChanged) {
          const summary = await workspace.getContextSummary(20);
          const section = summary
            ? buildWorkspaceFilesSection(summary, shouldInjectChanged ? "files_changed" : "initial")
            : buildWorkspaceFilesSection(
              "### Workspace Files\n- _No files currently available._",
              "files_changed",
            );
          injections.push(section);
        }

        lastInjectedWorkspaceSignature = signature;
      } else {
        lastInjectedWorkspaceSignature = undefined;
      }
    } catch {
      // ignore
    }

    try {
      const sel = await withTimeout(readSelectionContext().catch(() => null), 1500);
      if (sel) injections.push(sel.text);
    } catch {
      // ignore
    }

    const changes = changeTracker.flush();
    if (changes) injections.push(changes);
    if (injections.length === 0) return messages;

    const injection = injections.join("\n\n");
    const injectionMessage: AgentMessage = {
      role: "user",
      content: [{ type: "text", text: `${AUTO_CONTEXT_PREFIX}\n${injection}` }],
      timestamp: Date.now(),
    };

    const nextMessages = [...messages];
    let lastUserIdx = -1;
    for (let i = nextMessages.length - 1; i >= 0; i--) {
      if (nextMessages[i].role === "user") {
        lastUserIdx = i;
        break;
      }
    }

    if (lastUserIdx >= 0) {
      nextMessages.splice(lastUserIdx, 0, injectionMessage);
    } else {
      nextMessages.push(injectionMessage);
    }

    return nextMessages;
  };
}
