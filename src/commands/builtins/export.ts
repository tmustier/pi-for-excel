/** Builtin export and native Durable compaction commands. */
import { BACKGROUND_CONTEXT as context } from "@earendil-works/chord/context";
import type { AgentMessage } from "@earendil-works/pi-agent-core";
import type { StopReason, Usage } from "@earendil-works/pi-ai";
import type { ConversationPresenter } from "../../durable/conversation-view.js";
import type { DurableSessions } from "../../durable/sessions.js";
import type { SlashCommand } from "../types.js";
import type { ActiveConversationProvider } from "./model.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";
import { getErrorMessage } from "../../utils/errors.js";
import { summarizeContentForTranscript } from "../../utils/content.js";
import { getWorkbookChangeAuditLog } from "../../audit/workbook-change-audit.js";

type TranscriptEntry = {
  role: AgentMessage["role"];
  text: string;
  usage?: Usage;
  stopReason?: StopReason;
};

function messageToTranscriptText(message: AgentMessage): string {
  if (message.role === "archivedMessages") {
    return `[archived history: ${message.archivedChatMessageCount} chat messages]`;
  }
  if (message.role === "compactionSummary") return message.summary;
  return "content" in message ? summarizeContentForTranscript(message.content) : "";
}

type ExportDestination = "clipboard" | "download";

function parseExportDestination(raw: string, fallback: ExportDestination): ExportDestination {
  const normalized = raw.trim().toLowerCase();
  if (normalized === "clipboard") return "clipboard";
  if (normalized === "file" || normalized === "download") return "download";
  return fallback;
}

function triggerJsonDownload(fileName: string, content: string): void {
  const blob = new Blob([content], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  // Office WebViews may need a popup rather than an anchor download.
  const opened = window.open(url, "_blank");
  if (!opened) {
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = fileName;
    anchor.rel = "noopener";
    anchor.hidden = true;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
  }
  setTimeout(() => URL.revokeObjectURL(url), 60_000);
}

async function exportWorkbookAuditLog(rawArgs: string): Promise<void> {
  const destination = parseExportDestination(rawArgs, "download");
  const entries = await getWorkbookChangeAuditLog().list(500);
  const json = JSON.stringify({ exported: new Date().toISOString(), count: entries.length, entries }, null, 2);
  if (destination === "clipboard") {
    await navigator.clipboard.writeText(json);
    showToast(t("export.toast.audit_copied", {
      count: String(entries.length), size: (json.length / 1024).toFixed(0),
    }));
    return;
  }
  triggerJsonDownload(`pi-audit-log-${new Date().toISOString().slice(0, 10)}.json`, json);
  showToast(t("export.toast.audit_downloaded", { count: String(entries.length) }));
}

export function createExportCommands(
  getActiveConversation: ActiveConversationProvider,
  getSessions: () => DurableSessions,
): SlashCommand[] {
  return [{
    name: "export",
    description: t("command.export.json"),
    source: "builtin",
    execute: async (args: string) => {
      const parts = args.trim().split(/\s+/u).filter((part) => part.length > 0);
      const mode = parts[0]?.toLowerCase();
      if (mode === "audit" || mode === "audit-log") {
        try {
          await exportWorkbookAuditLog(parts.slice(1).join(" "));
        } catch (error) {
          showToast(t("export.toast.audit_export_failed", { error: getErrorMessage(error) }));
        }
        return;
      }
      const view = getActiveConversation();
      if (!view) {
        showToast(t("export.toast.no_session"));
        return;
      }
      // Export the immutable catalogue timeline, never the active UI projection.
      const session = await getSessions().loadSession(String(view.conversation.id));
      if (!session) {
        showToast(t("export.toast.no_session"));
        return;
      }
      const msgs = session.messages;
      if (msgs.length === 0) {
        showToast(t("export.toast.no_messages"));
        return;
      }
      const transcript: TranscriptEntry[] = msgs.map((message) => ({
        role: message.role,
        text: messageToTranscriptText(message),
        ...(message.role === "assistant" ? { usage: message.usage, stopReason: message.stopReason } : {}),
      }));
      const json = JSON.stringify({
        exported: new Date().toISOString(),
        id: session.id, title: session.title, createdAt: session.createdAt, lastModified: session.lastModified,
        model: session.model ? { id: session.model.id, name: session.model.name, provider: session.model.provider } : null,
        thinkingLevel: session.thinkingLevel,
        messageCount: msgs.length, transcript, raw: msgs,
      }, null, 2);
      if (parseExportDestination(args, "clipboard") === "clipboard") {
        try {
          await navigator.clipboard.writeText(json);
          showToast(t("export.toast.transcript_copied", {
            count: String(msgs.length), size: (json.length / 1024).toFixed(0),
          }));
        } catch (error) {
          showToast(t("export.toast.copy_failed", { error: getErrorMessage(error) }));
        }
        return;
      }
      triggerJsonDownload(`pi-session-${new Date().toISOString().slice(0, 10)}.json`, json);
      showToast(t("export.toast.transcript_downloaded", { count: String(msgs.length) }));
    },
  }];
}

/** Admission is not completion: wait for the native task's committed outcome. */
export async function runCompactCommand(view: ConversationPresenter, args: string): Promise<void> {
  showToast(t("export.toast.compact.compacting"), 60000);
  try {
    // Durable owns busy admission and summary placement; do not abort or drain its inbox.
    const task = await view.conversation.compact(args.trim() || undefined, context);
    const settled = await view.harness.waitForTask(task, context);
    const outcome = settled.state.outcome;
    if (outcome.status !== "completed") {
      const msg = outcome.status === "failed" ? outcome.error.message
        : outcome.reason ?? t("export.toast.compact.failed_error");
      showToast(t("export.toast.compact.failed", { msg }));
      return;
    }
    const result = outcome.result;
    showToast(result.entryId || result.submissionId
      ? `${t("status-popovers.compactTitle")} ✓`
      : t("export.toast.compact.nothing"));
  } catch (error) {
    showToast(t("export.toast.compact.failed", { msg: getErrorMessage(error) }));
  }
}

export function createCompactCommands(getActiveConversation: ActiveConversationProvider): SlashCommand[] {
  return [{
    name: "compact",
    description: t("command.export.summarize"),
    source: "builtin",
    execute: async (args: string) => {
      const view = getActiveConversation();
      if (!view) {
        showToast(t("export.toast.no_session"));
        return;
      }
      await runCompactCommand(view, args);
    },
  }];
}
