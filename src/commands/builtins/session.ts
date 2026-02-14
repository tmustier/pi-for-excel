/**
 * Builtin session management commands.
 */

import type { SlashCommand } from "../types.js";
import type { ResumeDialogTarget } from "./resume-target.js";
import { showToast } from "../../ui/toast.js";

export interface ManualFullBackupSummary {
  id: string;
  createdAt: number;
  sizeBytes: number;
}

export interface SessionCommandActions {
  renameActiveSession: (title: string) => Promise<void>;
  createRuntime: () => Promise<void>;
  openResumeDialog: (defaultTarget?: ResumeDialogTarget) => Promise<void>;
  openRecoveryDialog: () => Promise<void>;
  reopenLastClosed: () => Promise<void>;
  revertLatestCheckpoint: () => Promise<void>;
  createManualFullBackup: () => Promise<ManualFullBackupSummary>;
  listManualFullBackups: (limit?: number) => Promise<ManualFullBackupSummary[]>;
  restoreManualFullBackup: (backupId?: string) => Promise<ManualFullBackupSummary | null>;
  clearManualFullBackups: () => Promise<number>;
}

export function createSessionIdentityCommands(actions: SessionCommandActions): SlashCommand[] {
  return [
    {
      name: "name",
      description: "Name the current chat session",
      source: "builtin",
      execute: async (args: string) => {
        const title = args.trim();
        if (!title) {
          showToast("Usage: /name My Session Name");
          return;
        }

        await actions.renameActiveSession(title);
        showToast(`Session named: ${title}`);
      },
    },
    {
      name: "share-session",
      description: "Share session as a link",
      source: "builtin",
      execute: () => {
        showToast("Session sharing coming soon");
      },
    },
  ];
}

function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) {
    return "0 B";
  }

  if (bytes < 1024) {
    return `${Math.floor(bytes)} B`;
  }

  const units = ["KB", "MB", "GB"];
  let value = bytes / 1024;
  let unitIndex = 0;

  while (value >= 1024 && unitIndex < units.length - 1) {
    value /= 1024;
    unitIndex += 1;
  }

  return `${value.toFixed(value >= 10 ? 0 : 1)} ${units[unitIndex]}`;
}

function shortBackupId(id: string): string {
  return id.length > 24 ? `${id.slice(0, 24)}…` : id;
}

function backupUsage(): string {
  return "Usage: /backup [create|list [limit]|restore [id]|clear]";
}

export function createSessionLifecycleCommands(actions: SessionCommandActions): SlashCommand[] {
  return [
    {
      name: "new",
      description: "Start a new chat session tab",
      source: "builtin",
      execute: async () => {
        await actions.createRuntime();
      },
    },
    {
      name: "resume",
      description: "Resume a previous session (opens in new tab)",
      source: "builtin",
      execute: async () => {
        await actions.openResumeDialog("new_tab");
      },
    },
    {
      name: "resume-here",
      description: "Resume a previous session into the current tab",
      source: "builtin",
      execute: async () => {
        await actions.openResumeDialog("replace_current");
      },
    },
    {
      name: "history",
      description: "Open Backups (Beta)",
      source: "builtin",
      execute: async () => {
        await actions.openRecoveryDialog();
      },
    },
    {
      name: "backup",
      description: "Manual full-workbook backup (create/list/restore/clear)",
      source: "builtin",
      execute: async (rawArgs: string) => {
        try {
          const trimmed = rawArgs.trim();
          const [actionRaw, ...rest] = trimmed.length > 0 ? trimmed.split(/\s+/) : [];
          const action = (actionRaw ?? "create").toLowerCase();
          const tailText = rest.join(" ").trim();

          if (action === "help") {
            showToast(backupUsage());
            return;
          }

          if (action === "create") {
            const backup = await actions.createManualFullBackup();
            showToast(
              `Manual backup created: #${shortBackupId(backup.id)} (${formatBytes(backup.sizeBytes)}).`,
            );
            return;
          }

          if (action === "list") {
            const parsedLimit = tailText.length > 0 ? Number.parseInt(tailText, 10) : 5;
            const limit = Number.isFinite(parsedLimit)
              ? Math.max(1, Math.min(10, parsedLimit))
              : 5;

            const backups = await actions.listManualFullBackups(limit);
            if (backups.length === 0) {
              showToast("No manual full-workbook backups for this workbook.");
              return;
            }

            const preview = backups
              .map((backup) => shortBackupId(backup.id))
              .slice(0, 3)
              .join(", ");

            const hasMore = backups.length > 3;
            const previewText = hasMore ? `${preview}, …` : preview;
            showToast(`Manual backups (${backups.length} shown): ${previewText}`);
            return;
          }

          if (action === "restore") {
            const restored = await actions.restoreManualFullBackup(tailText.length > 0 ? tailText : undefined);
            if (!restored) {
              showToast("Backup not found for this workbook.");
              return;
            }

            showToast(`Downloaded backup #${shortBackupId(restored.id)}. Open the file in Excel to restore.`);
            return;
          }

          if (action === "clear") {
            const proceed = window.confirm("Delete all manual full-workbook backups for this workbook?");
            if (!proceed) {
              return;
            }

            const removed = await actions.clearManualFullBackups();
            showToast(`Deleted ${removed} manual backup${removed === 1 ? "" : "s"}.`);
            return;
          }

          showToast(backupUsage());
        } catch (error: unknown) {
          const message = error instanceof Error ? error.message : "Unknown error";
          showToast(`Backup command failed: ${message}`);
        }
      },
    },
    {
      name: "reopen",
      description: "Reopen the most recently closed session tab",
      source: "builtin",
      execute: async () => {
        await actions.reopenLastClosed();
      },
    },
    {
      name: "revert",
      description: "Revert the latest workbook backup",
      source: "builtin",
      execute: async () => {
        await actions.revertLatestCheckpoint();
      },
    },
  ];
}
