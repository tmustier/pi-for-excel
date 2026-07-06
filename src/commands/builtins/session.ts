/**
 * Builtin session management commands.
 */

import type { SlashCommand } from "../types.js";
import type { ResumeDialogTarget } from "./resume-target.js";
import { requestConfirmationDialog } from "../../ui/confirm-dialog.js";
import { showToast } from "../../ui/toast.js";
import { t } from "../../language/index.js";

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
      description: t("command.session.name"),
      source: "builtin",
      execute: async (args: string) => {
        const title = args.trim();
        if (!title) {
          showToast(t("session.toast.name_usage"));
          return;
        }

        await actions.renameActiveSession(title);
        showToast(t("session.toast.named", { title }));
      },
    },
    {
      name: "share-session",
      description: t("command.session.share"),
      source: "builtin",
      execute: () => {
        showToast(t("session.toast.sharing_soon"));
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
  return t("session.toast.backup_usage");
}

export function createSessionLifecycleCommands(actions: SessionCommandActions): SlashCommand[] {
  return [
    {
      name: "new",
      description: t("command.session.new"),
      source: "builtin",
      execute: async () => {
        await actions.createRuntime();
      },
    },
    {
      name: "resume",
      description: t("command.session.resume"),
      source: "builtin",
      execute: async () => {
        await actions.openResumeDialog("new_tab");
      },
    },
    {
      name: "resume-here",
      description: t("command.resume_here.desc"),
      source: "builtin",
      execute: async () => {
        await actions.openResumeDialog("replace_current");
      },
    },
    {
      name: "history",
      description: t("command.session.backups"),
      source: "builtin",
      execute: async () => {
        await actions.openRecoveryDialog();
      },
    },
    {
      name: "backup",
      description: t("command.session.backup"),
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
              t("session.toast.backup_created", { id: shortBackupId(backup.id), size: formatBytes(backup.sizeBytes) }),
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
              showToast(t("session.toast.no_backups"));
              return;
            }

            const preview = backups
              .map((backup) => shortBackupId(backup.id))
              .slice(0, 3)
              .join(", ");

            const hasMore = backups.length > 3;
            const previewText = hasMore ? `${preview}, …` : preview;
            showToast(t("session.toast.backups_list", { count: String(backups.length), preview: previewText }));
            return;
          }

          if (action === "restore") {
            const restored = await actions.restoreManualFullBackup(tailText.length > 0 ? tailText : undefined);
            if (!restored) {
              showToast(t("session.toast.backup_not_found"));
              return;
            }

            showToast(t("session.toast.backup_restored", { id: shortBackupId(restored.id) }));
            return;
          }

          if (action === "clear") {
            const proceed = await requestConfirmationDialog({
              title: t("session.confirm.clear_title"),
              message: t("session.confirm.clear_message"),
              confirmLabel: t("session.confirm.clear_label"),
              cancelLabel: t("session.confirm.cancel"),
              confirmButtonTone: "danger",
              restoreFocusOnClose: true,
            });
            if (!proceed) {
              return;
            }

            const removed = await actions.clearManualFullBackups();
            showToast(t("session.toast.backups_deleted", { count: String(removed) }));
            return;
          }

          showToast(backupUsage());
        } catch (error) {
          const message = error instanceof Error ? error.message : t("session.backup.unknown_error");
          showToast(t("session.toast.backup_failed", { message }));
        }
      },
    },
    {
      name: "reopen",
      description: t("command.session.reopen"),
      source: "builtin",
      execute: async () => {
        await actions.reopenLastClosed();
      },
    },
    {
      name: "revert",
      description: t("command.session.revert"),
      source: "builtin",
      execute: async () => {
        await actions.revertLatestCheckpoint();
      },
    },
  ];
}
