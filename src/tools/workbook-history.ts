import { Type, type Static } from "typebox";
import { recoveryTurn, type RecoveryAttribution } from "./recovery-metadata.js";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  getWorkbookRecoveryLog,
  type WorkbookRecoveryRestorePreview,
} from "../workbook/recovery-log.js";
import { RecoveryCursorSchema, type RecoveryPage, type RecoveryPageOptions, type RecoveryMetadata } from "../workbook/recovery/log-store.js";
import {
  getWorkbookChangeAuditLog,
  type AppendWorkbookChangeAuditEntryArgs,
} from "../audit/workbook-change-audit.js";
import { getErrorMessage } from "../utils/errors.js";
import { finalizeMutationOperation } from "./mutation/finalize.js";
import type { MutationFinalizeDependencies } from "./mutation/types.js";
import type { WorkbookHistoryDetails } from "./tool-details.js";

const schema = Type.Object({
  action: Type.Optional(
    Type.Union([
      Type.Literal("list"),
      Type.Literal("restore"),
      Type.Literal("delete"),
      Type.Literal("clear"),
    ], {
      description:
        "Operation to run. list (default): show recent backups; " +
        "restore: revert one backup; delete: remove one backup; clear: remove all backups for current workbook.",
    }),
  ),
  snapshot_id: Type.Optional(
    Type.String({
      description: "Backup id for restore/delete. If omitted, the latest backup is used.",
    }),
  ),
  overwrite_later_edits: Type.Optional(
    Type.Boolean({
      description:
        "Restore only. Default: false. If cells changed after the backed-up change, restore is refused and the " +
        "changed cells are returned; set true only after the user confirms those edits may be overwritten.",
    }),
  ),
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 50,
      description: "Max backups to list (list action only). Default: 10.",
    }),
  ),
  cursor: Type.Optional(RecoveryCursorSchema),
});

type Params = Static<typeof schema>;

type RestoreResult = {
  restoredSnapshotId: string;
  inverseSnapshotId: string | null;
  inverseSnapshotError?: string;
  address: string;
  changedCount: number;
};

interface WorkbookHistoryRecoveryLogLike {
  listForCurrentWorkbook(limit?: number): Promise<Array<Pick<RecoveryMetadata, "id">>>;
  pageForCurrentWorkbook(options?: RecoveryPageOptions): Promise<RecoveryPage>;
  restore(snapshotId: string, attribution?: { turnId?: string; toolCallId?: string }): Promise<RestoreResult>;
  previewRestore(snapshotId: string): Promise<WorkbookRecoveryRestorePreview>;
  delete(snapshotId: string): Promise<boolean>;
  clearForCurrentWorkbook(): Promise<number>;
}

interface WorkbookHistoryToolDependencies {
  getRecoveryLog: () => WorkbookHistoryRecoveryLogLike;
  appendAuditEntry: (entry: AppendWorkbookChangeAuditEntryArgs) => Promise<void>;
}

const defaultDependencies: WorkbookHistoryToolDependencies = {
  getRecoveryLog: () => getWorkbookRecoveryLog(),
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

function formatTimestamp(ts: number): string {
  try {
    return new Date(ts).toLocaleString();
  } catch {
    return String(ts);
  }
}

function buildListMarkdown(snapshots: RecoveryMetadata[]): string {
  const lines: string[] = [];
  lines.push("Recent backups (current workbook):");
  lines.push("");
  lines.push("| ID | Time | Tool | Range | Changed |");
  lines.push("| --- | --- | --- | --- | ---: |");

  for (const snapshot of snapshots) {
    const changed = snapshot.changedCount.toLocaleString();
    const toolLabel = snapshot.toolName === "restore_snapshot" ? "restore" : snapshot.toolName;
    lines.push(`| \`${snapshot.id}\` | ${formatTimestamp(snapshot.at)} | ${toolLabel} | ${snapshot.address} | ${changed} |`);
  }

  lines.push("");
  lines.push("Use `workbook_history` with `action: \"restore\"` and `snapshot_id` to revert a specific backup.");
  return lines.join("\n");
}

async function resolveSnapshotId(
  log: WorkbookHistoryRecoveryLogLike,
  params: Params,
): Promise<string | null> {
  const explicit = params.snapshot_id?.trim();
  if (explicit) return explicit;

  const latest = await log.listForCurrentWorkbook(1);
  return latest[0]?.id ?? null;
}

export function createWorkbookHistoryTool(
  dependencies: Partial<WorkbookHistoryToolDependencies> = {},
  attribution: RecoveryAttribution = {},
): AgentTool<typeof schema, WorkbookHistoryDetails> {
  const resolvedDependencies: WorkbookHistoryToolDependencies = {
    getRecoveryLog: dependencies.getRecoveryLog ?? defaultDependencies.getRecoveryLog,
    appendAuditEntry: dependencies.appendAuditEntry ?? defaultDependencies.appendAuditEntry,
  };
  const mutationFinalizeDependencies: MutationFinalizeDependencies = {
    appendAuditEntry: (entry) => resolvedDependencies.appendAuditEntry(entry),
  };

  return {
    name: "workbook_history",
    label: "Workbook History",
    description:
      "List, restore, and manage automatic workbook backups created before Pi edits.",
    parameters: schema,
    execute: async (toolCallId: string, params: Params): Promise<AgentToolResult<WorkbookHistoryDetails>> => {
      const action = params.action ?? "list";
      const log = resolvedDependencies.getRecoveryLog();

      try {
        if (action === "list") {
          const limit = params.limit ?? 10;
          const page = await log.pageForCurrentWorkbook({ limit, ...(params.cursor ? { cursor: params.cursor } : {}) });
          const snapshots = page.entries;

          if (snapshots.length === 0) {
            return {
              content: [{ type: "text", text: "No backups for this workbook yet." }],
              details: {
                kind: "workbook_history",
                action: "list",
                count: 0,
                snapshots: [],
              },
            };
          }

          return {
            content: [{ type: "text", text: buildListMarkdown(snapshots) + (page.nextCursor
              ? `\n\nMore backups: call list with cursor ${JSON.stringify(page.nextCursor)}.` : "") }],
            details: {
              kind: "workbook_history",
              action: "list",
              count: snapshots.length,
              ...(page.nextCursor ? { nextCursor: page.nextCursor } : {}),
              snapshots: snapshots.map((snapshot) => ({
                id: snapshot.id,
                at: snapshot.at,
                toolName: snapshot.toolName,
                address: snapshot.address,
                changedCount: snapshot.changedCount,
                cellCount: snapshot.cellCount,
                ...(snapshot.workbookId !== undefined ? { workbookId: snapshot.workbookId } : {}),
                ...(snapshot.workbookLabel !== undefined ? { workbookLabel: snapshot.workbookLabel } : {}),
              })),
            },
          };
        }

        if (action === "restore") {
          const snapshotId = await resolveSnapshotId(log, params);
          if (!snapshotId) {
            const restoreUnavailableResult: AgentToolResult<WorkbookHistoryDetails> = {
              content: [{ type: "text", text: "No backups available to restore." }],
              details: {
                kind: "workbook_history",
                action: "restore",
                error: "missing_snapshot",
              },
            };

            await finalizeMutationOperation(mutationFinalizeDependencies, {
              auditEntry: {
                toolName: "workbook_history",
                toolCallId,
                blocked: true,
                changedCount: 0,
                changes: [],
                summary: "error: no backups available to restore",
              },
            });

            return restoreUnavailableResult;
          }

          const preview = await log.previewRestore(snapshotId);
          const age = formatTimestamp(preview.at);
          if (preview.laterEdits && !params.overwrite_later_edits) {
            const sample = preview.laterEdits.sample.map((cell) => `- ${cell.address}: ${cell.content}`);
            await finalizeMutationOperation(mutationFinalizeDependencies, {
              auditEntry: {
                toolName: "workbook_history",
                toolCallId,
                blocked: true,
                outputAddress: preview.address,
                changedCount: 0,
                changes: [],
                summary: `restore of ${snapshotId} refused: ${preview.laterEdits.count} cell(s) edited since`,
              },
            });
            return {
              content: [{
                type: "text",
                text: [
                  `⛔ **Restore not done** — backup \`${snapshotId}\` of **${preview.address}** is from ${age}, and ` +
                    `${preview.laterEdits.count} cell(s) have changed since that change. Restoring would overwrite them:`,
                  ...sample,
                  "Ask the user; if they want those edits overwritten, retry with `overwrite_later_edits: true`.",
                ].join("\n"),
              }],
              details: {
                kind: "workbook_history",
                action: "restore",
                snapshotId,
                address: preview.address,
                error: "later_edits",
              },
            };
          }

          const restored = await log.restore(snapshotId, { ...recoveryTurn(attribution, toolCallId), toolCallId });
          const lines: string[] = [];
          lines.push(`✅ Restored backup \`${(restored.restoredSnapshotId)}\` (from ${age}) at **${restored.address}**.`);
          lines.push(`Changed item(s): ${restored.changedCount.toLocaleString()}.`);

          if (restored.inverseSnapshotId) {
            lines.push(`Rollback backup created: \`${(restored.inverseSnapshotId)}\`.`);
          } else if (restored.inverseSnapshotError) {
            lines.push(`⚠️ ${restored.inverseSnapshotError} This restore cannot itself be reverted from backups.`);
          }

          await finalizeMutationOperation(mutationFinalizeDependencies, {
            auditEntry: {
              toolName: "workbook_history",
              toolCallId,
              blocked: false,
              outputAddress: restored.address,
              changedCount: restored.changedCount,
              changes: [],
              summary: `restored backup ${(restored.restoredSnapshotId)} at ${restored.address}`,
            },
          });

          return {
            content: [{ type: "text", text: lines.join("\n\n") }],
            details: {
              kind: "workbook_history",
              action: "restore",
              snapshotId,
              restoredSnapshotId: restored.restoredSnapshotId,
              ...(typeof restored.inverseSnapshotId === "string" ? { inverseSnapshotId: restored.inverseSnapshotId } : {}),
              address: restored.address,
              changedCount: restored.changedCount,
            },
          };
        }

        if (action === "delete") {
          const snapshotId = await resolveSnapshotId(log, params);
          if (!snapshotId) {
            return {
              content: [{ type: "text", text: "No backups available to delete." }],
              details: {
                kind: "workbook_history",
                action: "delete",
                error: "missing_snapshot",
              },
            };
          }

          const deleted = await log.delete(snapshotId);
          if (!deleted) {
            return {
              content: [{ type: "text", text: "Backup not found." }],
              details: {
                kind: "workbook_history",
                action: "delete",
                snapshotId,
                deletedCount: 0,
              },
            };
          }

          return {
            content: [{ type: "text", text: `Deleted backup \`${(snapshotId)}\`.` }],
            details: {
              kind: "workbook_history",
              action: "delete",
              snapshotId,
              deletedCount: 1,
            },
          };
        }

        const removed = await log.clearForCurrentWorkbook();
        return {
          content: [{ type: "text", text: `Cleared ${removed} backup${removed === 1 ? "" : "s"} for this workbook.` }],
          details: {
            kind: "workbook_history",
            action: "clear",
            deletedCount: removed,
          },
        };
      } catch (error) {
        const message = getErrorMessage(error);

        if (action === "restore") {
          await finalizeMutationOperation(mutationFinalizeDependencies, {
            auditEntry: {
              toolName: "workbook_history",
              toolCallId,
              blocked: true,
              changedCount: 0,
              changes: [],
              summary: `error: ${message}`,
            },
          });
        }

        return {
          content: [{ type: "text", text: `Error: ${message}` }],
          details: {
            kind: "workbook_history",
            action,
            error: message,
          },
        };
      }
    },
  };
}
