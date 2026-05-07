/**
 * workbook_history — list / restore workbook backups.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";

import {
  getWorkbookRecoveryLog,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
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
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 50,
      description: "Max backups to list (list action only). Default: 10.",
    }),
  ),
});

type Params = Static<typeof schema>;

type RestoreResult = {
  restoredSnapshotId: string;
  inverseSnapshotId: string | null;
  address: string;
  changedCount: number;
};

interface WorkbookHistoryRecoveryLogLike {
  listForCurrentWorkbook(limit?: number): Promise<WorkbookRecoverySnapshot[]>;
  restore(snapshotId: string): Promise<RestoreResult>;
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

function shortId(id: string): string {
  return id.length > 12 ? id.slice(0, 12) : id;
}

function buildListMarkdown(snapshots: WorkbookRecoverySnapshot[]): string {
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
          const snapshots = await log.listForCurrentWorkbook(limit);

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
            content: [{ type: "text", text: buildListMarkdown(snapshots) }],
            details: {
              kind: "workbook_history",
              action: "list",
              count: snapshots.length,
              snapshots: snapshots.map((snapshot) => ({
                id: snapshot.id,
                at: snapshot.at,
                toolName: snapshot.toolName,
                address: snapshot.address,
                changedCount: snapshot.changedCount,
                cellCount: snapshot.cellCount,
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

          const restored = await log.restore(snapshotId);
          const lines: string[] = [];
          lines.push(`✅ Restored backup \`${shortId(restored.restoredSnapshotId)}\` at **${restored.address}**.`);
          lines.push(`Changed item(s): ${restored.changedCount.toLocaleString()}.`);

          if (restored.inverseSnapshotId) {
            lines.push(`Rollback backup created: \`${shortId(restored.inverseSnapshotId)}\`.`);
          }

          await finalizeMutationOperation(mutationFinalizeDependencies, {
            auditEntry: {
              toolName: "workbook_history",
              toolCallId,
              blocked: false,
              outputAddress: restored.address,
              changedCount: restored.changedCount,
              changes: [],
              summary: `restored backup ${shortId(restored.restoredSnapshotId)} at ${restored.address}`,
            },
          });

          return {
            content: [{ type: "text", text: lines.join("\n\n") }],
            details: {
              kind: "workbook_history",
              action: "restore",
              snapshotId,
              restoredSnapshotId: restored.restoredSnapshotId,
              inverseSnapshotId: restored.inverseSnapshotId ?? undefined,
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
            content: [{ type: "text", text: `Deleted backup \`${shortId(snapshotId)}\`.` }],
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
      } catch (error: unknown) {
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
