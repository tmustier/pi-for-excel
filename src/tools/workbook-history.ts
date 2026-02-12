/**
 * workbook_history — list / restore workbook recovery checkpoints.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";

import {
  getWorkbookRecoveryLog,
  type WorkbookRecoverySnapshot,
} from "../workbook/recovery-log.js";
import {
  getWorkbookChangeAuditLog,
  type AppendWorkbookChangeAuditEntryArgs,
} from "../audit/workbook-change-audit.js";
import { getErrorMessage } from "../utils/errors.js";
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
        "Operation to run. list (default): show recent checkpoints; " +
        "restore: revert one checkpoint; delete: remove one checkpoint; clear: remove all checkpoints for current workbook.",
    }),
  ),
  snapshot_id: Type.Optional(
    Type.String({
      description: "Checkpoint id for restore/delete. If omitted, the latest checkpoint is used.",
    }),
  ),
  limit: Type.Optional(
    Type.Integer({
      minimum: 1,
      maximum: 50,
      description: "Max checkpoints to list (list action only). Default: 10.",
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
  lines.push("Recent recovery checkpoints (current workbook):");
  lines.push("");
  lines.push("| ID | Time | Tool | Range | Changed |");
  lines.push("| --- | --- | --- | --- | ---: |");

  for (const snapshot of snapshots) {
    const changed = snapshot.changedCount.toLocaleString();
    const toolLabel = snapshot.toolName === "restore_snapshot" ? "restore" : snapshot.toolName;
    lines.push(`| \`${snapshot.id}\` | ${formatTimestamp(snapshot.at)} | ${toolLabel} | ${snapshot.address} | ${changed} |`);
  }

  lines.push("");
  lines.push("Use `workbook_history` with `action: \"restore\"` and `snapshot_id` to revert a specific checkpoint.");
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

  return {
    name: "workbook_history",
    label: "Workbook History",
    description:
      "List, restore, and manage automatic workbook recovery checkpoints created before agent edits.",
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
              content: [{ type: "text", text: "No recovery checkpoints for this workbook yet." }],
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
              content: [{ type: "text", text: "No recovery checkpoints available to restore." }],
              details: {
                kind: "workbook_history",
                action: "restore",
                error: "missing_snapshot",
              },
            };

            await resolvedDependencies.appendAuditEntry({
              toolName: "workbook_history",
              toolCallId,
              blocked: true,
              changedCount: 0,
              changes: [],
              summary: "error: no recovery checkpoints available to restore",
            });

            return restoreUnavailableResult;
          }

          const restored = await log.restore(snapshotId);
          const lines: string[] = [];
          lines.push(`✅ Restored checkpoint \`${shortId(restored.restoredSnapshotId)}\` at **${restored.address}**.`);
          lines.push(`Changed item(s): ${restored.changedCount.toLocaleString()}.`);

          if (restored.inverseSnapshotId) {
            lines.push(`Rollback checkpoint created: \`${shortId(restored.inverseSnapshotId)}\`.`);
          }

          await resolvedDependencies.appendAuditEntry({
            toolName: "workbook_history",
            toolCallId,
            blocked: false,
            outputAddress: restored.address,
            changedCount: restored.changedCount,
            changes: [],
            summary: `restored checkpoint ${shortId(restored.restoredSnapshotId)} at ${restored.address}`,
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
              content: [{ type: "text", text: "No recovery checkpoints available to delete." }],
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
              content: [{ type: "text", text: "Checkpoint not found." }],
              details: {
                kind: "workbook_history",
                action: "delete",
                snapshotId,
                deletedCount: 0,
              },
            };
          }

          return {
            content: [{ type: "text", text: `Deleted checkpoint \`${shortId(snapshotId)}\`.` }],
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
          content: [{ type: "text", text: `Cleared ${removed} checkpoint${removed === 1 ? "" : "s"} for this workbook.` }],
          details: {
            kind: "workbook_history",
            action: "clear",
            deletedCount: removed,
          },
        };
      } catch (error: unknown) {
        const message = getErrorMessage(error);

        if (action === "restore") {
          await resolvedDependencies.appendAuditEntry({
            toolName: "workbook_history",
            toolCallId,
            blocked: true,
            changedCount: 0,
            changes: [],
            summary: `error: ${message}`,
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
