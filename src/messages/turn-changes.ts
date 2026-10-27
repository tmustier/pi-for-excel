import { Type, type Static } from "typebox";
import { StringEnum } from "../tools/string-enum.js";
import type { WorkbookChangeAuditEntry } from "../audit/workbook-change-audit.js";
import type { WorkbookRecoverySnapshot } from "../workbook/recovery-log.js";
import type { RecoveryMetadata, RecoveryPage, RecoveryPageOptions } from "../workbook/recovery/log-store.js";

/** A turn may contain more than one metadata page; never load its payloads. */
export async function readTurnBackups(
  log: { pageForCurrentWorkbook(options: RecoveryPageOptions): Promise<RecoveryPage> },
  turnId: string,
): Promise<RecoveryMetadata[]> {
  let page = await log.pageForCurrentWorkbook({ turnId });
  const backups = [...page.entries];
  while (page.nextCursor) {
    page = await log.pageForCurrentWorkbook({ turnId, cursor: page.nextCursor });
    backups.push(...page.entries);
  }
  return backups;
}

export const turnChangesSchema = Type.Object({
  role: Type.Literal("turnChanges"),
  turnId: Type.String(),
  timestamp: Type.Number(),
  outcome: StringEnum(["finished", "stopped", "failed"]),
  ranOfficeJs: Type.Optional(Type.Boolean()),
  undoneBackupIds: Type.Optional(Type.Array(Type.String())),
  changes: Type.Array(Type.Object({
    address: Type.String(),
    changedCount: Type.Number(),
    backupIds: Type.Array(Type.String()),
  })),
});

export type TurnChangesMessage = Static<typeof turnChangesSchema>;

export const turnChangeActionSchema = Type.Object({
  action: StringEnum(["review", "undo"]),
  summary: turnChangesSchema,
});

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    turnChanges: TurnChangesMessage;
  }
}

export function summarizeTurnChanges(args: {
  turnId: string;
  timestamp: number;
  outcome: TurnChangesMessage["outcome"];
  toolCallIds: ReadonlySet<string>;
  ranOfficeJs?: boolean;
  audit: readonly WorkbookChangeAuditEntry[];
  backups: readonly Pick<WorkbookRecoverySnapshot, "id" | "at" | "toolCallId" | "address" | "changedCount">[];
}): TurnChangesMessage | null {
  const changes: TurnChangesMessage["changes"] = [];
  // Calls start in workbook-queue order. Preserve that order even for partial
  // failures with no audit entry, and when timestamps share a millisecond.
  for (const callId of args.toolCallIds) {
    const saved = args.backups.filter((backup) => backup.toolCallId === callId)
      .sort((a, b) => a.at - b.at || a.id.localeCompare(b.id));
    const entry = args.audit.find((item) => item.toolCallId === callId && !item.blocked && item.changedCount > 0);
    if (entry) {
      changes.push({ address: entry.outputAddress ?? entry.inputAddress ?? entry.toolName,
        changedCount: entry.changedCount, backupIds: saved.map((backup) => backup.id) });
    } else {
      // A failed tool can still have changed the workbook and saved a backup.
      for (const backup of saved) changes.push({ address: backup.address, changedCount: backup.changedCount, backupIds: [backup.id] });
    }
  }
  return changes.length === 0 && !args.ranOfficeJs ? null : {
    role: "turnChanges", turnId: args.turnId, timestamp: args.timestamp, outcome: args.outcome, changes,
    ...(args.ranOfficeJs ? { ranOfficeJs: true } : {}),
  };
}

/** Each restore is ordinary recovery: it captures its own inverse backup. */
export async function undoTurnChanges(
  summary: TurnChangesMessage,
  restore: (backupId: string) => Promise<boolean>,
): Promise<number> {
  let restored = 0;
  const undone = new Set(summary.undoneBackupIds);
  for (const change of [...summary.changes].reverse()) {
    for (const id of [...change.backupIds].reverse()) {
      if (undone.has(id)) continue;
      if (!await restore(id)) return restored;
      undone.add(id);
      summary.undoneBackupIds = [...undone];
      restored += 1;
    }
  }
  return restored;
}
