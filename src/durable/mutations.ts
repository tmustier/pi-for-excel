import type { WriteCellsGrid, WriteCellsHooks, WriteCellsIntent } from "../tools/write-cells.js";

export type MutationFaultPoint = "beforeSync" | "afterSync" | "beforeResult";
export type MutationClassification = "applied" | "unchanged" | "conflict";
export type MutationRecord = { [K in keyof WriteCellsIntent]: WriteCellsIntent[K] } & {
  /** Opaque workbook identity from getWorkbookContext; never a document URL. */
  workbookId: string;
  phase: "intended" | "synced" | "resultReady";
};
/** Parent supplies Conversation commit/defineDoc persistence. save must commit before resolving. */
export interface MutationStore {
  load(toolCallId: string): Promise<MutationRecord | null>;
  save(record: MutationRecord): Promise<void>;
}
export interface MutationController {
  hooks: WriteCellsHooks;
  configureFault(point: MutationFaultPoint | null): void;
  status(): { armed: MutationFaultPoint | null; paused: MutationFaultPoint | null; toolCallId: string | null };
  resume(): void;
  reconcile(toolCallId: string, read: (target: MutationRecord) => Promise<{
    values: WriteCellsGrid; formulas: WriteCellsGrid;
  }>): Promise<MutationClassification>;
}

/** Records cell content only; callers must not supply credentials or raw workbook URLs. */
export function createMutationController(options: {
  workbookId: string;
  store: MutationStore;
  allowFaultInjection?: boolean;
}): MutationController {
  let armed: MutationFaultPoint | null = null;
  let paused: MutationFaultPoint | null = null;
  let pausedCall: string | null = null;
  let release: (() => void) | null = null;

  const pause = async (point: MutationFaultPoint, toolCallId: string): Promise<void> => {
    if (armed !== point) return;
    armed = null;
    paused = point;
    pausedCall = toolCallId;
    await new Promise<void>((resolve) => { release = resolve; });
    paused = null;
    pausedCall = null;
    release = null;
  };
  const load = async (id: string): Promise<MutationRecord> => {
    const record = await options.store.load(id);
    if (!record || record.workbookId !== options.workbookId) throw new Error("Mutation record unavailable for workbook");
    return record;
  };
  const mark = async (id: string, phase: MutationRecord["phase"], point: MutationFaultPoint): Promise<void> => {
    const record = await load(id);
    await options.store.save({ ...record, phase });
    await pause(point, id);
  };
  return {
    hooks: {
      beforeSync: async (intent) => {
        if (await options.store.load(intent.toolCallId)) {
          throw new Error("Mutation already recorded; reconcile before any retry");
        }
        // Clone at the host boundary: later Office proxy changes cannot mutate the intent.
        await options.store.save(structuredClone({
          ...intent, workbookId: options.workbookId, phase: "intended" as const,
        }));
        await pause("beforeSync", intent.toolCallId);
      },
      afterSync: (id) => mark(id, "synced", "afterSync"),
      beforeResult: (id) => mark(id, "resultReady", "beforeResult"),
    },
    configureFault: (point) => {
      if (point !== null && options.allowFaultInjection !== true) throw new Error("Fault injection is disabled");
      if (paused !== null) throw new Error("Resume the paused mutation first");
      armed = point;
    },
    status: () => ({ armed, paused, toolCallId: pausedCall }),
    resume: () => { release?.(); },
    reconcile: async (id, read) => {
      const record = await load(id);
      const current = await read(structuredClone(record));
      const content = (values: WriteCellsGrid, formulas: WriteCellsGrid): WriteCellsGrid =>
        values.map((row, r) => row.map((value, c) => {
          const formula = formulas[r]?.[c];
          return typeof formula === "string" && formula.startsWith("=") ? formula : value ?? "";
        }));
      const actual = JSON.stringify(content(current.values, current.formulas));
      if (actual === JSON.stringify(content(record.intendedValues, record.intendedValues))) return "applied";
      if (actual === JSON.stringify(content(record.beforeValues, record.beforeFormulas))) return "unchanged";
      return "conflict";
    },
  };
}
