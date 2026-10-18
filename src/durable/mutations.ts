import type { JsonValue } from "@earendil-works/chord";
import { Type, type Static } from "typebox";
import { Value } from "typebox/value";
import type { WorkbookCoordinator } from "../workbook/coordinator.js";
import type { WorkbookMutationCall, WorkbookMutationGate } from "../tools/mutation/gate.js";
import { decodeToolDetails } from "../tools/tool-details.js";

export type MutationFaultPoint = "beforeDispatch" | "afterDispatch" | "beforeResult";
export type MutationClassification = "applied" | "unchanged" | "conflict";

/** Safe execution diagnostic for runtime/UI. Never carries the raw host/storage error or cause. */
export class MutationExecutionError extends Error {
  readonly mutationId: string;
  readonly outcome: "not-executed" | "uncertain";

  constructor(mutationId: string, outcome: MutationExecutionError["outcome"]) {
    super(outcome === "uncertain"
      ? "Workbook mutation may already have applied. Review the pending mutation before making further changes; do not retry automatically."
      : "Workbook mutation was not executed: host execution was never entered. Review any pending intent before making further changes.");
    this.name = "MutationExecutionError";
    this.mutationId = mutationId;
    this.outcome = outcome;
  }
}

const jsonSchema = Type.Cyclic({ Json: Type.Union([
  Type.String(), Type.Number(), Type.Boolean(), Type.Null(),
  Type.Array(Type.Ref("Json")), Type.Record(Type.String(), Type.Ref("Json")),
]) }, "Json");
const observationSchema = Type.Object({
  workbookId: Type.String({ minLength: 1 }), structure: jsonSchema, state: jsonSchema,
  complete: Type.Boolean(), limitation: Type.Union([Type.String(), Type.Null()]),
});
/** Evidence is observational, not an undo snapshot or a command to replay. */
export type MutationObservation = Static<typeof observationSchema>;
const planSchema = Type.Object({
  intent: jsonSchema, target: jsonSchema, before: observationSchema,
  expected: Type.Union([observationSchema, Type.Null()]),
});
export type MutationPlan = Static<typeof planSchema>;
const reconciliationSchema = Type.Object({
  classification: Type.Union([Type.Literal("applied"), Type.Literal("unchanged"), Type.Literal("conflict")]),
  reason: Type.String(), observation: observationSchema, observationToken: Type.String({ minLength: 1 }),
});
export type MutationReconciliation = Static<typeof reconciliationSchema>;

/** The runtime may use this parser at its persisted document boundary. Invalid records block writes. */
export const mutationRecordSchema = Type.Intersect([planSchema, Type.Object({
  version: Type.Literal(3), mutationId: Type.String({ minLength: 1 }),
  toolCallId: Type.String({ minLength: 1 }), toolName: Type.String({ minLength: 1 }),
  workbookId: Type.String({ minLength: 1 }), sessionId: Type.String(),
  /** Failed receipt, not incomplete observation; remains pending even after result commit. */
  requiresReview: Type.Boolean(),
  phase: Type.Union([Type.Literal("prepared"), Type.Literal("dispatched"), Type.Literal("returned"),
    Type.Literal("resultCommitted"), Type.Literal("resolved")]),
  after: Type.Union([observationSchema, Type.Null()]),
  resolution: Type.Union([reconciliationSchema, Type.Null()]),
})]);
export type MutationRecord = Static<typeof mutationRecordSchema>;

export function parseMutationRecord(value: JsonValue): MutationRecord {
  if (!Value.Check(mutationRecordSchema, value)) throw new Error("Invalid or incompatible mutation record; explicit recovery is required");
  if ((value.phase === "resolved" && !value.resolution)
    || ((value.phase === "returned" || value.phase === "resultCommitted") && !value.after)
    || value.before.workbookId !== value.workbookId
    || (value.expected && value.expected.workbookId !== value.workbookId)
    || (value.after && value.after.workbookId !== value.workbookId)) {
    throw new Error("Inconsistent mutation record; explicit recovery is required");
  }
  return value;
}

export interface MutationHost {
  prepare(call: WorkbookMutationCall): Promise<MutationPlan>;
  /** Must be read-only and resolve targets by stable identity, not their former names. */
  observe(record: MutationRecord): Promise<MutationObservation>;
}

export interface MutationStore {
  load(mutationId: string): Promise<MutationRecord | null>;
  /** All records for this workbook, including other sessions; never truncate or apply retention. */
  list(): Promise<MutationRecord[]>;
  /** Resolve only after durable commit. Failed reads/writes must reject, never look absent. */
  save(record: MutationRecord): Promise<void>;
}

export interface MutationController extends WorkbookMutationGate {
  pending(): Promise<MutationRecord[]>;
  reconcile(mutationId: string): Promise<MutationReconciliation>;
  /** Explicit human acceptance of current state. Does not write the workbook or replay the call. */
  resolve(mutationId: string, observationToken: string): Promise<MutationReconciliation>;
  /** Invoke only AFTER the tool result is committed; never from a tool-finally block. */
  resultCommitted(mutationId: string): Promise<void>;
  configureFault(point: MutationFaultPoint | null): void;
  status(): { armed: MutationFaultPoint | null; paused: MutationFaultPoint | null; mutationId: string | null; toolCallId: string | null };
  resume(): void;
}

function canonical(value: JsonValue): string {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map(canonical).join(",")}]`;
  return `{${Object.keys(value).sort().map((key) => `${JSON.stringify(key)}:${canonical(value[key] ?? null)}`).join(",")}}`;
}

function observationValue(observation: MutationObservation): JsonValue {
  return { ...observation };
}

function same(left: MutationObservation, right: MutationObservation): boolean {
  return canonical(observationValue(left)) === canonical(observationValue(right));
}

function isPending(record: MutationRecord): boolean {
  return record.phase !== "resolved" && (record.phase !== "resultCommitted" || record.requiresReview);
}

/** One controller per workbook, sharing the existing coordinator with every mutation producer. */
export function createMutationController(options: {
  workbookId: string;
  store: MutationStore;
  host: MutationHost;
  coordinator: WorkbookCoordinator;
  allowFaultInjection?: boolean;
}): MutationController {
  if (!options.workbookId || options.workbookId === "workbook:unknown") {
    throw new Error("Durable mutations require a stable workbook identity");
  }
  let armed: MutationFaultPoint | null = null;
  let paused: MutationFaultPoint | null = null;
  let pausedCall: string | null = null;
  let pausedMutation: string | null = null;
  let release: (() => void) | null = null;
  const active = new Set<string>();

  const pause = async (point: MutationFaultPoint, call: WorkbookMutationCall): Promise<void> => {
    if (armed !== point) return;
    armed = null;
    paused = point;
    pausedCall = call.toolCallId;
    pausedMutation = call.mutationId;
    await new Promise<void>((resolve) => { release = resolve; });
    paused = null;
    pausedCall = null;
    pausedMutation = null;
    release = null;
  };
  const load = async (id: string): Promise<MutationRecord> => {
    const record = parseMutationRecord(await options.store.load(id));
    if (record.mutationId !== id || record.workbookId !== options.workbookId) {
      throw new Error("Mutation record unavailable for workbook");
    }
    return structuredClone(record);
  };
  const pending = async (): Promise<MutationRecord[]> => {
    const records = (await options.store.list()).map(parseMutationRecord);
    // A mis-scoped store is an integration error, not permission to ignore another workbook.
    if (records.some((record) => record.workbookId !== options.workbookId)) {
      throw new Error("Mutation store contains incompatible or cross-workbook records");
    }
    return structuredClone(records.filter(isPending));
  };
  const save = (record: MutationRecord): Promise<void> => options.store.save(structuredClone(parseMutationRecord(record)));
  const reconcile = async (id: string): Promise<MutationReconciliation> => {
    const record = await load(id);
    const observation = await options.host.observe(structuredClone(record));
    if (observation.workbookId !== options.workbookId) throw new Error("Workbook identity changed during reconciliation");
    let classification: MutationClassification = "conflict";
    let reason = observation.limitation ?? "Current state does not prove the complete intended outcome or original state.";
    if (record.phase === "prepared") {
      classification = observation.complete && record.before.complete && same(observation, record.before) ? "unchanged" : "conflict";
      reason = "The execution boundary was never entered; this call made no workbook changes. "
        + (classification === "unchanged" ? "Current state matches the original state."
          : "Current state differs from the original state or cannot be completely observed.");
    } else if (observation.complete && record.expected?.complete && same(observation, record.expected)) {
      classification = "applied";
      reason = "The complete intended postcondition matches at the recorded identities and structural context.";
    } else if (observation.complete && record.before.complete && same(observation, record.before)) {
      classification = "unchanged";
      reason = "The complete observed scope matches the original state; this is not permission to retry.";
    }
    const bytes = new TextEncoder().encode(canonical({
      id, phase: record.phase, observation: observationValue(observation), classification,
    }));
    const digest = await crypto.subtle.digest("SHA-256", bytes);
    const observationToken = Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
    return { classification, reason, observation: structuredClone(observation), observationToken };
  };
  const serialized = async <T>(id: string, action: () => Promise<T>): Promise<T> => {
    const out = await options.coordinator.runWrite({
      workbookId: options.workbookId, sessionId: "mutation-resolution", opId: id, journalOnly: true,
    }, action);
    return out.result;
  };

  return {
    // Called INSIDE withWorkbookCoordinator's write lock, never recursively acquire that lock.
    execute: async (call, run) => {
      if (call.workbookId !== options.workbookId) throw new Error("Workbook identity changed before mutation");
      if (!call.mutationId?.trim()) throw new Error("Durable mutation requires a stable logical mutationId");
      if (await options.store.load(call.mutationId)) throw new Error("Mutation already recorded; automatic replay is forbidden");
      if ((await pending()).length) throw new Error("Resolve pending workbook mutations before making another change");
      let hostEntered = false;
      active.add(call.mutationId);
      try {
        const plan = await options.host.prepare(call);
        if (plan.before.workbookId !== options.workbookId) throw new Error("Workbook identity changed while preparing mutation");
        let record: MutationRecord = structuredClone({
          ...plan, version: 3, mutationId: call.mutationId, toolCallId: call.toolCallId, toolName: call.toolName,
          workbookId: options.workbookId, sessionId: call.sessionId,
          phase: "prepared", after: null, resolution: null, requiresReview: false,
        });
        await save(record);
        await pause("beforeDispatch", call);
        const current = await options.host.observe(structuredClone(record));
        if (!same(current, record.before)) {
          throw new Error("Workbook changed while mutation intent was being committed; review before continuing");
        }
        // Commit before even entering host code: it may mutate synchronously (WPS/extensions).
        record = { ...record, phase: "dispatched" };
        await save(record);
        hostEntered = true;
        const result = await run();
        await pause("afterDispatch", call);
        const after = await options.host.observe(structuredClone(record));
        if (after.workbookId !== options.workbookId) throw new Error("Workbook identity changed after mutation");
        // Existing core tools return error receipts rather than throwing. A host sync may
        // already have applied part of the operation, so a committed error must not unblock retries.
        const details = decodeToolDetails(result.details);
        const structuredError = details !== undefined && (
          ("ok" in details && details.ok === false)
          || ("error" in details && details.error !== undefined)
          || ("blocked" in details && details.blocked)
          || details.kind === "skills_error"
        );
        const errorReceipt = structuredError || result.content.some((block) => block.type === "text" && /^\s*(?:⚠️\s*)?Error\b/iu.test(block.text));
        // Incomplete observation is not failed execution. Only a failed receipt
        // keeps a normally returned call pending after its result is durably committed.
        record = { ...record, phase: "returned", after, requiresReview: errorReceipt };
        await save(record);
        await pause("beforeResult", call);
        return result;
      } catch {
        // Do not replace the durable intent, roll back, or expose exception details.
        // A failed commit may itself have committed; leave that evidence untouched.
        throw new MutationExecutionError(call.mutationId, hostEntered ? "uncertain" : "not-executed");
      } finally {
        active.delete(call.mutationId);
      }
    },
    pending,
    reconcile,
    resolve: (id, observationToken) => serialized(id, async () => {
      if (active.has(id)) throw new Error("Cannot resolve a running mutation");
      const record = await load(id);
      if (!isPending(record)) throw new Error("Mutation is already settled");
      const resolution = await reconcile(id);
      if (resolution.observationToken !== observationToken) {
        throw new Error("Workbook state changed since review; review the new observation before resolving");
      }
      await save({ ...record, phase: "resolved", resolution });
      return resolution;
    }),
    resultCommitted: (id) => serialized(id, async () => {
      const record = await load(id);
      if (record.phase === "resultCommitted" || record.phase === "resolved") return;
      if (record.phase !== "returned" || active.has(id)) throw new Error("Mutation has not returned a result");
      await save({ ...record, phase: "resultCommitted" });
    }),
    configureFault: (point) => {
      if (point !== null && options.allowFaultInjection !== true) throw new Error("Fault injection is disabled");
      if (paused !== null) throw new Error("Resume the paused mutation first");
      armed = point;
    },
    status: () => ({ armed, paused, mutationId: pausedMutation, toolCallId: pausedCall }),
    resume: () => { release?.(); },
  };
}
