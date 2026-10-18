import type { MutationObservation, MutationReconciliation, MutationRecord } from "./durable/mutations.js";
import { mountDurableControls, type DurableControlsRuntime, type DurableHistoryCallbacks } from "./ui/durable-controls.js";

const workbookId = "workbook:gallery";
const sheetId = "{4F1C2B9E-7A31-4C5D-9E0B-2D6F8A1C3B57}";
const sheets = [{ id: sheetId, name: "Budget", position: 0, visibility: "Visible", active: true, tables: [] }];
const cellsLimitation = "Evidence covers target values/formulas, sheet/table metadata and up to eight adjacent cells per range. It does not prove unrelated sheet state or persistent row/column identity, and never authorizes replay.";

function cells(total: number): MutationObservation {
  return {
    workbookId, structure: { sheets, anchors: [{ range: "C3", content: [["Total"]] }] },
    state: [{ range: "C4:D4", content: [[total, "=C4*1.2"]] }], complete: true, limitation: cellsLimitation,
  };
}

/** Same shape the runtime persists for an interrupted write_cells call. */
const record: MutationRecord = {
  version: 3, mutationId: "mutation:7c1e0f4a-2b9d-4e63-a8f5-91d3c6b20e47", toolCallId: "toolu_01HXJ8Q2ZK7M", toolName: "write_cells",
  workbookId, sessionId: "session:0d6b9f2e", requiresReview: false, phase: "dispatched",
  intent: { start_cell: "Budget!C4", values: [[1200, "=C4*1.2"]] },
  target: { kind: "cells", sheetId, ranges: ["C4:D4"], chartId: null, action: "" },
  before: cells(1000), expected: cells(1200), after: null, resolution: null,
};

/** Deterministic gallery fixture; all state transitions still go through real buttons. */
export function mountDurableGallery(host: HTMLElement, scenario: "pending" | "conflict" | "error" | "resolved" | "storage"): void {
  let pending = scenario !== "resolved" && scenario !== "storage";
  let attempts = 0;
  let observation = 1;
  let deleted = false;
  const reconciliation = (): MutationReconciliation => scenario === "conflict"
    ? { classification: "conflict", reason: cellsLimitation, observation: cells(1160 + observation * 15), observationToken: `gallery-observation-${observation}` }
    : { classification: "applied", reason: "The complete intended postcondition matches at the recorded identities and structural context.",
      observation: cells(1200), observationToken: "gallery-completed-observation" };
  const unused = (): never => { throw new Error("Gallery must not execute runtime operations"); };
  const runtime: DurableControlsRuntime = {
    get harness() { return unused(); },
    reviews: () => Promise.resolve(!pending ? [] : scenario === "error"
      ? [{ mutationId: "legacy:prototype:3", toolName: "Legacy workbook change", legacy: true, evidence: { toolName: "fill_formula", range: "Budget!E2:E40" } }]
      : [{ mutationId: record.mutationId, toolName: record.toolName, legacy: false, evidence: record }]),
    mutations: {
      pending: unused, execute: unused, resultCommitted: unused, configureFault: unused, status: unused, resume: unused,
      reconcile: () => {
        if (scenario === "error") return Promise.reject(new Error("Stable workbook identity is unavailable"));
        return Promise.resolve(reconciliation());
      },
      resolve: (_id: string, token: string) => {
        attempts++;
        if (scenario === "conflict" && attempts === 1) observation++;
        if (token !== reconciliation().observationToken) return Promise.reject(new Error("Workbook state changed since review; review the new observation before resolving"));
        pending = false;
        return Promise.resolve(reconciliation());
      },
    },
  };
  const history: DurableHistoryCallbacks = {
    listScopes: () => Promise.resolve([
      { id: "current", name: "WPS", generation: "2", current: true, identityDurable: false },
      ...deleted ? [] : [{ id: "retained", name: "Budget.xlsx", generation: "1", current: false, identityDurable: true }],
    ]),
    inspectSessions: () => Promise.resolve([{ id: "completed", name: "Budget review", completed: true }, { id: "pending", name: "Uncertain change", completed: false }]),
    copyCompletedHistory: (_scope, session) => { if (session !== "completed") return Promise.reject(new Error("Pending fixture copied")); return Promise.resolve(); },
    deleteGeneration: () => { deleted = true; return Promise.resolve(); },
  };
  mountDurableControls(host, runtime, history);
}
