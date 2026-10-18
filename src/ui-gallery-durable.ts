import type { MutationReconciliation } from "./durable/mutations.js";
import { mountDurableControls, type DurableControlsRuntime, type DurableHistoryCallbacks } from "./ui/durable-controls.js";
import { t } from "./language/index.js";

/** Deterministic gallery fixture; all state transitions still go through real buttons. */
export function mountDurableGallery(host: HTMLElement, scenario: "pending" | "conflict" | "error" | "resolved" | "storage"): void {
  let pending = scenario !== "resolved" && scenario !== "storage";
  let attempts = 0;
  let observation = 1;
  let deleted = false;
  const reconciliation = (): MutationReconciliation => ({
    classification: scenario === "conflict" ? "conflict" : "unchanged",
    reason: t("durable.uncertainty"),
    observation: { workbookId: "gallery", structure: { sheet: "Sheet1" }, state: { A1: observation }, complete: false, limitation: t("durable.uncertainty") },
    observationToken: `observation-${observation}`,
  });
  const unused = (): never => { throw new Error("Gallery must not execute runtime operations"); };
  const runtime: DurableControlsRuntime = {
    get harness() { return unused(); },
    reviews: () => Promise.resolve(pending ? [{ mutationId: scenario, toolName: "write_cells", legacy: true, evidence: { before: { A1: 0 }, expected: { A1: 1 }, uncertainty: t("durable.uncertainty") } }] : []),
    mutations: {
      pending: unused, execute: unused, resultCommitted: unused, configureFault: unused, status: unused, resume: unused,
      reconcile: () => {
        if (scenario === "error") return Promise.reject(new Error(t("durable.loading")));
        return Promise.resolve(reconciliation());
      },
      resolve: (_id: string, token: string) => {
        attempts++;
        if (scenario === "conflict" && attempts === 1) observation++;
        if (token !== reconciliation().observationToken) return Promise.reject(new Error(t("durable.reviewAgain")));
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
