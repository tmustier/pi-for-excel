import { Type } from "typebox";
import { Value } from "typebox/value";
import { parseMutationRecord, type MutationObservation, type MutationReconciliation, type MutationRecord } from "../durable/mutations.js";
import type { MutationReview, openDurableRuntime } from "../durable/runtime.js";
import { t } from "../language/index.js";
import { createOverlayButton, createOverlayDialog, createOverlayHeader, type OverlayDialogController } from "./overlay-dialog.js";

export type DurableControlsRuntime = Pick<Awaited<ReturnType<typeof openDurableRuntime>>, "reviews" | "mutations" | "harness">;
export interface RetainedWorkbookScope {
  id: string;
  generation: string;
  name: string;
  current: boolean;
  identityDurable: boolean;
}
export interface RetainedSessionMetadata {
  id: string;
  name: string;
  completed: boolean;
}
/** Implementations must only copy completed history, never pending execution. */
export interface DurableHistoryCallbacks {
  listScopes(): Promise<RetainedWorkbookScope[]>;
  inspectSessions(scopeId: string): Promise<RetainedSessionMetadata[]>;
  copyCompletedHistory(scopeId: string, sessionId: string): Promise<void>;
  deleteGeneration(scopeId: string): Promise<void>;
}
export interface DurableControls {
  refresh(): Promise<void>;
  dispose(): void;
}

function text(tag: string, value: string, className?: string): HTMLElement {
  const node = document.createElement(tag);
  node.textContent = value;
  if (className) node.className = className;
  return node;
}

// Read-only display views of the mutation host's persisted shapes. Anything that
// does not match stays out of the summary and remains in the technical evidence.
const cellStateSchema = Type.Array(Type.Object({
  range: Type.String(),
  content: Type.Array(Type.Array(Type.Union([Type.String(), Type.Number(), Type.Boolean(), Type.Null()]))),
}), { minItems: 1 });
const targetSchema = Type.Object({ sheetId: Type.Union([Type.String(), Type.Null()]), ranges: Type.Array(Type.String()), action: Type.String() });
const sheetsSchema = Type.Array(Type.Object({ id: Type.String(), name: Type.String() }));
const cellStructureSchema = Type.Object({ sheets: sheetsSchema });
const SUMMARY_CELLS = 6;

function parseRecord(review: MutationReview): MutationRecord | null {
  if (review.legacy) return null;
  try { return parseMutationRecord(review.evidence); } catch { return null; }
}

function operationLabel(review: MutationReview, record: MutationRecord | null): string {
  if (!record) return review.legacy ? t("durable.legacy") : review.toolName;
  const action = Value.Check(targetSchema, record.target) ? record.target.action : "";
  return action ? `${record.toolName} · ${action}` : record.toolName;
}

/** Sheet names come from the identities recorded with the original state, never from a name lookup. */
function targetLabel(record: MutationRecord | null): string {
  if (!record || !Value.Check(targetSchema, record.target) || !record.target.ranges.length) return t("durable.targetUnknown");
  const { sheetId, ranges } = record.target;
  const structure = record.before.structure;
  const sheets = Value.Check(sheetsSchema, structure) ? structure : Value.Check(cellStructureSchema, structure) ? structure.sheets : [];
  const sheet = sheets.find((entry) => entry.id === sheetId)?.name;
  return ranges.map((range) => sheet ? `${sheet}!${range}` : range).join(", ");
}

function stageLabel(record: MutationRecord | null): string | null {
  if (!record) return null;
  if (record.requiresReview) return t("durable.stageError");
  if (record.phase === "prepared") return t("durable.stagePrepared");
  if (record.phase === "dispatched") return t("durable.stageDispatched");
  if (record.phase === "returned") return t("durable.stageReturned");
  return null;
}

function cellValues(observation: MutationObservation | null | undefined): string | null {
  const state = observation?.state;
  if (!Value.Check(cellStateSchema, state)) return null;
  return state.map(({ range, content }) => {
    const cells = content.flat();
    const shown = cells.slice(0, SUMMARY_CELLS).map((value) => value === "" || value === null ? t("durable.emptyCell") : String(value));
    const more = cells.length > SUMMARY_CELLS ? ` ${t("durable.moreCells", { count: cells.length })}` : "";
    return `${range}: ${shown.join(", ")}${more}`;
  }).join("\n");
}

function classificationText(classification: MutationReconciliation["classification"]): { label: string; help: string } {
  if (classification === "applied") return { label: t("durable.classApplied"), help: t("durable.classAppliedHelp") };
  if (classification === "unchanged") return { label: t("durable.classUnchanged"), help: t("durable.classUnchangedHelp") };
  return { label: t("durable.classConflict"), help: t("durable.classConflictHelp") };
}

function facts(rows: [string, string | null][]): HTMLElement {
  const list = document.createElement("dl");
  list.className = "pi-durable-facts";
  for (const [label, value] of rows) {
    if (value === null) continue;
    list.append(text("dt", label), text("dd", value));
  }
  return list;
}

/** Mount outside model transcript. Parent refreshes after committed runtime changes. */
export function mountDurableControls(host: HTMLElement, runtime: DurableControlsRuntime, history: DurableHistoryCallbacks): DurableControls {
  const root = document.createElement("section");
  root.className = "pi-durable-controls";
  const pending = document.createElement("div");
  pending.className = "pi-durable-status";
  pending.setAttribute("role", "status");
  const list = text("div", "", "pi-durable-list");
  const actions = text("div", "", "pi-durable-actions");
  const error = text("p", "");
  error.setAttribute("role", "alert");
  const notice = text("p", "");
  notice.setAttribute("role", "status");
  let dialog: OverlayDialogController | null = null;
  let disposed = false;
  let refreshVersion = 0;
  const fail = (cause: unknown): void => {
    error.textContent = t("durable.error", { error: cause instanceof Error ? cause.message : String(cause) });
    if (dialog?.overlay.isConnected) {
      dialog.card.querySelector("[role=alert]")?.remove();
      const alert = text("p", error.textContent);
      alert.setAttribute("role", "alert");
      (dialog.card.querySelector(".pi-durable-alerts") ?? dialog.card).append(alert);
    }
  };
  const button = (label: string, action: () => Promise<void> | void, target: HTMLElement, className?: string): HTMLButtonElement => {
    const node = createOverlayButton(className ? { text: label, className } : { text: label });
    node.addEventListener("click", () => {
      node.disabled = true;
      void Promise.resolve().then(action).catch(fail).finally(() => { node.disabled = false; });
    });
    target.append(node);
    return node;
  };
  const open = (title: string): OverlayDialogController => {
    dialog?.close();
    const previous = document.activeElement;
    const current = createOverlayDialog({ overlayId: "pi-durable-dialog", cardClassName: "pi-durable-dialog", restoreFocusOnClose: false });
    dialog = current;
    current.overlay.setAttribute("aria-label", title);
    const header = createOverlayHeader({ title, onClose: current.close, closeLabel: t("durable.close") });
    current.card.append(header.header);
    current.addCleanup(() => {
      if (previous instanceof HTMLElement && previous.isConnected) previous.focus();
      else root.querySelector<HTMLButtonElement>("button:not(:disabled)")?.focus();
    });
    current.overlay.addEventListener("keydown", (event) => {
      if (event.key !== "Tab") return;
      const items = [...current.card.querySelectorAll<HTMLElement>("button:not(:disabled), summary")];
      const first = items[0];
      const last = items.at(-1);
      if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus(); }
      else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus(); }
    });
    current.mount();
    header.closeButton.focus();
    return current;
  };
  let inspectionVersion = 0;
  const inspect = async (review: MutationReview, record: MutationRecord | null, current: OverlayDialogController): Promise<void> => {
    const version = ++inspectionVersion;
    current.card.querySelector(".pi-durable-inspection")?.remove();
    const content = text("div", "", "pi-durable-inspection");
    const outcome = text("div", "", "pi-durable-outcome");
    outcome.append(text("p", t("durable.checking"), "pi-durable-muted"));
    const values = facts([
      [t("durable.valuesBefore"), cellValues(record?.before)],
      [t("durable.valuesExpected"), cellValues(record?.expected)],
    ]);
    const reviewActions = text("div", "", "pi-durable-review-actions");
    // Forensic data stays available but collapsed: the decision must not require reading it.
    const evidence = document.createElement("details");
    evidence.className = "pi-durable-evidence";
    evidence.append(text("summary", t("durable.technical")), text("p", t("durable.recordId", { id: review.mutationId })),
      text("h4", t("durable.evidence")), text("pre", JSON.stringify(review.evidence, null, 2)), text("p", t("durable.uncertainty")));
    content.append(outcome, values, reviewActions, evidence);
    (current.card.querySelector(".pi-durable-review") ?? current.card).append(content);
    button(t("durable.recheck"), () => inspect(review, record, current), reviewActions);
    current.card.scrollTop = 0;
    let observation: MutationReconciliation;
    try {
      observation = await runtime.mutations.reconcile(review.mutationId);
    } catch (cause) {
      if (version === inspectionVersion) outcome.replaceChildren(text("p", t("durable.checkFailed")));
      throw cause;
    }
    if (!current.overlay.isConnected || version !== inspectionVersion) return;
    const { label, help } = classificationText(observation.classification);
    const heading = text("h3", label, "pi-durable-outcome-title");
    heading.tabIndex = -1;
    outcome.replaceChildren(heading, text("p", help), text("p", t("durable.reason", { reason: observation.reason }), "pi-durable-muted"));
    const limitation = observation.observation.limitation;
    if (limitation && limitation !== observation.reason) outcome.append(text("p", t("durable.limits", { limits: limitation }), "pi-durable-muted"));
    values.append(...facts([
      [t("durable.valuesCurrent"), cellValues(observation.observation)],
      [t("durable.coverage"), observation.observation.complete ? t("durable.coverageComplete") : t("durable.coveragePartial")],
    ]).childNodes);
    content.insertBefore(text("p", t("durable.noReplay"), "pi-durable-warning"), reviewActions);
    evidence.append(text("h4", t("durable.observation")), text("pre", JSON.stringify(observation, null, 2)));
    const acknowledge = button(t("durable.acknowledge"), async () => {
      // Invalidate acceptance before awaiting: every failure requires fresh inspection.
      content.remove();
      try {
        await runtime.mutations.resolve(review.mutationId, observation.observationToken);
        error.textContent = "";
        notice.textContent = t("durable.acknowledged");
        await refresh();
        current.close();
      } catch (cause) {
        fail(cause);
        const alerts = current.card.querySelector(".pi-durable-alerts") ?? current.card;
        alerts.querySelector(".pi-durable-review-again")?.remove();
        alerts.append(text("p", t("durable.reviewAgain"), "pi-durable-review-again"));
        await inspect(review, record, current);
        // Land on the new result, never on the new acknowledgement button.
        current.card.querySelector<HTMLElement>(".pi-durable-outcome-title")?.focus();
      }
    }, reviewActions, "pi-overlay-btn--primary");
    reviewActions.prepend(acknowledge);
    current.card.scrollTop = 0;
  };
  const review = async (item: MutationReview): Promise<void> => {
    const current = open(t("durable.review"));
    const record = parseRecord(item);
    const body = text("div", "", "pi-durable-review");
    body.append(facts([
      [t("durable.operation"), operationLabel(item, record)],
      [t("durable.target"), targetLabel(record)],
      [t("durable.stage"), stageLabel(record)],
    ]), text("div", "", "pi-durable-alerts"));
    current.card.append(body);
    await inspect(item, record, current);
  };
  const refresh = async (): Promise<void> => {
    const version = ++refreshVersion;
    try {
      const reviews = await runtime.reviews();
      if (disposed || version !== refreshVersion) return;
      root.classList.toggle("pi-durable-controls--pending", reviews.length > 0);
      pending.replaceChildren(text("p", reviews.length ? t("durable.pending", { count: reviews.length }) : t("durable.none")));
      list.replaceChildren();
      for (const item of reviews) {
        const record = parseRecord(item);
        const row = text("div", "", "pi-durable-row");
        row.append(text("span", `${operationLabel(item, record)} · ${targetLabel(record)}`));
        button(t("durable.inspect"), () => review(item), row, "pi-overlay-btn--compact");
        list.append(row);
      }
    } catch (cause) { fail(cause); }
  };
  root.append(pending, list, notice, error, actions);
  button(t("durable.refresh"), refresh, actions, "pi-overlay-btn--compact");
  button(t("durable.storage"), async () => {
    const current = open(t("durable.storage"));
    current.card.append(text("p", t("durable.transferHelp")));
    const scopes = await history.listScopes();
    for (const scope of scopes) {
      const row = text("section", "");
      row.append(text("h3", scope.name), text("p", t("durable.scope", { generation: scope.generation, current: scope.current ? t("durable.current") : t("durable.retained"), identity: scope.identityDurable ? t("durable.durableIdentity") : t("durable.unidentified") })));
      button(t("durable.sessions"), async () => {
        const sessions = await history.inspectSessions(scope.id);
        row.querySelector(".pi-durable-sessions")?.remove();
        const list = text("div", "");
        list.className = "pi-durable-sessions";
        if (!sessions.length) list.append(text("p", t("durable.noSessions")));
        for (const session of sessions) {
          const item = text("div", session.name);
          item.append(text("p", session.id));
          item.append(text("p", session.completed ? t("durable.completed") : t("durable.pendingHistory")));
          if (session.completed && !scope.current) button(t("durable.copy"), async () => {
            await history.copyCompletedHistory(scope.id, session.id);
            notice.textContent = t("durable.copied");
            item.querySelector("[role=status]")?.remove();
            const copied = text("p", t("durable.copied"));
            copied.setAttribute("role", "status");
            item.append(copied);
          }, item);
          list.append(item);
        }
        row.append(list);
      }, row);
      if (!scope.current) button(t("durable.delete"), () => {
        const previous = document.activeElement;
        const confirmation = text("div", "");
        confirmation.append(text("p", t("durable.deleteConfirm", { name: scope.name, generation: scope.generation })));
        const cancel = button(t("durable.cancel"), () => { confirmation.remove(); if (previous instanceof HTMLElement) previous.focus(); }, confirmation);
        button(t("durable.confirmDelete"), async () => {
          await history.deleteGeneration(scope.id);
          row.remove();
          current.card.querySelector<HTMLButtonElement>("button")?.focus();
        }, confirmation);
        row.querySelector(".pi-durable-confirm")?.remove();
        confirmation.className = "pi-durable-confirm";
        row.append(confirmation);
        cancel.focus();
      }, row);
      current.card.append(row);
    }
  }, actions, "pi-overlay-btn--compact");
  host.append(root);
  pending.append(text("p", t("durable.loading")));
  void refresh();
  return { refresh, dispose: () => { disposed = true; dialog?.close(); root.remove(); } };
}
