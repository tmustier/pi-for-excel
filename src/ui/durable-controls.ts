import type { openDurableRuntime } from "../durable/runtime.js";
import { t } from "../language/index.js";
import { createOverlayButton, createOverlayDialog, createOverlayHeader, type OverlayDialogController } from "./overlay-dialog.js";
import "./durable-controls.css";

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

function text(tag: string, value: string): HTMLElement {
  const node = document.createElement(tag);
  node.textContent = value;
  return node;
}

/** Mount outside model transcript. Parent refreshes after committed runtime changes. */
export function mountDurableControls(host: HTMLElement, runtime: DurableControlsRuntime, history: DurableHistoryCallbacks): DurableControls {
  const root = document.createElement("section");
  root.className = "pi-durable-controls";
  const pending = document.createElement("div");
  pending.setAttribute("role", "status");
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
      dialog.card.append(alert);
    }
  };
  const button = (label: string, action: () => Promise<void> | void, target: HTMLElement): HTMLButtonElement => {
    const node = createOverlayButton({ text: label });
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
      const items = [...current.card.querySelectorAll<HTMLButtonElement>("button:not(:disabled)")];
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
  const inspect = async (review: Awaited<ReturnType<DurableControlsRuntime["reviews"]>>[number], current: OverlayDialogController): Promise<void> => {
    const version = ++inspectionVersion;
    current.card.querySelector(".pi-durable-inspection")?.remove();
    const content = text("div", "");
    content.className = "pi-durable-inspection";
    content.append(text("h3", t("durable.evidence")), text("pre", JSON.stringify(review.evidence, null, 2)), text("p", t("durable.uncertainty")));
    current.card.append(content);
    current.card.scrollTop = 0;
    const observation = await runtime.mutations.reconcile(review.mutationId);
    if (!current.overlay.isConnected || version !== inspectionVersion) return;
    content.append(text("h3", t("durable.observation")), text("pre", JSON.stringify(observation, null, 2)), text("p", t("durable.noReplay")));
    button(t("durable.acknowledge"), async () => {
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
        current.card.append(text("p", t("durable.reviewAgain")));
        await inspect(review, current);
      }
    }, content);
    current.card.scrollTop = 0;
  };
  const refresh = async (): Promise<void> => {
    const version = ++refreshVersion;
    try {
      const reviews = await runtime.reviews();
      if (disposed || version !== refreshVersion) return;
      pending.replaceChildren(text("p", reviews.length ? t("durable.pending", { count: reviews.length }) : t("durable.none")));
      for (const review of reviews) {
        const row = text("div", review.toolName);
        button(t("durable.inspect"), async () => {
          const current = open(t("durable.review"));
          current.card.append(text("p", review.mutationId));
          button(t("durable.inspect"), () => inspect(review, current), current.card);
          await inspect(review, current);
        }, row);
        pending.append(row);
      }
    } catch (cause) { fail(cause); }
  };
  root.append(pending, notice, error);
  button(t("durable.refresh"), refresh, root);
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
  }, root);
  host.append(root);
  pending.append(text("p", t("durable.loading")));
  void refresh();
  return { refresh, dispose: () => { disposed = true; dialog?.close(); root.remove(); } };
}
