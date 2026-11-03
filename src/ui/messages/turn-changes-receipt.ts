import { html, type TemplateResult } from "lit";

import { t } from "../../language/index.js";
import { describeOfficeJsChanges, type TurnChangesMessage } from "../../messages/turn-changes.js";

export function renderTurnChangesReceipt(message: TurnChangesMessage, placement: "footer" | "standalone"): TemplateResult {
  const hasBackups = message.changes.some((change) => change.backupIds.some((id) => !message.undoneBackupIds?.includes(id)));
  const partial = message.changes.some((change) => change.backupIds.length === 0);
  const cellsOnly = !message.changes.some((change) => change.unit === "changes");
  const count = cellsOnly ? message.changes.reduce((total, change) => total + change.changedCount, 0) : message.changes.length;
  const summary = cellsOnly
    ? t(count === 1 ? "turnChanges.summaryOne" : "turnChanges.summary", { count })
    : t(count === 1 ? "turnChanges.changesOne" : "turnChanges.changes", { count });
  const action = (name: "review" | "undo"): void => {
    document.dispatchEvent(new CustomEvent("pi:turn-change-action", { detail: { action: name, summary: message } }));
  };
  return html`<div class="pi-turn-changes pi-turn-changes--${placement}" data-outcome=${message.outcome} role="status">
    ${message.changes.length > 0 ? html`<div class="pi-turn-changes__line">
      <span class="pi-turn-changes__summary">${summary}</span>
      <span class="pi-turn-changes__actions">
        <button type="button" class="pi-btn pi-btn--text" ?disabled=${!hasBackups} @click=${() => action("undo")}>${t(message.undoneBackupIds?.length && !hasBackups ? "turnChanges.undone" : "turnChanges.undo")}</button>
        <button type="button" class="pi-btn pi-btn--secondary" @click=${() => action("review")}>${t("turnChanges.review")}</button>
      </span>
    </div>` : ""}
    ${partial ? html`<span class="pi-turn-changes__warning">${t("turnChanges.partial")}</span>` : ""}
    ${message.ranOfficeJs ? html`<span class="pi-turn-changes__warning">${describeOfficeJsChanges(message)}</span>` : ""}
  </div>`;
}
