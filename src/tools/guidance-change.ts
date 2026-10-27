/** Consent and reversible feedback for changes to Pi's standing guidance. */
import { t } from "../language/index.js";
import { showActionToast, showToast } from "../ui/toast.js";
import type { ExperimentalToolGateDependencies } from "./experimental-tool-gates/types.js";

export type GuidanceChangeDependencies = Pick<ExperimentalToolGateDependencies,
  "getExecutionMode" | "requestLocalChangeApproval">;

export async function approveGuidanceChange(
  dependencies: GuidanceChangeDependencies,
  toolName: string,
  message: string,
  signal: AbortSignal | undefined,
  alwaysAsk = false,
): Promise<void> {
  signal?.throwIfAborted();
  const mode = await dependencies.getExecutionMode?.() ?? "safe";
  if (alwaysAsk || mode === "safe") {
    if (!dependencies.requestLocalChangeApproval) {
      throw new Error("This guidance change needs your approval, but the approval dialog is unavailable.");
    }
    const approved = await dependencies.requestLocalChangeApproval({
      toolName,
      title: t("guidance.approval.title"),
      message,
      confirmLabel: t("guidance.approval.confirm"),
      destructive: false,
      ...(signal ? { signal } : {}),
    });
    signal?.throwIfAborted();
    if (!approved) throw new Error("Guidance change cancelled by the user.");
  }
  signal?.throwIfAborted();
}

export function showGuidanceUndo(message: string, undo: () => Promise<void>): void {
  if (typeof document === "undefined") return;
  showActionToast({
    message,
    actionLabel: t("guidance.undo"),
    onAction: () => {
      void undo().then(
        () => showToast(t("guidance.undone")),
        () => showToast(t("guidance.undoFailed"), { variant: "error" }),
      );
    },
  });
}
