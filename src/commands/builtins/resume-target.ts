/**
 * Resume target semantics shared across commands + overlays.
 */

export type ResumeDialogTarget = "new_tab" | "replace_current";

import { t } from "../../language/index.js";

export function getResumeTargetLabel(target: ResumeDialogTarget): string {
  if (target === "replace_current") {
    return t("resume.replaceCurrent");
  }

  return t("resume.openInNewTab");
}

export function getCrossWorkbookResumeConfirmMessage(target: ResumeDialogTarget): string {
  if (target === "replace_current") {
    return t("resume.crossWorkbookReplace");
  }

  return t("resume.crossWorkbookNewTab");
}
