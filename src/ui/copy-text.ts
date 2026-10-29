import { t } from "../language/index.js";
import {
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
} from "./overlay-dialog.js";
import { COPY_TEXT_OVERLAY_ID } from "./overlay-ids.js";

export type CopyTextOutcome = "copied" | "manual" | "unavailable";

export function copyViaHiddenTextarea(text: string): boolean {
  const textarea = document.createElement("textarea");
  textarea.value = text;
  textarea.setAttribute("readonly", "");
  textarea.style.position = "fixed";
  textarea.style.opacity = "0";
  document.body.appendChild(textarea);
  textarea.select();
  try {
    return document.execCommand("copy");
  } catch {
    return false;
  } finally {
    textarea.remove();
  }
}

export async function copyText(
  text: string,
  options: { manualFallback?: boolean } = {},
): Promise<CopyTextOutcome> {
  try {
    await navigator.clipboard.writeText(text);
    return "copied";
  } catch {
  }

  if (copyViaHiddenTextarea(text)) return "copied";
  if (options.manualFallback === false) return "unavailable";

  const dialog = createOverlayDialog({
    overlayId: COPY_TEXT_OVERLAY_ID,
    cardClassName: "pi-welcome-card pi-overlay-card pi-overlay-card--m",
  });

  const { header } = createOverlayHeader({
    onClose: dialog.close,
    closeLabel: t("dialog.close"),
    title: t("copyText.title"),
    subtitle: t("copyText.hint"),
  });

  const textarea = document.createElement("textarea");
  textarea.className = "pi-overlay-textarea pi-copy-text__value";
  textarea.readOnly = true;
  textarea.value = text;
  textarea.rows = 10;
  textarea.setAttribute("aria-label", t("copyText.title"));

  const body = document.createElement("div");
  body.className = "pi-overlay-body";
  body.appendChild(textarea);

  const actions = document.createElement("div");
  actions.className = "pi-overlay-actions";
  const done = createOverlayButton({ text: t("copyText.done"), className: "pi-overlay-btn--primary" });
  done.addEventListener("click", dialog.close);
  actions.appendChild(done);

  dialog.card.append(header, body, actions);
  dialog.mount();
  textarea.focus();
  textarea.select();
  return "manual";
}
