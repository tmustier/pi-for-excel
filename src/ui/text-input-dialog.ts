import {
  closeOverlayById,
  createOverlayButton,
  createOverlayDialog,
  createOverlayHeader,
  createOverlayInput,
} from "./overlay-dialog.js";
import { TEXT_INPUT_DIALOG_OVERLAY_ID } from "./overlay-ids.js";

const TEXT_INPUT_UI_UNAVAILABLE_ERROR =
  "Text input UI is unavailable in this environment.";

export interface TextInputDialogOptions {
  title: string;
  message?: string;
  confirmLabel?: string;
  cancelLabel?: string;
  placeholder?: string;
  initialValue?: string;
  overlayId?: string;
  restoreFocusOnClose?: boolean;
  cardClassName?: string;
}

function canRenderTextInputDialog(): boolean {
  if (typeof document === "undefined") {
    return false;
  }

  return document.body instanceof HTMLElement;
}

export function requestTextInputDialog(options: TextInputDialogOptions): Promise<string | null> {
  if (!canRenderTextInputDialog()) {
    return Promise.reject(new Error(TEXT_INPUT_UI_UNAVAILABLE_ERROR));
  }

  const overlayId = options.overlayId ?? TEXT_INPUT_DIALOG_OVERLAY_ID;
  closeOverlayById(overlayId);

  return new Promise((resolve) => {
    const dialog = createOverlayDialog({
      overlayId,
      cardClassName: options.cardClassName ?? "pi-welcome-card pi-overlay-card pi-overlay-card--s",
      restoreFocusOnClose: options.restoreFocusOnClose,
    });

    let settled = false;

    const settle = (value: string | null): void => {
      if (settled) {
        return;
      }

      settled = true;
      resolve(value);
    };

    const cancel = (): void => {
      settle(null);
      dialog.close();
    };

    const { header } = createOverlayHeader({
      onClose: cancel,
      closeLabel: options.cancelLabel ?? "Cancel",
      title: options.title,
    });

    const body = document.createElement("div");
    body.className = "pi-overlay-body";

    if (options.message) {
      const message = document.createElement("p");
      message.className = "pi-overlay-subtitle pi-confirm-dialog__message";
      message.textContent = options.message;
      body.appendChild(message);
    }

    const input = createOverlayInput({
      placeholder: options.placeholder ?? "",
    });
    input.value = options.initialValue ?? "";
    input.setAttribute("autocomplete", "off");
    body.appendChild(input);

    const submit = (): void => {
      settle(input.value);
      dialog.close();
    };

    const onInputKeydown = (event: KeyboardEvent): void => {
      if (event.key !== "Enter") {
        return;
      }

      event.preventDefault();
      submit();
    };

    input.addEventListener("keydown", onInputKeydown);

    const actions = document.createElement("div");
    actions.className = "pi-overlay-actions";

    const cancelButton = createOverlayButton({
      text: options.cancelLabel ?? "Cancel",
    });

    const confirmButton = createOverlayButton({
      text: options.confirmLabel ?? "Save",
      className: "pi-overlay-btn--primary",
    });

    cancelButton.addEventListener("click", cancel);
    confirmButton.addEventListener("click", submit);

    dialog.addCleanup(() => {
      input.removeEventListener("keydown", onInputKeydown);
      cancelButton.removeEventListener("click", cancel);
      confirmButton.removeEventListener("click", submit);
      settle(null);
    });

    actions.append(cancelButton, confirmButton);
    dialog.card.append(header, body, actions);
    dialog.mount();

    const focusInput = (): void => {
      if (typeof input.focus === "function") {
        input.focus();
      }

      if (typeof input.select === "function") {
        input.select();
      }
    };

    if (typeof requestAnimationFrame === "function") {
      requestAnimationFrame(focusInput);
    } else {
      focusInput();
    }
  });
}
