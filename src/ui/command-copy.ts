import { t } from "../language/index.js";
import { Check, Copy, lucide } from "./lucide-icons.js";

const COPY_RESET_DELAY_MS = 1400;

function selectElementText(element: HTMLElement): void {
  const selection = window.getSelection();
  if (!selection) {
    return;
  }

  const range = document.createRange();
  range.selectNodeContents(element);
  selection.removeAllRanges();
  selection.addRange(range);
}

/**
 * Excel on the web frames the taskpane without clipboard-write permission, so
 * the async API is refused there; `execCommand("copy")` from the click still
 * works. Focus returns to the button afterwards.
 */
function copyViaHiddenTextarea(text: string): boolean {
  const previous = document.activeElement instanceof HTMLElement ? document.activeElement : null;
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
    previous?.focus({ preventScroll: true });
  }
}

function copyToClipboard(text: string, onCopied: () => void, fallbackElement: HTMLElement): void {
  const fallback = (): void => {
    if (copyViaHiddenTextarea(text)) onCopied();
    else selectElementText(fallbackElement);
  };

  if (!navigator.clipboard?.writeText) {
    fallback();
    return;
  }

  void navigator.clipboard.writeText(text).then(onCopied, fallback);
}

/** Create the shared command row used by inline setup cards. */
export function createCopyableCommand(command: string): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-command-copy";

  const code = document.createElement("code");
  code.textContent = command;

  const copyButton = document.createElement("button");
  copyButton.type = "button";
  copyButton.className = "pi-command-copy__button";

  const showCopyState = (): void => {
    const label = t("bridge-setup.copyCommandTitle");
    copyButton.replaceChildren(lucide(Copy));
    copyButton.title = label;
    copyButton.setAttribute("aria-label", label);
  };

  showCopyState();

  let resetTimeout: ReturnType<typeof setTimeout> | null = null;

  copyButton.addEventListener("click", () => {
    copyToClipboard(command, () => {
      const label = t("bridge-setup.copiedTitle");
      copyButton.replaceChildren(lucide(Check));
      copyButton.title = label;
      copyButton.setAttribute("aria-label", label);

      if (resetTimeout !== null) {
        clearTimeout(resetTimeout);
      }

      resetTimeout = setTimeout(() => {
        showCopyState();
        resetTimeout = null;
      }, COPY_RESET_DELAY_MS);
    }, code);
  });

  row.append(code, copyButton);
  return row;
}
