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

function copyToClipboard(text: string, onCopied: () => void, fallbackElement: HTMLElement): void {
  if (!navigator.clipboard?.writeText) {
    selectElementText(fallbackElement);
    return;
  }

  void navigator.clipboard.writeText(text).then(onCopied, () => selectElementText(fallbackElement));
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
