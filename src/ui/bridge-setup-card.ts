import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
} from "../tools/experimental-tool-gates.js";
import { probeBridgeHealth } from "../tools/bridge-service-utils.js";
import {
  isLibreOfficeBridgeDetails,
  isPythonBridgeDetails,
  isPythonTransformRangeDetails,
  isTmuxBridgeDetails,
  type BridgeGateReason,
  type LibreOfficeBridgeDetails,
  type PythonBridgeDetails,
  type PythonTransformRangeDetails,
  type TmuxBridgeDetails,
} from "../tools/tool-details.js";
import { t } from "../language/index.js";
import { AlertTriangle, Check, Copy, Terminal, lucide } from "./lucide-icons.js";

export const PYTHON_BRIDGE_SETUP_COMMAND = "npx pi-for-excel-python-bridge";
export const TMUX_BRIDGE_SETUP_COMMAND = "npx pi-for-excel-tmux-bridge";

export type BridgeSetupCardDetails =
  | TmuxBridgeDetails
  | PythonBridgeDetails
  | LibreOfficeBridgeDetails
  | PythonTransformRangeDetails;

interface BridgeSetupCardModel {
  title: string;
  command: string;
  probeUrl: string | null;
}

interface BridgeSetupCardDependencies {
  probeBridge?: (bridgeUrl: string) => Promise<boolean>;
}

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

function createCopyableCommand(command: string): HTMLDivElement {
  const row = document.createElement("div");
  row.className = "pi-bridge-setup__code";

  const code = document.createElement("code");
  code.textContent = command;

  const copyBtn = document.createElement("button");
  copyBtn.type = "button";
  copyBtn.className = "pi-bridge-setup__copy";
  copyBtn.title = t("bridge-setup.copyCommandTitle");
  copyBtn.setAttribute("aria-label", t("bridge-setup.copyCommandTitle"));
  copyBtn.replaceChildren(lucide(Copy));

  let resetTimeout: ReturnType<typeof setTimeout> | null = null;

  copyBtn.addEventListener("click", () => {
    copyToClipboard(command, () => {
      copyBtn.replaceChildren(lucide(Check));
      copyBtn.title = t("bridge-setup.copiedTitle");
      copyBtn.setAttribute("aria-label", t("bridge-setup.copiedTitle"));

      if (resetTimeout !== null) {
        clearTimeout(resetTimeout);
      }

      resetTimeout = setTimeout(() => {
        copyBtn.replaceChildren(lucide(Copy));
        copyBtn.title = t("bridge-setup.copyCommandTitle");
        copyBtn.setAttribute("aria-label", t("bridge-setup.copyCommandTitle"));
        resetTimeout = null;
      }, 1400);
    }, code);
  });

  row.append(code, copyBtn);
  return row;
}

function isSetupFailure(args: {
  error: string | undefined;
  gateReason: BridgeGateReason | undefined;
}): boolean {
  if (args.gateReason !== undefined) {
    return true;
  }

  const error = args.error;
  if (typeof error !== "string") {
    return false;
  }

  const normalized = error.toLowerCase();
  return normalized === "no_python_runtime" || normalized.includes("bridge");
}

function resolveProbeUrl(args: {
  bridgeUrl: string | undefined;
  gateReason: BridgeGateReason | undefined;
  defaultUrl: string;
}): string | null {
  if (args.bridgeUrl) {
    return args.bridgeUrl;
  }

  if (args.gateReason === "invalid_bridge_url") {
    return null;
  }

  return args.defaultUrl;
}

function toTmuxModel(details: TmuxBridgeDetails): BridgeSetupCardModel | null {
  if (details.ok !== false) {
    return null;
  }

  if (details.skillHint !== "tmux-bridge") {
    return null;
  }

  if (!isSetupFailure({
    error: details.error,
    gateReason: details.gateReason,
  })) {
    return null;
  }

  return {
    title: t("bridge-setup.tmuxTitle"),
    command: TMUX_BRIDGE_SETUP_COMMAND,
    probeUrl: resolveProbeUrl({
      bridgeUrl: details.bridgeUrl,
      gateReason: details.gateReason,
      defaultUrl: DEFAULT_TMUX_BRIDGE_URL,
    }),
  };
}

function toPythonModel(details: PythonBridgeDetails): BridgeSetupCardModel | null {
  if (details.ok !== false) {
    return null;
  }

  if (details.skillHint !== "python-bridge") {
    return null;
  }

  if (!isSetupFailure({
    error: details.error,
    gateReason: details.gateReason,
  })) {
    return null;
  }

  const title = details.error === "no_python_runtime"
    ? t("bridge-setup.pythonUnavailable")
    : "Python bridge is unavailable";

  return {
    title,
    command: PYTHON_BRIDGE_SETUP_COMMAND,
    probeUrl: resolveProbeUrl({
      bridgeUrl: details.bridgeUrl,
      gateReason: details.gateReason,
      defaultUrl: DEFAULT_PYTHON_BRIDGE_URL,
    }),
  };
}

function toLibreOfficeModel(details: LibreOfficeBridgeDetails): BridgeSetupCardModel | null {
  if (details.ok !== false) {
    return null;
  }

  if (details.skillHint !== "python-bridge") {
    return null;
  }

  if (!isSetupFailure({
    error: details.error,
    gateReason: details.gateReason,
  })) {
    return null;
  }

  return {
    title: t("bridge-setup.fileConversionUnavailable"),
    command: PYTHON_BRIDGE_SETUP_COMMAND,
    probeUrl: resolveProbeUrl({
      bridgeUrl: details.bridgeUrl,
      gateReason: details.gateReason,
      defaultUrl: DEFAULT_PYTHON_BRIDGE_URL,
    }),
  };
}

function toTransformRangeModel(details: PythonTransformRangeDetails): BridgeSetupCardModel | null {
  if (details.blocked !== false) {
    return null;
  }

  if (details.skillHint !== "python-bridge") {
    return null;
  }

  if (!isSetupFailure({
    error: details.error,
    gateReason: details.gateReason,
  })) {
    return null;
  }

  return {
    title: t("bridge-setup.pythonTransformUnavailable"),
    command: PYTHON_BRIDGE_SETUP_COMMAND,
    probeUrl: resolveProbeUrl({
      bridgeUrl: details.bridgeUrl,
      gateReason: details.gateReason,
      defaultUrl: DEFAULT_PYTHON_BRIDGE_URL,
    }),
  };
}

export function resolveBridgeSetupCardModel(details: unknown): BridgeSetupCardModel | null {
  if (isTmuxBridgeDetails(details)) {
    return toTmuxModel(details);
  }

  if (isPythonBridgeDetails(details)) {
    return toPythonModel(details);
  }

  if (isLibreOfficeBridgeDetails(details)) {
    return toLibreOfficeModel(details);
  }

  if (isPythonTransformRangeDetails(details)) {
    return toTransformRangeModel(details);
  }

  return null;
}

export function shouldShowBridgeSetupCard(details: unknown): details is BridgeSetupCardDetails {
  return resolveBridgeSetupCardModel(details) !== null;
}

export async function testBridgeSetupConnection(
  details: unknown,
  probeBridge: (bridgeUrl: string) => Promise<boolean> = probeBridgeHealth,
): Promise<boolean> {
  const model = resolveBridgeSetupCardModel(details);
  if (!model || !model.probeUrl) {
    return false;
  }

  return probeBridge(model.probeUrl);
}

export function mountBridgeSetupCard(
  container: HTMLElement,
  details: BridgeSetupCardDetails,
  dependencies: BridgeSetupCardDependencies = {},
): void {
  if (container.dataset.mounted === "true") {
    return;
  }

  const model = resolveBridgeSetupCardModel(details);
  if (!model) {
    return;
  }

  container.dataset.mounted = "true";

  const probeBridge = dependencies.probeBridge ?? probeBridgeHealth;

  const card = document.createElement("div");
  card.className = "pi-bridge-setup";

  const header = document.createElement("div");
  header.className = "pi-bridge-setup__header";

  const iconGlyph = model.command === TMUX_BRIDGE_SETUP_COMMAND ? Terminal : AlertTriangle;
  const icon = lucide(iconGlyph);
  icon.classList.add("pi-bridge-setup__icon");

  const titleEl = document.createElement("span");
  titleEl.className = "pi-bridge-setup__title";
  titleEl.textContent = model.title;

  header.append(icon, titleEl);

  const body = document.createElement("div");
  body.className = "pi-bridge-setup__body";

  const intro = document.createElement("p");
  intro.className = "pi-bridge-setup__text";
  intro.textContent = t("bridge-setup.intro");

  const hint = document.createElement("p");
  hint.className = "pi-bridge-setup__hint";
  hint.textContent = t("bridge-setup.keepRunning");

  const actions = document.createElement("div");
  actions.className = "pi-bridge-setup__actions";

  const testButton = document.createElement("button");
  testButton.type = "button";
  testButton.className = "pi-bridge-setup__test";
  testButton.textContent = t("bridge-setup.testConnection");

  const status = document.createElement("span");
  status.className = "pi-bridge-setup__status";
  status.setAttribute("role", "status");
  status.setAttribute("aria-live", "polite");

  let checking = false;

  if (!model.probeUrl) {
    testButton.disabled = true;
    status.textContent = t("bridge-setup.setValidUrlFirst");
    status.className = "pi-bridge-setup__status is-warn";
  }

  testButton.addEventListener("click", () => {
    const probeUrl = model.probeUrl;
    if (!probeUrl || checking) {
      return;
    }

    checking = true;
    testButton.disabled = true;
    testButton.textContent = t("bridge-setup.checking");
    status.textContent = t("bridge-setup.checkingBridge");
    status.className = "pi-bridge-setup__status";

    void probeBridge(probeUrl).then(
      (reachable) => {
        if (reachable) {
          status.textContent = t("bridge-setup.bridgeDetected");
          status.className = "pi-bridge-setup__status is-ok";
          return;
        }

        status.textContent = t("bridge-setup.bridgeNotDetected");
        status.className = "pi-bridge-setup__status is-warn";
      },
      () => {
        status.textContent = t("bridge-setup.cannotCheckBridge");
        status.className = "pi-bridge-setup__status is-error";
      },
    ).finally(() => {
      checking = false;
      testButton.disabled = false;
      testButton.textContent = t("bridge-setup.testConnection");
    });
  });

  actions.append(testButton, status);

  const footer = document.createElement("div");
  footer.className = "pi-bridge-setup__footer";

  const dismissButton = document.createElement("button");
  dismissButton.type = "button";
  dismissButton.className = "pi-bridge-setup__dismiss";
  dismissButton.textContent = t("bridge-setup.dismiss");
  dismissButton.addEventListener("click", () => {
    card.classList.add("is-dismissed");
    setTimeout(() => card.remove(), 200);
  });

  footer.append(dismissButton);

  body.append(intro, createCopyableCommand(model.command), hint, actions);
  card.append(header, body, footer);
  container.append(card);
}
