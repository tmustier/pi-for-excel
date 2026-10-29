import { t } from "../language/index.js";
import { AlertTriangle, CircleX, Info, lucide } from "./lucide-icons.js";

export type BannerTone = "error" | "warning" | "info";

export interface BannerAction {
  label: string;
  onClick: () => void;
  primary?: boolean;
  expanded?: boolean;
}

export interface BannerOptions {
  id: string;
  tone: BannerTone;
  message: string;
  actions?: BannerAction[];
  details?: HTMLElement;
  onDismiss?: () => void;
}

export interface BannerHandle {
  root: HTMLElement;
  setActions: (actions: BannerAction[]) => void;
}

const TONE_ORDER: Record<BannerTone, number> = { error: 0, warning: 1, info: 2 };

function renderActions(row: HTMLElement, actions: readonly BannerAction[]): void {
  const focusedIndex = Array.from(row.children).findIndex((child) => child === document.activeElement);
  row.replaceChildren();
  row.hidden = actions.length === 0;
  for (const action of actions) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = action.primary ? "pi-banner__action pi-banner__action--primary" : "pi-banner__action";
    button.textContent = action.label;
    if (action.expanded !== undefined) {
      button.setAttribute("aria-expanded", String(action.expanded));
    }
    button.addEventListener("click", action.onClick);
    row.appendChild(button);
  }
  if (focusedIndex >= 0) row.querySelectorAll("button")[focusedIndex]?.focus({ preventScroll: true });
}

export function createBanner(options: BannerOptions): BannerHandle {
  const root = document.createElement("section");
  root.className = `pi-banner pi-banner--${options.tone}`;
  root.dataset.bannerId = options.id;
  root.dataset.tone = options.tone;
  root.setAttribute("role", options.tone === "error" ? "alert" : "status");

  const body = document.createElement("div");
  body.className = "pi-banner__body";

  const message = document.createElement("p");
  message.className = "pi-banner__message";
  message.textContent = options.message;

  const actionRow = document.createElement("div");
  actionRow.className = "pi-banner__actions";
  renderActions(actionRow, options.actions ?? []);

  body.append(message);
  if (options.details) body.append(options.details);
  body.append(actionRow);

  const icon = lucide(options.tone === "error" ? CircleX : options.tone === "warning" ? AlertTriangle : Info);
  icon.classList.add("pi-banner__icon");
  icon.setAttribute("aria-hidden", "true");
  root.append(icon, body);

  if (options.onDismiss) {
    const onDismiss = options.onDismiss;
    const dismiss = document.createElement("button");
    dismiss.type = "button";
    dismiss.className = "pi-banner__dismiss";
    dismiss.textContent = "×";
    dismiss.setAttribute("aria-label", t("banner.dismiss"));
    dismiss.title = t("banner.dismiss");
    dismiss.addEventListener("click", () => {
      root.remove();
      onDismiss();
    });
    root.appendChild(dismiss);
  }

  return {
    root,
    setActions: (actions) => renderActions(actionRow, actions),
  };
}

export function mountBanner(root: HTMLElement): boolean {
  const area = document.getElementById("pi-notices");
  if (!area) return false;

  const id = root.dataset.bannerId;
  const existing = id ? area.querySelector<HTMLElement>(`[data-banner-id="${CSS.escape(id)}"]`) : null;
  if (existing && existing !== root) {
    existing.replaceWith(root);
    return true;
  }

  const tone = root.dataset.tone;
  const rank = tone === "error" || tone === "warning" || tone === "info" ? TONE_ORDER[tone] : 3;
  const before = Array.from(area.children).find((child) => {
    if (!(child instanceof HTMLElement) || child === root) return false;
    const childTone = child.dataset.tone;
    const childRank = childTone === "error" || childTone === "warning" || childTone === "info" ? TONE_ORDER[childTone] : 3;
    return childRank > rank;
  });
  area.insertBefore(root, before ?? null);
  return true;
}

export function showBanner(options: BannerOptions): BannerHandle | null {
  const handle = createBanner(options);
  return mountBanner(handle.root) ? handle : null;
}

export function hideBanner(id: string): void {
  document.getElementById("pi-notices")?.querySelector(`[data-banner-id="${CSS.escape(id)}"]`)?.remove();
}
