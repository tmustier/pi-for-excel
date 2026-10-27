/**
 * Banners: a condition that stays true until it is fixed or acknowledged
 * (the last run failed, the proxy is down, the first-run disclosure).
 * See src/ui/README.md, "Notifications".
 *
 * Every banner lives in the notice area under the tab strip (`#pi-notices`),
 * most severe first, and can be dismissed. Banners never move focus.
 */

import { t } from "../language/index.js";
import { AlertTriangle, CircleX, Info, lucide } from "./lucide-icons.js";

export type BannerTone = "error" | "warning" | "info";

export interface BannerAction {
  label: string;
  onClick: () => void;
  /** Primary actions render as a filled button; others as text links. */
  primary?: boolean;
  /** Disclosure state, omitted for actions that do not toggle details. */
  expanded?: boolean;
}

export interface BannerOptions {
  /** One banner per id; showing an id again replaces it in place. */
  id: string;
  tone: BannerTone;
  message: string;
  actions?: BannerAction[];
  /** Optional extra content under the message (e.g. expandable steps). */
  details?: HTMLElement;
  /** Called after the user dismisses the banner with ×. Omit to hide ×. */
  onDismiss?: () => void;
}

export interface BannerHandle {
  root: HTMLElement;
  /** Re-render the action row (labels and attributes). */
  setActions: (actions: BannerAction[]) => void;
}

const TONE_ORDER: Record<BannerTone, number> = { error: 0, warning: 1, info: 2 };

function renderActions(row: HTMLElement, actions: readonly BannerAction[]): void {
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
}

/** Build a banner element without mounting it. */
export function createBanner(options: BannerOptions): BannerHandle {
  const root = document.createElement("section");
  root.className = `pi-banner pi-banner--${options.tone}`;
  root.dataset.bannerId = options.id;
  root.dataset.tone = options.tone;
  // Errors interrupt; everything else waits for the user.
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

/** Insert `root` into the notice area, keeping banners ordered by severity. */
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

/** Show (or replace) a banner in the notice area. */
export function showBanner(options: BannerOptions): BannerHandle | null {
  const handle = createBanner(options);
  return mountBanner(handle.root) ? handle : null;
}

export function hideBanner(id: string): void {
  document.getElementById("pi-notices")?.querySelector(`[data-banner-id="${CSS.escape(id)}"]`)?.remove();
}
