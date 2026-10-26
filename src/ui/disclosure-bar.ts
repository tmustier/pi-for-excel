/**
 * First-run disclosure, built on the shared banner primitive.
 *
 * Tells the user once that Pi can reach external services, with a way to
 * review the related settings. Dismissed for good (per add-in) either way.
 */

import { t } from "../language/index.js";
import { createBanner, type BannerHandle } from "./banner.js";

const ACKNOWLEDGED_KEY = "pi.onboarding.disclosure.acknowledged";

function isAcknowledged(): boolean {
  try {
    return localStorage.getItem(ACKNOWLEDGED_KEY) === "1";
  } catch {
    return false;
  }
}

function setAcknowledged(): void {
  try {
    localStorage.setItem(ACKNOWLEDGED_KEY, "1");
  } catch {
    // ignore — private mode / storage unavailable
  }
}

export interface DisclosureBarOptions {
  /** Number of configured providers (shows only when ≥1). */
  providerCount: number;
  /** Opens the settings that control external access. */
  onOpenSettings: () => void;
}

/** Create the disclosure banner, or `null` if it was already acknowledged. */
export function createDisclosureBar(options: DisclosureBarOptions): BannerHandle | null {
  if (isAcknowledged() || options.providerCount < 1) {
    return null;
  }

  let handle: BannerHandle | null = null;
  const dismiss = (): void => {
    setAcknowledged();
    handle?.root.remove();
  };

  handle = createBanner({
    id: "disclosure",
    tone: "info",
    message: t("disclosure-bar.text"),
    actions: [
      { label: t("disclosure-bar.gotIt"), primary: true, onClick: dismiss },
      {
        label: t("disclosure-bar.review"),
        onClick: () => {
          dismiss();
          options.onOpenSettings();
        },
      },
    ],
    onDismiss: setAcknowledged,
  });

  return handle;
}
