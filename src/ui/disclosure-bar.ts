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
  }
}

export interface DisclosureBarOptions {
  providerCount: number;
  onOpenSettings: () => void;
}

export function createDisclosureBar(options: DisclosureBarOptions): BannerHandle | null {
  if (isAcknowledged() || options.providerCount < 1) {
    return null;
  }

  const dismiss = (): void => {
    setAcknowledged();
    handle.root.remove();
  };

  const handle = createBanner({
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
