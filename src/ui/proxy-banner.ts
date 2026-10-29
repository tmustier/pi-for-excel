import { t } from "../language/index.js";
import { PROXY_HELPER_COMMAND } from "../auth/proxy-validation.js";
import { createBanner, hideBanner, mountBanner } from "./banner.js";
import { createCopyableCommand } from "./command-copy.js";

const INSTALL_GUIDE_URL = "https://pi.dev/excel#connect";
const BANNER_ID = "proxy";

export type ProxyBannerState = "detected" | "not-detected" | "unknown";

export interface ProxyBannerController {
  update: (state: ProxyBannerState, blockedByBrowser?: boolean) => Promise<void>;
}

export function createProxyBannerController(opts: {
  isProxyEnabled: () => Promise<boolean>;
}): ProxyBannerController {
  let dismissed = false;

  return {
    update: async (state, blockedByBrowser = false) => {
      if (state === "detected") dismissed = false;
      if (state !== "not-detected" || dismissed) {
        hideBanner(BANNER_ID);
        return;
      }

      const enabled = await opts.isProxyEnabled().catch(() => false);
      if (!enabled) {
        hideBanner(BANNER_ID);
        return;
      }

      const details = document.createElement("div");
      details.className = "pi-banner__details";
      details.hidden = true;

      const intro = document.createElement("p");
      intro.textContent = blockedByBrowser ? t("proxy-banner.blocked_intro") : t("proxy-banner.intro");
      if (blockedByBrowser) {
        details.append(intro);
      } else {
        const hint = document.createElement("p");
        hint.className = "pi-banner__hint";
        hint.textContent = t("proxy-banner.hint");

        const guide = document.createElement("a");
        guide.className = "pi-banner__link";
        guide.href = INSTALL_GUIDE_URL;
        guide.target = "_blank";
        guide.rel = "noopener noreferrer";
        guide.textContent = t("proxy-banner.guideLink");

        details.append(intro, createCopyableCommand(PROXY_HELPER_COMMAND), hint, guide);
      }
      const handle = createBanner({
        id: BANNER_ID,
        tone: "warning",
        message: blockedByBrowser ? t("proxy-banner.blocked") : t("proxy-banner.warning"),
        details,
        onDismiss: () => {
          dismissed = true;
        },
      });

      const renderToggle = (): void => {
        const open = !details.hidden;
        handle.setActions([{
          label: open ? t("proxy-banner.hideSteps") : t("proxy-banner.action"),
          expanded: open,
          onClick: () => {
            details.hidden = open;
            renderToggle();
          },
        }]);
      };
      renderToggle();
      mountBanner(handle.root);
    },
  };
}
