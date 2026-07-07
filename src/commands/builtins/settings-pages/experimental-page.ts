/**
 * Experimental features page.
 */

import { t } from "../../../language/index.js";
import type { SettingsShellPage } from "../../../ui/settings-shell.js";
import {
  buildExperimentalFeatureContent,
  buildExperimentalFeatureFooter,
} from "../experimental-overlay.js";

export function createExperimentalPage(): SettingsShellPage {
  return {
    id: "experimental",
    parentId: "root",
    title: () => t("settings.section.experimental"),
    // No subtitle: the page content renders its own section headings + hints.
    render: (ctx) => {
      ctx.body.appendChild(buildExperimentalFeatureContent());
      ctx.body.appendChild(buildExperimentalFeatureFooter());
    },
  };
}
