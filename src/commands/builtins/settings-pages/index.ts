/**
 * Settings pages — registry and public API.
 *
 * All configuration surfaces live as pages inside one settings shell
 * (see `src/ui/settings-shell.ts` and `docs/settings-ux-redesign.md`):
 *
 *   root ─┬─ providers        model provider logins
 *         ├─ gateway          custom OpenAI-compatible gateway
 *         ├─ proxy            local proxy helper
 *         ├─ rules            rules + workbook rules + format conventions
 *         ├─ backups          workbook recovery snapshots
 *         ├─ connections      web search / extension connections / MCP
 *         ├─ plugins          installed extensions
 *         ├─ skills           bundled + external skills
 *         ├─ shortcuts        keyboard shortcuts reference
 *         └─ experimental     experimental feature toggles
 */

import { createSettingsShell, type SettingsShellPage } from "../../../ui/settings-shell.js";
import { SETTINGS_OVERLAY_ID } from "../../../ui/overlay-ids.js";
import { t } from "../../../language/index.js";
import {
  getSettingsPagesDependencies,
  configureSettingsPages,
  type SettingsPagesDependencies,
} from "./dependencies.js";
import { createRootPage } from "./root-page.js";
import { createProvidersPage } from "./providers-page.js";
import { createGatewayPage } from "./gateway-page.js";
import { createProxyPage } from "./proxy-page.js";
import { createRulesPage } from "./rules-page.js";
import { createBackupsPage } from "./backups-page.js";
import {
  createConnectionsPage,
  createPluginsPage,
  createSkillsPage,
} from "./extensions-pages.js";
import { createShortcutsPage } from "./shortcuts-page.js";
import { createExperimentalPage } from "./experimental-page.js";

export { configureSettingsPages, type SettingsPagesDependencies };

/** Legacy alias for the extensions pages (used by slash commands). */
export type ExtensionsHubTab = "connections" | "plugins" | "skills";

export type SettingsPageId =
  | "root"
  | "providers"
  | "gateway"
  | "proxy"
  | "rules"
  | "backups"
  | "connections"
  | "plugins"
  | "skills"
  | "shortcuts"
  | "experimental";

let pageRegistry: Map<string, SettingsShellPage> | null = null;

function getPageRegistry(): Map<string, SettingsShellPage> {
  if (pageRegistry) return pageRegistry;

  const pages: SettingsShellPage[] = [
    createRootPage(),
    createProvidersPage(),
    createGatewayPage(),
    createProxyPage(),
    createRulesPage(),
    createBackupsPage(),
    createConnectionsPage(),
    createPluginsPage(),
    createSkillsPage(),
    createShortcutsPage(),
    createExperimentalPage(),
  ];

  pageRegistry = new Map(pages.map((page) => [page.id, page]));
  return pageRegistry;
}

const shell = createSettingsShell({
  overlayId: SETTINGS_OVERLAY_ID,
  rootId: "root",
  getPage: (pageId) => getPageRegistry().get(pageId),
  backLabel: () => t("settings.shell.back"),
  closeLabel: () => t("settings.close"),
});

/**
 * Open the settings overlay at a page. Calling with no page while the
 * overlay is open toggles it closed (slash-command toggle behavior).
 */
export async function openSettings(pageId?: SettingsPageId): Promise<void> {
  if (shell.isOpen() && pageId === undefined) {
    await shell.requestClose();
    return;
  }

  await shell.open(pageId);
}

export { getSettingsPagesDependencies };
