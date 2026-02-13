/**
 * Boot — runs before any pi-web-ui components mount.
 *
 * 1. Imports Tailwind CSS (pi-web-ui/app.css)
 * 2. Installs compatibility patches (Lit class-field shadowing, markdown safety)
 * 3. Installs a thinking-label patch ("Thinking…" → "Thought for …")
 *
 * MUST be imported as the first module in taskpane.ts.
 */

import "@mariozechner/pi-web-ui/app.css";
import "./ui/theme.css";

import { installLitClassFieldShadowingPatch } from "./compat/lit-class-field-shadowing.js";
import { installMarkedSafetyPatch } from "./compat/marked-safety.js";
import { installThinkingDurationPatch } from "./compat/thinking-duration.js";
import { installDialogStyleHooks } from "./ui/dialog-style-hooks.js";
import { installThemeModeSync } from "./ui/theme-mode.js";

installLitClassFieldShadowingPatch();
installMarkedSafetyPatch();
installThinkingDurationPatch();
installDialogStyleHooks();
installThemeModeSync();
