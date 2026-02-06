/**
 * Boot — runs before any pi-web-ui components mount.
 *
 * 1. Imports Tailwind CSS (pi-web-ui/app.css)
 * 2. Patches Lit's ReactiveElement to fix tsgo class field shadowing
 *
 * MUST be imported as the first module in taskpane.ts.
 */

import "@mariozechner/pi-web-ui/app.css";
import "./ui/theme.css";

import { installLitClassFieldShadowingPatch } from "./compat/lit-class-field-shadowing.js";

installLitClassFieldShadowingPatch();
