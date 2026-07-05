/**
 * Tool registry — creates all built-in tools for the agent.
 *
 * Canonical source of truth for core tools lives in `src/tools/registry.ts`.
 * Experimental/non-core tools are appended here.
 */

import { createCoreTools } from "./registry.js";
import type { SkillReadCache } from "../skills/read-cache.js";
import { createTmuxTool } from "./tmux.js";
import { createPythonRunTool } from "./python-run.js";
import { createLibreOfficeConvertTool } from "./libreoffice-convert.js";
import { createPythonTransformRangeTool } from "./python-transform-range.js";
import { createFilesTool } from "./files.js";
import { createExecuteOfficeJsTool } from "./execute-office-js.js";
import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
} from "./extensions-manager.js";

export interface CreateAllToolsOptions {
  /**
   * Host discriminator for future host-specific tool availability. Currently
   * accepted for taskpane host abstraction without changing the default tool set.
   */
  hostKind?: "office" | "wps" | "browser";
  getExtensionManager?: () => ExtensionsManagerToolRuntime | null;
  getSessionId?: () => string | null;
  skillReadCache?: SkillReadCache;
}

export function createAllTools(options: CreateAllToolsOptions = {}) {
  const getExtensionManager = options.getExtensionManager ?? (() => null);

  return [
    ...createCoreTools({
      skills: {
        getSessionId: options.getSessionId,
        readCache: options.skillReadCache,
      },
    }),
    createTmuxTool(),
    createPythonRunTool(),
    createLibreOfficeConvertTool(),
    createPythonTransformRangeTool(),
    createFilesTool(),
    createExecuteOfficeJsTool(),
    createExtensionsManagerTool({ getManager: getExtensionManager }),
  ];
}
