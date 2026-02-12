/**
 * Tool registry — creates all built-in tools for the agent.
 *
 * Canonical source of truth for core tools lives in `src/tools/registry.ts`.
 * Experimental/non-core tools are appended here.
 */

import { createCoreTools } from "./registry.js";
import { createTmuxTool } from "./tmux.js";
import { createPythonRunTool } from "./python-run.js";
import { createLibreOfficeConvertTool } from "./libreoffice-convert.js";
import { createPythonTransformRangeTool } from "./python-transform-range.js";
import { createFilesTool } from "./files.js";
import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
} from "./extensions-manager.js";

export interface CreateAllToolsOptions {
  getExtensionManager?: () => ExtensionsManagerToolRuntime | null;
}

export function createAllTools(options: CreateAllToolsOptions = {}) {
  const getExtensionManager = options.getExtensionManager ?? (() => null);

  return [
    ...createCoreTools(),
    createTmuxTool(),
    createPythonRunTool(),
    createLibreOfficeConvertTool(),
    createPythonTransformRangeTool(),
    createFilesTool(),
    createExtensionsManagerTool({ getManager: getExtensionManager }),
  ];
}
