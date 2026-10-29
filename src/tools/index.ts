import type { SpreadsheetHostKind } from "../host/index.js";
import { createCoreTools, type AnyCoreTool } from "./registry.js";
import { selectOfficeCoupledToolForHost } from "./host-selection.js";
import type { SkillReadCache } from "../skills/read-cache.js";
import type { GuidanceChangeDependencies } from "./guidance-change.js";
import { createTmuxTool } from "./tmux.js";
import { createPythonRunTool } from "./python-run.js";
import { createLibreOfficeConvertTool } from "./libreoffice-convert.js";
import { createPythonTransformRangeTool } from "./python-transform-range.js";
import { createFilesTool } from "./files.js";
import { createExecuteOfficeJsTool } from "./execute-office-js.js";
import { createExecuteWpsJsTool } from "./execute-wps-js.js";
import type { RecoveryAttribution } from "./recovery-metadata.js";
import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
} from "./extensions-manager.js";

export interface CreateAllToolsOptions extends RecoveryAttribution {
  hostKind?: SpreadsheetHostKind;
  getExtensionManager?: () => ExtensionsManagerToolRuntime | null;
  getSessionId?: () => string | null;
  skillReadCache?: SkillReadCache;
  guidance?: GuidanceChangeDependencies;
}

export function createAllTools(options: CreateAllToolsOptions = {}): AnyCoreTool[] {
  const getExtensionManager = options.getExtensionManager ?? (() => null);
  const hostKind = options.hostKind ?? "office";

  const skills = {
    ...(options.getSessionId !== undefined ? { getSessionId: options.getSessionId } : {}),
    ...(options.skillReadCache !== undefined ? { readCache: options.skillReadCache } : {}),
  };

  return [
    ...createCoreTools({
      hostKind,
      skills,
      recovery: options,
      ...(options.guidance ? { guidance: options.guidance } : {}),
    }),
    createTmuxTool(),
    createPythonRunTool(),
    createLibreOfficeConvertTool(),
    selectOfficeCoupledToolForHost(createPythonTransformRangeTool({}, options), hostKind),
    createFilesTool(),
    selectOfficeCoupledToolForHost(createExecuteOfficeJsTool(), hostKind),
    ...(hostKind === "wps" ? [createExecuteWpsJsTool()] : []),
    createExtensionsManagerTool({ getManager: getExtensionManager }),
  ];
}
