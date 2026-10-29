function isToolsBridgeHealthPayloadShape(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}


import {
  fetchBridgeHealthJson,
  getBridgeSetting,
  resolveValidatedBridgeUrl,
} from "./bridge-service-utils.js";
import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
} from "./experimental-tool-gates/types.js";

export type LocalServiceStatus = "running" | "not_running" | "partial";

export interface PythonServiceEntry {
  name: "python";
  displayName: "Python (native)";
  status: LocalServiceStatus;
  pythonVersion?: string;
  libreofficeAvailable?: boolean;
  libreofficeVersion?: string;
  skillName: "python-bridge";
}

export interface TmuxServiceEntry {
  name: "tmux";
  displayName: "Terminal (tmux)";
  status: LocalServiceStatus;
  tmuxVersion?: string;
  tmuxSessions?: number;
  skillName: "tmux-bridge";
}

export type LocalServiceEntry = PythonServiceEntry | TmuxServiceEntry;

export interface BridgeHealthDependencies {
  getPythonBridgeUrl?: () => Promise<string | undefined>;
  getTmuxBridgeUrl?: () => Promise<string | undefined>;
  fetchHealth?: (url: string) => Promise<unknown>;
}

function parseLibreofficeVersion(rawVersion: string | undefined): string | undefined {
  if (!rawVersion) {
    return undefined;
  }

  const trimmed = rawVersion.trim();
  if (trimmed.length === 0) {
    return undefined;
  }

  const matched = trimmed.match(/\d+\.\d+(?:\.\d+){0,2}/);
  return matched?.[0];
}

function parsePythonHealth(payload: unknown): PythonServiceEntry {
  const base: PythonServiceEntry = {
    name: "python",
    displayName: "Python (native)",
    status: "not_running",
    skillName: "python-bridge",
  };

  if (!isToolsBridgeHealthPayloadShape(payload)) return base;
  if (payload.ok !== true) return base;

  const python = isToolsBridgeHealthPayloadShape(payload.python) ? payload.python : undefined;
  const pythonAvailable = python?.available === true;
  const pythonVersion = typeof python?.version === "string" ? python.version : undefined;

  const libreoffice = isToolsBridgeHealthPayloadShape(payload.libreoffice) ? payload.libreoffice : undefined;
  const libreofficeAvailable = libreoffice?.available === true;
  const libreofficeVersionRaw = typeof libreoffice?.version === "string" ? libreoffice.version : undefined;
  const libreofficeVersion = parseLibreofficeVersion(libreofficeVersionRaw);

  if (!pythonAvailable) {
    return base;
  }

  const status: LocalServiceStatus = libreofficeAvailable ? "running" : "partial";

  return {
    ...base,
    status,
    libreofficeAvailable,
    ...(pythonVersion !== undefined ? { pythonVersion } : {}),
    ...(libreofficeVersion !== undefined ? { libreofficeVersion } : {}),
  };
}

function parseTmuxHealth(payload: unknown): TmuxServiceEntry {
  const base: TmuxServiceEntry = {
    name: "tmux",
    displayName: "Terminal (tmux)",
    status: "not_running",
    skillName: "tmux-bridge",
  };

  if (!isToolsBridgeHealthPayloadShape(payload)) return base;
  if (payload.ok !== true) return base;

  const tmuxVersion = typeof payload.tmuxVersion === "string" ? payload.tmuxVersion : undefined;
  const tmuxSessions = typeof payload.sessions === "number" ? payload.sessions : undefined;

  if (payload.mode === "stub" || payload.backend === "stub") {
    return {
      ...base,
      status: "partial",
      ...(tmuxVersion !== undefined ? { tmuxVersion } : {}),
      ...(tmuxSessions !== undefined ? { tmuxSessions } : {}),
    };
  }

  return {
    ...base,
    status: "running",
    ...(tmuxVersion !== undefined ? { tmuxVersion } : {}),
    ...(tmuxSessions !== undefined ? { tmuxSessions } : {}),
  };
}

export async function probeLocalServices(
  deps: BridgeHealthDependencies = {},
): Promise<LocalServiceEntry[]> {
  const getPythonUrl = deps.getPythonBridgeUrl ?? (() => getBridgeSetting(PYTHON_BRIDGE_URL_SETTING_KEY));
  const getTmuxUrl = deps.getTmuxBridgeUrl ?? (() => getBridgeSetting(TMUX_BRIDGE_URL_SETTING_KEY));
  const fetchHealth = deps.fetchHealth ?? fetchBridgeHealthJson;

  const [pythonConfiguredUrl, tmuxConfiguredUrl] = await Promise.all([
    getPythonUrl(),
    getTmuxUrl(),
  ]);

  const pythonUrl = resolveValidatedBridgeUrl(pythonConfiguredUrl, DEFAULT_PYTHON_BRIDGE_URL).bridgeUrl;
  const tmuxUrl = resolveValidatedBridgeUrl(tmuxConfiguredUrl, DEFAULT_TMUX_BRIDGE_URL).bridgeUrl;

  const [pythonPayload, tmuxPayload] = await Promise.all([
    pythonUrl ? fetchHealth(pythonUrl) : Promise.resolve(null),
    tmuxUrl ? fetchHealth(tmuxUrl) : Promise.resolve(null),
  ]);

  return [
    parsePythonHealth(pythonPayload),
    parseTmuxHealth(tmuxPayload),
  ];
}
