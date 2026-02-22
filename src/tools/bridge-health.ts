/**
 * Bridge health probing for system prompt Local Services section.
 *
 * Probes Python and tmux bridges once at session init, returning a compact
 * status snapshot that feeds into `buildSystemPrompt({ localServices })`.
 *
 * Reuses shared bridge URL/probe helpers used by the per-call gates in
 * `evaluation.ts`, but parses the full `/health` JSON payload for richer
 * status (python version, libreoffice availability, tmux sessions).
 */

import { isRecord } from "../utils/type-guards.js";

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

// ---------------------------------------------------------------------------
// Public types
// ---------------------------------------------------------------------------

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

// ---------------------------------------------------------------------------
// Dependency injection for testability
// ---------------------------------------------------------------------------

export interface BridgeHealthDependencies {
  getPythonBridgeUrl?: () => Promise<string | undefined>;
  getTmuxBridgeUrl?: () => Promise<string | undefined>;
  fetchHealth?: (url: string) => Promise<unknown>;
}

// ---------------------------------------------------------------------------
// Health payload parsing
// ---------------------------------------------------------------------------

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

  if (!isRecord(payload)) return base;
  if (payload.ok !== true) return base;

  // Extract python version
  const python = isRecord(payload.python) ? payload.python : undefined;
  const pythonAvailable = python?.available === true;
  const pythonVersion = typeof python?.version === "string" ? python.version : undefined;

  // Extract libreoffice availability/version
  const libreoffice = isRecord(payload.libreoffice) ? payload.libreoffice : undefined;
  const libreofficeAvailable = libreoffice?.available === true;
  const libreofficeVersionRaw = typeof libreoffice?.version === "string" ? libreoffice.version : undefined;
  const libreofficeVersion = parseLibreofficeVersion(libreofficeVersionRaw);

  if (!pythonAvailable) {
    // Bridge process is running but Python binary is missing — treat as not_running
    // since python_run calls will fail with 501.
    return base;
  }

  const status: LocalServiceStatus = libreofficeAvailable ? "running" : "partial";

  return {
    ...base,
    status,
    pythonVersion,
    libreofficeAvailable,
    libreofficeVersion,
  };
}

function parseTmuxHealth(payload: unknown): TmuxServiceEntry {
  const base: TmuxServiceEntry = {
    name: "tmux",
    displayName: "Terminal (tmux)",
    status: "not_running",
    skillName: "tmux-bridge",
  };

  if (!isRecord(payload)) return base;
  if (payload.ok !== true) return base;

  const tmuxVersion = typeof payload.tmuxVersion === "string" ? payload.tmuxVersion : undefined;
  const tmuxSessions = typeof payload.sessions === "number" ? payload.sessions : undefined;

  // Stub mode: bridge is running but tmux is not installed
  if (payload.mode === "stub" || payload.backend === "stub") {
    return { ...base, status: "partial", tmuxVersion, tmuxSessions };
  }

  return {
    ...base,
    status: "running",
    tmuxVersion,
    tmuxSessions,
  };
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Probe Python and tmux bridges in parallel. Returns a snapshot of local
 * service status suitable for `buildSystemPrompt({ localServices })`.
 *
 * Called once at session init. The result is stable for the session.
 */
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
