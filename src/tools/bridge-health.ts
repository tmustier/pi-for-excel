/**
 * Bridge health probing for system prompt Local Services section.
 *
 * Probes Python and tmux bridges once at session init, returning a compact
 * status snapshot that feeds into `buildSystemPrompt({ localServices })`.
 *
 * Reuses the same URL resolution + validation logic as the per-call gates
 * in `evaluation.ts`, but parses the full `/health` JSON payload for
 * richer status (python version, libreoffice availability, tmux sessions).
 */

import { validateOfficeProxyUrl } from "../auth/proxy-validation.js";
import { isRecord } from "../utils/type-guards.js";

import {
  DEFAULT_PYTHON_BRIDGE_URL,
  DEFAULT_TMUX_BRIDGE_URL,
  PYTHON_BRIDGE_URL_SETTING_KEY,
  TMUX_BRIDGE_URL_SETTING_KEY,
} from "./experimental-tool-gates/types.js";

const BRIDGE_HEALTH_PATH = "/health";
const BRIDGE_HEALTH_TIMEOUT_MS = 900;

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
// URL resolution (same logic as evaluation.ts)
// ---------------------------------------------------------------------------

async function defaultGetBridgeUrl(settingKey: string): Promise<string | undefined> {
  try {
    const storageModule = await import("@mariozechner/pi-web-ui/dist/storage/app-storage.js");
    const storage = storageModule.getAppStorage();
    const value = await storage.settings.get<string>(settingKey);
    if (typeof value !== "string") return undefined;

    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : undefined;
  } catch {
    return undefined;
  }
}

function resolveAndValidateUrl(configuredUrl: string | undefined, defaultUrl: string): string | null {
  const rawUrl = configuredUrl ?? defaultUrl;
  try {
    return validateOfficeProxyUrl(rawUrl);
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------------------
// Health fetch with timeout
// ---------------------------------------------------------------------------

async function defaultFetchHealth(url: string): Promise<unknown> {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), BRIDGE_HEALTH_TIMEOUT_MS);

  try {
    const target = `${url.replace(/\/+$/, "")}${BRIDGE_HEALTH_PATH}`;
    const response = await fetch(target, {
      method: "GET",
      signal: controller.signal,
    });
    if (!response.ok) return null;
    return await response.json() as unknown;
  } catch {
    return null;
  } finally {
    clearTimeout(timeoutId);
  }
}

// ---------------------------------------------------------------------------
// Health payload parsing
// ---------------------------------------------------------------------------

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

  // Extract libreoffice availability
  const libreoffice = isRecord(payload.libreoffice) ? payload.libreoffice : undefined;
  const libreofficeAvailable = libreoffice?.available === true;

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
  const getPythonUrl = deps.getPythonBridgeUrl ?? (() => defaultGetBridgeUrl(PYTHON_BRIDGE_URL_SETTING_KEY));
  const getTmuxUrl = deps.getTmuxBridgeUrl ?? (() => defaultGetBridgeUrl(TMUX_BRIDGE_URL_SETTING_KEY));
  const fetchHealth = deps.fetchHealth ?? defaultFetchHealth;

  const [pythonConfiguredUrl, tmuxConfiguredUrl] = await Promise.all([
    getPythonUrl(),
    getTmuxUrl(),
  ]);

  const pythonUrl = resolveAndValidateUrl(pythonConfiguredUrl, DEFAULT_PYTHON_BRIDGE_URL);
  const tmuxUrl = resolveAndValidateUrl(tmuxConfiguredUrl, DEFAULT_TMUX_BRIDGE_URL);

  const [pythonPayload, tmuxPayload] = await Promise.all([
    pythonUrl ? fetchHealth(pythonUrl) : Promise.resolve(null),
    tmuxUrl ? fetchHealth(tmuxUrl) : Promise.resolve(null),
  ]);

  return [
    parsePythonHealth(pythonPayload),
    parseTmuxHealth(tmuxPayload),
  ];
}
