import type { Api, Context, Model, Tool } from "@mariozechner/pi-ai";

export type PrefixChangeReason = "model" | "systemPrompt" | "tools";

export interface PrefixFingerprint {
  modelKey: string;
  systemPromptKey: string;
  toolsKey: string;
}

const FNV_OFFSET_BASIS = 0x811c9dc5;
const FNV_PRIME = 0x01000193;
const NO_SESSION_KEY = "__no-session__";
const DEFAULT_MAX_TRACKED_SESSIONS = 64;

function hashString(value: string): string {
  let hash = FNV_OFFSET_BASIS;

  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, FNV_PRIME) >>> 0;
  }

  return hash.toString(16).padStart(8, "0");
}

function serializeToolParameters(parameters: unknown): string {
  try {
    const serialized = JSON.stringify(parameters);
    return serialized ?? "null";
  } catch {
    return "[unserializable]";
  }
}

function serializeToolSchemas(tools: readonly Tool[] | undefined): string {
  if (!tools || tools.length === 0) {
    return "";
  }

  const parts: string[] = [];

  for (const tool of tools) {
    const params = serializeToolParameters(tool.parameters);
    parts.push(`${tool.name}\u001f${tool.description}\u001f${params}`);
  }

  return parts.join("\u001e");
}

function normalizeSessionKey(sessionId: string | undefined): string {
  if (typeof sessionId !== "string") {
    return NO_SESSION_KEY;
  }

  const normalized = sessionId.trim();
  return normalized.length > 0 ? normalized : NO_SESSION_KEY;
}

export function createPrefixFingerprint(model: Model<Api>, context: Context): PrefixFingerprint {
  const modelIdentity = `${model.provider}/${model.id}`;
  const systemPrompt = context.systemPrompt ?? "";
  const toolSchemas = serializeToolSchemas(context.tools);

  return {
    modelKey: hashString(modelIdentity),
    systemPromptKey: hashString(systemPrompt),
    toolsKey: hashString(toolSchemas),
  };
}

export function getPrefixChangeReasons(
  previous: PrefixFingerprint | undefined,
  current: PrefixFingerprint,
): PrefixChangeReason[] {
  if (!previous) {
    return [];
  }

  const reasons: PrefixChangeReason[] = [];

  if (previous.modelKey !== current.modelKey) {
    reasons.push("model");
  }

  if (previous.systemPromptKey !== current.systemPromptKey) {
    reasons.push("systemPrompt");
  }

  if (previous.toolsKey !== current.toolsKey) {
    reasons.push("tools");
  }

  return reasons;
}

export class PrefixChangeTracker {
  private readonly lastBySession = new Map<string, PrefixFingerprint>();
  private readonly maxTrackedSessions: number;

  constructor(opts?: { maxTrackedSessions?: number }) {
    const maxTrackedSessions = opts?.maxTrackedSessions;
    this.maxTrackedSessions =
      typeof maxTrackedSessions === "number" && Number.isFinite(maxTrackedSessions) && maxTrackedSessions > 0
        ? Math.floor(maxTrackedSessions)
        : DEFAULT_MAX_TRACKED_SESSIONS;
  }

  observe(sessionId: string | undefined, fingerprint: PrefixFingerprint): PrefixChangeReason[] {
    const sessionKey = normalizeSessionKey(sessionId);
    const previous = this.lastBySession.get(sessionKey);
    const reasons = getPrefixChangeReasons(previous, fingerprint);

    if (this.lastBySession.has(sessionKey)) {
      this.lastBySession.delete(sessionKey);
    }

    this.lastBySession.set(sessionKey, fingerprint);
    this.trimToBudget();

    return reasons;
  }

  private trimToBudget(): void {
    while (this.lastBySession.size > this.maxTrackedSessions) {
      const oldestSessionKey = this.lastBySession.keys().next().value;
      if (typeof oldestSessionKey !== "string") {
        break;
      }

      this.lastBySession.delete(oldestSessionKey);
    }
  }
}
