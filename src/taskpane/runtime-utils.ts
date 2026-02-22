import type { AgentTool } from "@mariozechner/pi-agent-core";

import { isRecord } from "../utils/type-guards.js";

const FNV_OFFSET_BASIS = 0x811c9dc5;
const FNV_PRIME = 0x01000193;

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

/**
 * Build a stable fingerprint for runtime tool metadata.
 *
 * We intentionally exclude function identity (`execute`) so refresh passes that
 * rebuild tool objects without schema changes can be no-ops.
 */
export function createRuntimeToolFingerprint(tools: readonly AgentTool[]): string {
  if (tools.length === 0) {
    return "";
  }

  const parts: string[] = [];

  for (const tool of tools) {
    parts.push(
      `${tool.name}\u001f${tool.label}\u001f${tool.description}\u001f${serializeToolParameters(tool.parameters)}`,
    );
  }

  return hashString(parts.join("\u001e"));
}

/**
 * Decide whether a runtime refresh pass should call `agent.setTools(...)`.
 *
 * We update tools when either metadata changed (fingerprint delta) or when
 * extension-owned tool behavior changed without metadata deltas (revision delta).
 */
export function shouldApplyRuntimeToolUpdate(args: {
  previousFingerprint: string;
  nextFingerprint: string;
  previousExtensionToolRevision: number;
  nextExtensionToolRevision: number;
}): boolean {
  if (args.previousFingerprint !== args.nextFingerprint) {
    return true;
  }

  return args.previousExtensionToolRevision !== args.nextExtensionToolRevision;
}

export function isRuntimeAgentTool(value: unknown): value is AgentTool {
  if (!isRecord(value)) return false;

  return typeof value.name === "string"
    && typeof value.label === "string"
    && typeof value.description === "string"
    && "parameters" in value
    && typeof value.execute === "function";
}

export function normalizeRuntimeTools(candidates: readonly unknown[]): AgentTool[] {
  const seen = new Set<string>();
  const out: AgentTool[] = [];

  for (const candidate of candidates) {
    if (!isRuntimeAgentTool(candidate)) {
      console.warn("[pi] Ignoring invalid runtime tool payload", candidate);
      continue;
    }

    if (seen.has(candidate.name)) {
      console.warn(`[pi] Ignoring duplicate runtime tool name: ${candidate.name}`);
      continue;
    }

    seen.add(candidate.name);
    out.push(candidate);
  }

  return out;
}

export function isLikelyCorsErrorMessage(msg: string): boolean {
  const m = msg.toLowerCase();

  if (m.includes("failed to fetch")) return true;
  if (m.includes("load failed")) return true;
  if (m.includes("networkerror")) return true;

  if (m.includes("cors") || m.includes("cross-origin")) return true;
  if (m.includes("cors requests are not allowed")) return true;

  return false;
}

export function createAsyncCoalescer(task: () => Promise<void>): () => Promise<void> {
  let inFlight: Promise<void> | null = null;
  let rerunRequested = false;

  const run = async (): Promise<void> => {
    do {
      rerunRequested = false;
      await task();
    } while (rerunRequested);
  };

  return async (): Promise<void> => {
    if (inFlight) {
      rerunRequested = true;
      await inFlight;
      return;
    }

    inFlight = run();
    try {
      await inFlight;
    } finally {
      inFlight = null;
    }
  };
}

export async function awaitWithTimeout<T>(
  label: string,
  timeoutMs: number,
  task: Promise<T>,
): Promise<T> {
  let timeoutId: ReturnType<typeof setTimeout> | null = null;

  try {
    const timeoutPromise = new Promise<never>((_, reject) => {
      timeoutId = setTimeout(() => {
        reject(new Error(`${label} timed out after ${timeoutMs}ms`));
      }, timeoutMs);
    });

    return await Promise.race([task, timeoutPromise]);
  } finally {
    if (timeoutId) {
      clearTimeout(timeoutId);
    }
  }
}
