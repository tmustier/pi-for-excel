import type { AgentTool } from "@mariozechner/pi-agent-core";

import { isRecord } from "../utils/type-guards.js";

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
