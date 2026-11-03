import { isAbortError } from "../utils/network.js";

export class ToolStoppedError extends Error {
  constructor(message = "Stopped by the user.") {
    super(message);
    this.name = "ToolStoppedError";
  }
}

export function isStop(error: unknown, signal: AbortSignal | undefined): boolean {
  return error instanceof ToolStoppedError || (signal?.aborted === true && isAbortError(error));
}
