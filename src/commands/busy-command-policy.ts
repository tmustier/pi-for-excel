/**
 * Slash-command policy for command execution while a runtime is busy.
 *
 * Keep this list centralized so keyboard-enter execution and command-menu
 * execution cannot drift.
 */

import type { SlashCommand } from "./types.js";

export function isBusyAllowedCommand(command: Pick<SlashCommand, "name" | "source" | "busyAllowed">): boolean {
  // Prompt commands send into this conversation. Local commands do not;
  // /compact is routed to the per-runtime queue by slash-command-execution.
  return command.busyAllowed ?? command.source !== "prompt";
}
