import type { SlashCommand } from "./types.js";

export function isBusyAllowedCommand(command: Pick<SlashCommand, "name" | "source" | "busyAllowed">): boolean {
  return command.busyAllowed ?? command.source !== "prompt";
}
