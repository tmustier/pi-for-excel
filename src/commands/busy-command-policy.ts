/**
 * Slash-command policy for command execution while a runtime is busy.
 *
 * Keep this list centralized so keyboard-enter execution and command-menu
 * execution cannot drift.
 */

import { TOOLS_COMMAND_NAME } from "../integrations/naming.js";
import type { SlashCommand } from "./types.js";

const BUSY_ALLOWED_COMMANDS = new Set<string>([
  "compact",
  "new",
  "rules",
  "resume",
  "history",
  "reopen",
  "yolo",
  "extensions",
  "plugins",
  "skills",
  "files",
  TOOLS_COMMAND_NAME,
]);

export function isBusyAllowedCommand(command: Pick<SlashCommand, "name" | "source" | "busyAllowed">): boolean {
  if (BUSY_ALLOWED_COMMANDS.has(command.name)) {
    return true;
  }

  if (command.source === "extension") {
    return command.busyAllowed ?? true;
  }

  return command.busyAllowed === true;
}
