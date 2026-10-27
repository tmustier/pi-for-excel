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
  // Backups stay available while Pi works: restores and full backups wait
  // for the workbook lock instead of interleaving with Pi's writes.
  "history",
  "revert",
  "backup",
  "reopen",
  // Read-only: they copy or download what is already in the transcript.
  "copy",
  "export",
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
