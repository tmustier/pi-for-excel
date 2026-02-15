/**
 * Slash-command policy for command execution while a runtime is busy.
 *
 * Keep this list centralized so keyboard-enter execution and command-menu
 * execution cannot drift.
 */

import { INTEGRATIONS_COMMAND_NAME, TOOLS_COMMAND_NAME } from "../integrations/naming.js";

const BUSY_ALLOWED_COMMANDS = new Set<string>([
  "compact",
  "new",
  "rules",
  "resume",
  "history",
  "reopen",
  "yolo",
  "addons",
  "extensions",
  "skills",
  TOOLS_COMMAND_NAME,
  INTEGRATIONS_COMMAND_NAME,
]);

export function isBusyAllowedCommand(commandName: string): boolean {
  return BUSY_ALLOWED_COMMANDS.has(commandName);
}
