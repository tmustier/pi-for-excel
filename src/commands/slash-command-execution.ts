import { isBusyAllowedCommand } from "./busy-command-policy.js";
import { commandRegistry } from "./types.js";

export type SlashCommandExecutionResult = "not-found" | "busy-blocked" | "missing-queue" | "queued" | "executed";

export interface ExecuteSlashCommandOptions {
  name: string;
  args: string;
  busy: boolean;
  enqueueCommand?: (name: string, args: string) => void;
  beforeExecute?: () => void;
}

export function executeSlashCommand(options: ExecuteSlashCommandOptions): SlashCommandExecutionResult {
  const command = commandRegistry.get(options.name);
  if (!command) {
    return "not-found";
  }

  if (options.busy && !isBusyAllowedCommand(command)) {
    return "busy-blocked";
  }

  options.beforeExecute?.();

  if (options.name === "compact") {
    if (!options.enqueueCommand) {
      return "missing-queue";
    }

    options.enqueueCommand(options.name, options.args);
    return "queued";
  }

  void command.execute(options.args);
  return "executed";
}
