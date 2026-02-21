/**
 * Slash command system — types and registry.
 */

export type CommandSource = "builtin" | "extension" | "integration" | "prompt";

export interface SlashCommand {
  /** Command name (without the leading `/`) */
  name: string;
  /** Short description shown in the menu */
  description: string;
  /** Source type for visual badge */
  source: CommandSource;
  /** Execute the command. `args` is everything after the command name. */
  execute: (args: string) => void | Promise<void>;
  /**
   * Optional busy policy override.
   *
   * When true, command is allowed while the active runtime is busy.
   * Extension commands default to busy-allowed unless explicitly set to false.
   */
  busyAllowed?: boolean;
  /** Optional: is this command available right now? */
  enabled?: () => boolean;
}

class CommandRegistry {
  private commands = new Map<string, SlashCommand>();

  register(cmd: SlashCommand): void {
    this.commands.set(cmd.name, cmd);
  }

  unregister(name: string): void {
    this.commands.delete(name);
  }

  get(name: string): SlashCommand | undefined {
    return this.commands.get(name);
  }

  /** Get all commands, optionally filtered by prefix */
  list(filter?: string): SlashCommand[] {
    const all = Array.from(this.commands.values());
    if (!filter) return all;
    const q = filter.toLowerCase();
    return all.filter(
      (c) => c.name.toLowerCase().includes(q) || c.description.toLowerCase().includes(q),
    );
  }
}

export const commandRegistry = new CommandRegistry();
