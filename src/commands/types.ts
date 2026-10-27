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
  /** Menu presentation only; hidden commands remain callable by their full name. */
  menuVisibility?: "hidden" | "exact";
  menuGroup?: "chat" | "workbook" | "setup";
  /** Execute the command. `args` is everything after the command name. */
  execute: (args: string) => void | Promise<void>;
  /**
   * Optional busy policy override.
   *
   * When true, command is allowed while the active runtime is busy.
   * Local and extension commands default to allowed. Prompt commands conflict
   * with the running conversation and default to refused.
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
    if (filter === undefined) return all;
    const q = filter.toLowerCase().split(/\s/)[0] ?? "";
    // A fully typed alias dispatches directly on Enter. A description match
    // must not substitute the parent command and discard the alias's target.
    if (this.commands.get(q)?.menuVisibility === "hidden") return [];
    const matches = all.filter(
      (c) => c.menuVisibility !== "hidden"
        && (c.menuVisibility !== "exact" || c.name.toLowerCase() === q)
        && (c.name.toLowerCase().includes(q) || c.description.toLowerCase().includes(q)),
    );
    const rank = (command: SlashCommand): number => command.name.toLowerCase() === q ? 0
      : command.name.toLowerCase().startsWith(q) ? 1
      : command.name.toLowerCase().includes(q) ? 2 : 3;
    const groups = { chat: 0, workbook: 1, setup: 2 };
    return matches.sort((left, right) => {
      return rank(left) - rank(right)
        || groups[left.menuGroup ?? "setup"] - groups[right.menuGroup ?? "setup"];
    });
  }
}

export const commandRegistry = new CommandRegistry();
