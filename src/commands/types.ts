export type CommandSource = "builtin" | "extension" | "integration" | "prompt";

export interface SlashCommand {
  name: string;
  description: string;
  source: CommandSource;
  menuVisibility?: "hidden" | "exact";
  menuGroup?: "chat" | "workbook" | "setup";
  execute: (args: string) => void | Promise<void>;
  busyAllowed?: boolean;
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

  list(filter?: string): SlashCommand[] {
    const all = Array.from(this.commands.values());
    if (filter === undefined) return all;
    if (/\s/.test(filter)) return [];
    const q = filter.toLowerCase();
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
