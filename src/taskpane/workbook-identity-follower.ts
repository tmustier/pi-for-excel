import { getWorkbookRules, setWorkbookRules, workbookRulesKey } from "../rules/store.js";
import type { SettingsWriter } from "../storage/local/settings-store.js";

export interface WorkbookIdentityFollower {
  checkOnce(): Promise<void>;
}

export function createWorkbookIdentityFollower(options: {
  settings: SettingsWriter;
  resolveIdentity: () => Promise<string | null>;
  onIdentityChanged: (change: { from: string | null; to: string }) => Promise<void>;
}): WorkbookIdentityFollower {
  let current: string | null | undefined;
  let checking: Promise<void> | null = null;

  const check = async (): Promise<void> => {
    const next = await options.resolveIdentity();
    const previous = current;
    if (previous === undefined) {
      current = next;
      return;
    }
    if (next === null || next === previous) return;
    const rules = await getWorkbookRules(options.settings, previous);
    const destinationExists = (await options.settings.get(workbookRulesKey(next))) != null;
    if (rules !== null && !destinationExists) {
      await setWorkbookRules(options.settings, next, rules);
    }
    await options.onIdentityChanged({ from: previous, to: next });
    current = next;
  };

  return {
    checkOnce(): Promise<void> {
      checking ??= check().finally(() => {
        checking = null;
      });
      return checking;
    },
  };
}
