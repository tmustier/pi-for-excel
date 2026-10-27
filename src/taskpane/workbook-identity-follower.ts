/**
 * Notices when the open workbook's identity changes while the taskpane runs.
 *
 * Identity is derived from the file path, so Save As (and the first save of a
 * new workbook) gives the same open document a new identity. Browser-stored
 * rules follow it without replacing existing destination rules. The caller
 * moves tab state along too; old identities keep their copies.
 */

import { getWorkbookRules, setWorkbookRules, workbookRulesKey } from "../rules/store.js";
import type { SettingsWriter } from "../storage/local/settings-store.js";

export interface WorkbookIdentityFollower {
  checkOnce(): Promise<void>;
}

export function createWorkbookIdentityFollower(options: {
  settings: SettingsWriter;
  resolveIdentity: () => Promise<string | null>;
  /** `from` is null when the workbook had no identity before (never saved). */
  onIdentityChanged: (change: { from: string | null; to: string }) => Promise<void>;
}): WorkbookIdentityFollower {
  let current: string | null | undefined;
  let checking: Promise<void> | null = null;

  const check = async (): Promise<void> => {
    const next = await options.resolveIdentity();
    const previous = current;
    // The first read only records where the pane started. A failed or empty
    // read later is not a move and keeps the identity the tabs belong to.
    if (previous === undefined) {
      current = next;
      return;
    }
    if (next === null || next === previous) return;
    const rules = await getWorkbookRules(options.settings, previous);
    // An empty entry is an explicit choice to clear rules, not a missing entry.
    const destinationExists = (await options.settings.get(workbookRulesKey(next))) != null;
    if (rules !== null && !destinationExists) {
      await setWorkbookRules(options.settings, next, rules);
    }
    await options.onIdentityChanged({ from: previous, to: next });
    current = next;
  };

  return {
    checkOnce(): Promise<void> {
      // Checks never overlap, so one change is handled once.
      checking ??= check().finally(() => {
        checking = null;
      });
      return checking;
    },
  };
}
