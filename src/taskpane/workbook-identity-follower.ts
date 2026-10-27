/**
 * Notices when the open workbook's identity changes while the taskpane runs.
 *
 * Identity is derived from the file path, so Save As (and the first save of a
 * new workbook) gives the same open document a new identity. Per-workbook tab
 * state is stored under the identity, so the caller moves it along; otherwise
 * the saved file reopens with a fresh tab and the old file keeps the tabs.
 */

export interface WorkbookIdentityFollower {
  checkOnce(): Promise<void>;
}

export function createWorkbookIdentityFollower(options: {
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
    current = next;
    await options.onIdentityChanged({ from: previous, to: next });
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
