import type { WorkbookTabLayout } from "./tab-layout.js";

export interface TabLayoutPersistenceController {
  enable(): void;
  persist(layout: WorkbookTabLayout): void;
  flush(): Promise<void>;
}

export interface CreateTabLayoutPersistenceOptions {
  resolveWorkbookId: () => Promise<string | null>;
  saveLayout: (workbookId: string | null, layout: WorkbookTabLayout) => Promise<void>;
  warn?: (message: string, error: DynamicValue) => void;
}

function tabLayoutSignature(layout: WorkbookTabLayout): string {
  return JSON.stringify(layout);
}

export function createTabLayoutPersistence(
  options: CreateTabLayoutPersistenceOptions,
): TabLayoutPersistenceController {
  let enabled = false;
  let lastPersistedSignature: string | null = null;
  let persistChain: Promise<void> = Promise.resolve();

  const warn = options.warn ?? ((message: string, error: DynamicValue) => {
    console.warn(message, error);
  });

  return {
    enable(): void {
      enabled = true;
    },

    persist(layout: WorkbookTabLayout): void {
      if (!enabled) return;

      const layoutSignature = tabLayoutSignature(layout);

      persistChain = persistChain
        .then(
          async () => {
            const workbookId = await options.resolveWorkbookId();
            const persistSignature = `${workbookId ?? "__global__"}|${layoutSignature}`;
            if (persistSignature === lastPersistedSignature) {
              return;
            }

            await options.saveLayout(workbookId, layout);
            lastPersistedSignature = persistSignature;
          },
          () => undefined,
        )
        .catch((error: DynamicValue) => {
          warn("[pi] Failed to persist tab layout:", error);
        });
    },

    flush(): Promise<void> {
      return persistChain;
    },
  };
}
