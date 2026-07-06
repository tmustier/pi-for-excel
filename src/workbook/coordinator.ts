/**
 * Workbook operation coordinator.
 *
 * Core invariant: parallel reasoning, serialized workbook mutation.
 */

import { getErrorMessage } from "../utils/errors.js";

export type WorkbookOperationType = "read" | "write";

export interface WorkbookOperationContext {
  workbookId: string;
  sessionId: string;
  opId: string;
  toolName?: string;
  expectedRevision?: number;
}

export type WorkbookCoordinatorEvent =
  | {
    type: "queued";
    operationType: WorkbookOperationType;
    context: WorkbookOperationContext;
    queuedWrites: number;
    revision: number;
  }
  | {
    type: "started";
    operationType: WorkbookOperationType;
    context: WorkbookOperationContext;
    queuedWrites: number;
    revision: number;
  }
  | {
    type: "completed";
    operationType: WorkbookOperationType;
    context: WorkbookOperationContext;
    queuedWrites: number;
    revision: number;
  }
  | {
    type: "failed";
    operationType: WorkbookOperationType;
    context: WorkbookOperationContext;
    queuedWrites: number;
    revision: number;
    errorMessage: string;
  };

export interface WorkbookQueueSnapshot {
  revision: number;
  queuedWrites: number;
  activeWrite: WorkbookOperationContext | null;
}

export interface WorkbookCoordinator {
  runRead<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<T>;
  runWrite<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<{ result: T; revision: number }>;
  getRevision(workbookId: string): number;
  getSnapshot(workbookId: string): WorkbookQueueSnapshot;
  subscribe(listener: (event: WorkbookCoordinatorEvent) => void): () => void;
}

interface QueuedWriteOperation {
  context: WorkbookOperationContext;
  execute: () => Promise<void>;
}

interface WorkbookQueueState {
  revision: number;
  running: boolean;
  activeWrite: WorkbookOperationContext | null;
  queue: QueuedWriteOperation[];
}

const UNKNOWN_WORKBOOK = "workbook:unknown";

function toError(error: DynamicValue): Error {
  if (error instanceof Error) return error;
  return new Error(getErrorMessage(error));
}

function normalizeWorkbookId(workbookId: string): string {
  const trimmed = workbookId.trim();
  return trimmed.length > 0 ? trimmed : UNKNOWN_WORKBOOK;
}

export function createWorkbookCoordinator(): WorkbookCoordinator {
  const states = new Map<string, WorkbookQueueState>();
  const listeners = new Set<(event: WorkbookCoordinatorEvent) => void>();

  function ensureState(workbookId: string): WorkbookQueueState {
    const key = normalizeWorkbookId(workbookId);
    const existing = states.get(key);
    if (existing) return existing;

    const created: WorkbookQueueState = {
      revision: 0,
      running: false,
      activeWrite: null,
      queue: [],
    };
    states.set(key, created);
    return created;
  }

  function emit(event: WorkbookCoordinatorEvent): void {
    for (const listener of listeners) {
      listener(event);
    }
  }

  async function runRead<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<T> {
    emit({
      type: "started",
      operationType: "read",
      context: ctx,
      queuedWrites: 0,
      revision: getRevision(ctx.workbookId),
    });

    try {
      const result = await fn();
      emit({
        type: "completed",
        operationType: "read",
        context: ctx,
        queuedWrites: 0,
        revision: getRevision(ctx.workbookId),
      });
      return result;
    } catch (error) {
      emit({
        type: "failed",
        operationType: "read",
        context: ctx,
        queuedWrites: 0,
        revision: getRevision(ctx.workbookId),
        errorMessage: getErrorMessage(error),
      });
      throw error;
    }
  }

  function processQueue(workbookId: string): void {
    const state = ensureState(workbookId);
    if (state.running) return;

    const next = state.queue.shift();
    if (!next) return;

    state.running = true;
    state.activeWrite = next.context;

    emit({
      type: "started",
      operationType: "write",
      context: next.context,
      queuedWrites: state.queue.length,
      revision: state.revision,
    });

    void (async () => {
      try {
        await next.execute();
      } finally {
        state.running = false;
        state.activeWrite = null;
        processQueue(workbookId);
      }
    })();
  }

  async function runWrite<T>(ctx: WorkbookOperationContext, fn: () => Promise<T>): Promise<{ result: T; revision: number }> {
    const state = ensureState(ctx.workbookId);

    return new Promise<{ result: T; revision: number }>((resolve, reject) => {
      const operation: QueuedWriteOperation = {
        context: ctx,
        execute: async () => {
          try {
            const result = await fn();
            state.revision += 1;
            emit({
              type: "completed",
              operationType: "write",
              context: ctx,
              queuedWrites: state.queue.length,
              revision: state.revision,
            });
            resolve({ result, revision: state.revision });
          } catch (error) {
            emit({
              type: "failed",
              operationType: "write",
              context: ctx,
              queuedWrites: state.queue.length,
              revision: state.revision,
              errorMessage: getErrorMessage(error),
            });
            reject(toError(error));
          }
        },
      };

      if (state.running || state.queue.length > 0) {
        emit({
          type: "queued",
          operationType: "write",
          context: ctx,
          queuedWrites: state.queue.length + 1,
          revision: state.revision,
        });
      }

      state.queue.push(operation);
      processQueue(ctx.workbookId);
    });
  }

  function getRevision(workbookId: string): number {
    const state = states.get(normalizeWorkbookId(workbookId));
    return state?.revision ?? 0;
  }

  function getSnapshot(workbookId: string): WorkbookQueueSnapshot {
    const state = ensureState(workbookId);
    return {
      revision: state.revision,
      queuedWrites: state.queue.length,
      activeWrite: state.activeWrite,
    };
  }

  function subscribe(listener: (event: WorkbookCoordinatorEvent) => void): () => void {
    listeners.add(listener);
    return () => listeners.delete(listener);
  }

  return {
    runRead,
    runWrite,
    getRevision,
    getSnapshot,
    subscribe,
  };
}
