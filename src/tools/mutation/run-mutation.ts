/**
 * `excelRun` for workbook mutations.
 *
 * When `excelRun` gives up on a batch (its timeout, #792), the tool call ends
 * but `Excel.run` cannot be cancelled: the callback keeps running and would go
 * on to send its later writes after the tool already reported. Here, once the
 * call has failed, every later `context.sync()` of that run rejects without
 * reaching Excel. Only the batch already in flight may still land, which is
 * what the tools' "may or may not have been applied" wording allows for.
 */

import { excelRun, type ExcelRunOptions } from "../../excel/helpers.js";

export class AbandonedBatchError extends Error {
  constructor() {
    super("This Excel batch was abandoned after the tool call failed; nothing more was sent to Excel.");
    this.name = "AbandonedBatchError";
  }
}

export async function runMutation<T>(
  fn: (context: Excel.RequestContext) => Promise<T>,
  options: ExcelRunOptions = {},
): Promise<T> {
  let abandoned = false;
  try {
    return await excelRun((context) => {
      const sync = context.sync.bind(context);
      context.sync = <V>(passThroughValue?: V): Promise<V> =>
        abandoned ? Promise.reject(new AbandonedBatchError()) : sync(passThroughValue);
      return fn(context);
    }, options);
  } catch (error) {
    abandoned = true;
    throw error;
  }
}
