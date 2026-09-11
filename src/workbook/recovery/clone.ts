/**
 * Clone helpers for recovery state snapshots.
 *
 * A snapshot must not alias the object a caller handed in, and it must carry
 * only the properties its schema declares (host objects and older persisted
 * payloads can bring extras). One schema-driven clone does both.
 */

import { type Static, type TSchema } from "typebox";
import { Value } from "typebox/value";

import {
  RecoveryChartStateSchema,
  RecoveryCommentThreadStateSchema,
  RecoveryConditionalFormatRuleSchema,
  RecoveryFormatRangeStateSchema,
  RecoveryFormatSelectionSchema,
  RecoveryModifyStructureStateSchema,
} from "./schemas.js";
import type {
  RecoveryChartState,
  RecoveryCommentThreadState,
  RecoveryConditionalFormatRule,
  RecoveryFormatRangeState,
  RecoveryFormatSelection,
  RecoveryModifyStructureState,
} from "./types.js";

export function cloneRecoveryState<S extends TSchema>(schema: S, value: Static<S>): Static<S> {
  // Clean only removes properties the schema does not declare, so a value that
  // is valid going in is valid coming out.
  return Value.Clean(schema, Value.Clone(value)) as Static<S>;
}

export function cloneRecoveryChartState(state: RecoveryChartState): RecoveryChartState {
  return cloneRecoveryState(RecoveryChartStateSchema, state);
}

export function cloneRecoveryModifyStructureState(state: RecoveryModifyStructureState): RecoveryModifyStructureState {
  return cloneRecoveryState(RecoveryModifyStructureStateSchema, state);
}

export function cloneRecoveryConditionalFormatRules(
  rules: readonly RecoveryConditionalFormatRule[],
): RecoveryConditionalFormatRule[] {
  return rules.map((rule) => cloneRecoveryState(RecoveryConditionalFormatRuleSchema, rule));
}

export function cloneRecoveryCommentThreadState(state: RecoveryCommentThreadState): RecoveryCommentThreadState {
  return cloneRecoveryState(RecoveryCommentThreadStateSchema, state);
}

export function cloneRecoveryFormatSelection(selection: RecoveryFormatSelection): RecoveryFormatSelection {
  return cloneRecoveryState(RecoveryFormatSelectionSchema, selection);
}

export function cloneStringGrid(grid: readonly string[][]): string[][] {
  return grid.map((row) => [...row]);
}

export function cloneRecoveryFormatRangeState(state: RecoveryFormatRangeState): RecoveryFormatRangeState {
  return cloneRecoveryState(RecoveryFormatRangeStateSchema, state);
}
