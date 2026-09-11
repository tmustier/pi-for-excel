import assert from "node:assert/strict";
import { test } from "node:test";

import {
  decodeToolDetails,
  decodeToolDetailsOfKind,
  getToolOutputTruncationDetails,
  isBridgeGateError,
  type ToolOutputTruncationDetails,
} from "../src/tools/tool-details.ts";

const truncation: ToolOutputTruncationDetails = {
  version: 1,
  strategy: "head",
  truncated: true,
  truncatedBy: "lines",
  totalLines: 40,
  totalBytes: 800,
  outputLines: 20,
  outputBytes: 400,
  maxLines: 20,
  maxBytes: 10_000,
};

const writeCells = {
  kind: "write_cells",
  blocked: false,
  address: "Sheet1!A1:B2",
  changes: {
    changedCount: 1,
    truncated: false,
    sample: [{ address: "Sheet1!A1", beforeValue: "", afterValue: "1" }],
  },
  recovery: { status: "checkpoint_created", snapshotId: "snap-1" },
};

void test("a well-formed payload decodes to its kind and keeps merged extras such as outputTruncation", () => {
  const raw: unknown = { ...writeCells, outputTruncation: truncation };

  const details = decodeToolDetails(raw);

  assert.equal(details?.kind, "write_cells");
  assert.equal(details, raw);
  assert.deepEqual(getToolOutputTruncationDetails(raw), truncation);
});

void test("payloads that are not a known, well-formed details object decode to undefined", () => {
  const cases: Record<string, unknown> = {
    null: null,
    array: [writeCells],
    "no kind": { blocked: false },
    "unknown kind": { ...writeCells, kind: "write_cellz" },
    "wrong scalar type": { ...writeCells, blocked: "no" },
    "enum outside its values": { ...writeCells, recovery: { status: "maybe" } },
    "malformed nested sample": {
      ...writeCells,
      changes: { changedCount: 1, truncated: false, sample: [{ address: 1 }] },
    },
  };

  for (const [name, raw] of Object.entries(cases)) {
    assert.equal(decodeToolDetails(raw), undefined, name);
  }
});

void test("decoding for an expected kind rejects a valid payload of another kind", () => {
  assert.equal(decodeToolDetailsOfKind(writeCells, "charts"), undefined);
  assert.equal(decodeToolDetailsOfKind(writeCells, "write_cells")?.address, "Sheet1!A1:B2");
});

void test("a dependency tree is validated to its leaves, not only at the root", () => {
  const leaf = { address: "Sheet1!C1", value: 3, precedents: [] };
  const tree = {
    kind: "trace_dependencies",
    root: { address: "Sheet1!A1", value: 1, formula: "=B1", precedents: [{ address: "Sheet1!B1", value: 2, precedents: [leaf] }] },
    mode: "precedents",
  };

  assert.equal(decodeToolDetails(tree)?.kind, "trace_dependencies");

  const brokenGrandchild = {
    ...tree,
    root: { ...tree.root, precedents: [{ address: "Sheet1!B1", value: 2, precedents: [{ ...leaf, address: 7 }] }] },
  };
  assert.equal(decodeToolDetails(brokenGrandchild), undefined);
});

void test("truncation metadata is readable whether or not the tool attached its own details", () => {
  assert.deepEqual(getToolOutputTruncationDetails({ outputTruncation: truncation }), truncation);
  assert.equal(getToolOutputTruncationDetails({ outputTruncation: { ...truncation, version: 2 } }), undefined);
  assert.equal(getToolOutputTruncationDetails(writeCells), undefined);
});

void test("a bridge gate error is a failed bridge result that carries a gate reason and a skill hint", () => {
  const gate = decodeToolDetails({
    kind: "tmux_bridge",
    ok: false,
    action: "list_sessions",
    gateReason: "bridge_unreachable",
    skillHint: "tmux-bridge",
  });
  const noHint = decodeToolDetails({ kind: "tmux_bridge", ok: false, action: "list_sessions", gateReason: "bridge_unreachable" });
  const blockedTransform = decodeToolDetails({
    kind: "python_transform_range",
    blocked: true,
    error: "blocked",
    gateReason: "bridge_unreachable",
    skillHint: "python-bridge",
  });
  const succeeded = decodeToolDetails({ kind: "python_bridge", ok: true, action: "run_python", gateReason: "bridge_unreachable", skillHint: "python-bridge" });

  assert.ok(gate && noHint && blockedTransform && succeeded);
  assert.equal(isBridgeGateError(gate), true);
  assert.equal(isBridgeGateError(noHint), false);
  assert.equal(isBridgeGateError(blockedTransform), false);
  assert.equal(isBridgeGateError(succeeded), false);
});
