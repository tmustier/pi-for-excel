import { readFile, writeFile } from "node:fs/promises";
import { createHash } from "node:crypto";

// Deliberately pinned to the published 1.3.0 bytes, including declarations.
const root = new URL("../node_modules/@earendil-works/pi-durable/", import.meta.url);
const pkg = JSON.parse(await readFile(new URL("package.json", root), "utf8"));
if (pkg.version !== "1.3.0") throw new Error("Durable compaction patch requires exactly 1.3.0");
const callback = "readonly compactionForModel?: (model: Model<Api>) => Partial<CompactionPolicy>;";
const patches = [
  ["agent.js", "762e8a4743ec1da42ffcdcff7e7d4803e45e04c37f2d1047eee0507060a4b88c", [
    ["        compaction: { ...DEFAULT_COMPACTION_POLICY, ...settings?.compaction },", "        compaction: { ...DEFAULT_COMPACTION_POLICY, ...settings?.compaction },\n        compactionForModel: settings?.compactionForModel,"],
  ]],
  ["generation.js", "c7902ec218c6268474969142a70d4ba485eb75c7ee18f6d75c3923a60510c34e", [
    ["thresholdCompaction(view, entries, resolved.contextWindow, settings.compaction)", "thresholdCompaction(view, entries, resolved.contextWindow, { ...settings.compaction, ...settings.compactionForModel?.(resolved) })"],
    ["const blocking = contextWindow - policy.reserveTokens;", "const blocking = Math.min(contextWindow - policy.reserveTokens, policy.hardTriggerTokens ?? Number.POSITIVE_INFINITY);"],
    ["if (overflow && compacted === undefined && settings.compaction.enabled) {\n        const policy = settings.compaction;", "const overflowModel = runtime.models.getModel(ref.provider, ref.modelId);\n    const compactionPolicy = { ...settings.compaction, ...(overflowModel === undefined ? {} : settings.compactionForModel?.(overflowModel)) };\n    if (overflow && compacted === undefined && compactionPolicy.enabled) {\n        const policy = compactionPolicy;"],
  ]],
  ["compaction.js", "70b9218282ffe80fa20aaf98356531f3f27de8ab6ed0ba27b174324267aaa3a4", [
    ["const policy = settings.compaction;", "const policy = { ...settings.compaction, ...settings.compactionForModel?.(model) };"],
  ]],
  ["types.d.ts", "baf6e81f1c8e2b0b13d74cb4af02eb93d3d2bcd041920c99c8895f88bb486123", [
    ["import type { AssistantMessage, CacheRetention, Message, Models,", "import type { Api, Model, AssistantMessage, CacheRetention, Message, Models,"],
    ["    reserveTokens: number;", "    reserveTokens: number;\n    /** Optional blocking cap, independent of the summarization output budget. */\n    hardTriggerTokens?: number;"],
    ["    readonly compaction?: Partial<CompactionPolicy>;", `    readonly compaction?: Partial<CompactionPolicy>;\n    /** Pure synchronous per-model overrides, freshly merged at each built-in consumption site. */\n    ${callback}`],
    ["    readonly compaction: CompactionPolicy;", `    readonly compaction: CompactionPolicy;\n    ${callback}`],
  ]],
];
const hash = (value) => createHash("sha256").update(value).digest("hex");
const pending = [];
for (const [file, expected, changes] of patches) {
  const url = new URL(`dist/harness/${file}`, root);
  const source = await readFile(url, "utf8");
  let original = source;
  if (hash(original) !== expected) {
    for (const [before, after] of changes.toReversed()) original = original.replace(after, before);
    if (hash(original) !== expected) throw new Error(`Unrecognized Durable source: ${file}; refusing to patch`);
  }
  let patched = original;
  for (const [before, after] of changes) {
    if (patched.split(before).length !== 2) throw new Error(`Ambiguous Durable patch: ${file}`);
    patched = patched.replace(before, after);
  }
  if (source !== original && source !== patched) throw new Error(`Partial Durable patch: ${file}`);
  pending.push([url, patched]);
}
// Validate every file before writing any of them; repeat installs are harmless.
for (const [url, patched] of pending) await writeFile(url, patched);
