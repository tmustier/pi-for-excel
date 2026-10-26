import assert from "node:assert/strict";
import { test } from "node:test";

import { Agent, type AgentMessage, type AgentTool } from "@earendil-works/pi-agent-core";
import {
  createModels,
  fauxAssistantMessage,
  fauxProvider,
  fauxToolCall,
  type TranscriptContext,
} from "@earendil-works/pi-ai";
import { Type } from "typebox";

import { ChangeTracker } from "../src/context/change-tracker.ts";
import { readSelectionContext } from "../src/context/selection.ts";
import { createConvertToLlm } from "../src/messages/convert-to-llm.ts";
import { createAutoContext } from "../src/taskpane/context-injection.ts";

// ---------------------------------------------------------------------------
// Host boundary: a workbook that emits Office.js worksheet events.
// ---------------------------------------------------------------------------

type Handler = (event: never) => Promise<void>;

class FakeWorkbook {
  readonly sheets = new Map<string, string>([["sheet-1", "Data"]]);
  private readonly handlers = new Map<string, Handler>();

  install(): () => void {
    const handlerSlot = (name: string) => ({ add: (handler: Handler) => this.handlers.set(name, handler) });
    const worksheets = {
      items: [] as Array<{ id: string; name: string }>,
      load: () => {
        worksheets.items = [...this.sheets].map(([id, name]) => ({ id, name }));
      },
      getItem: (id: string) => ({ name: this.sheets.get(id) ?? id, load: () => {} }),
      onChanged: handlerSlot("changed"),
      onAdded: handlerSlot("added"),
      onDeleted: handlerSlot("deleted"),
      onNameChanged: handlerSlot("nameChanged"),
      onMoved: handlerSlot("moved"),
      onVisibilityChanged: handlerSlot("visibility"),
    };
    const context = { workbook: { worksheets }, sync: () => Promise.resolve() };
    const previous: unknown = Reflect.get(globalThis, "Excel");
    Reflect.set(globalThis, "Excel", {
      run: <T>(callback: (ctx: unknown) => Promise<T>): Promise<T> => callback(context),
    });
    return () => Reflect.set(globalThis, "Excel", previous);
  }

  private async emit(name: string, event: object): Promise<void> {
    const handler = this.handlers.get(name);
    if (!handler) throw new Error(`no ${name} handler registered`);
    await (handler as (event: object) => Promise<void>)(event);
  }

  edit(sheetId: string, address: string, triggerSource: "Unknown" | "ThisLocalAddin"): Promise<void> {
    return this.emit("changed", {
      worksheetId: sheetId,
      address,
      changeType: "RangeEdited",
      source: "Local",
      triggerSource,
    });
  }

  async addSheet(id: string, name: string): Promise<void> {
    this.sheets.set(id, name);
    await this.emit("added", { worksheetId: id, source: "Local" });
  }

  blueprint(): string {
    return [...this.sheets.values()].map((name, i) => `${i + 1}. **${name}**`).join("\n");
  }
}

function textOf(message: TranscriptContext["messages"][number]): string {
  if (message.role !== "user" || typeof message.content === "string") return "";
  return message.content.map((block) => (block.type === "text" ? block.text : "")).join("");
}

function autoContextTexts(request: TranscriptContext): string[] {
  return request.messages.map(textOf).filter((text) => text.startsWith("[Auto-context]"));
}

// ---------------------------------------------------------------------------
// Auto-context through a real agent loop
// ---------------------------------------------------------------------------

void test("auto-context stores workbook and edit notes once, keeps them for later calls, and re-reads only the selection", async () => {
  const workbook = new FakeWorkbook();
  const uninstall = workbook.install();
  try {
    let blueprint = workbook.blueprint();
    const tracker = new ChangeTracker({ onStructureChange: () => { blueprint = workbook.blueprint(); } });
    await tracker.start();

    const faux = fauxProvider({ models: [{ id: "auto-context", contextWindow: 200_000, maxTokens: 4_096 }] });
    const models = createModels();
    models.setProvider(faux.provider);
    faux.setResponses([
      fauxAssistantMessage(fauxToolCall("write", {}, { id: "write-1" }), { stopReason: "toolUse" }),
      fauxAssistantMessage(fauxToolCall("write", {}, { id: "write-2" }), { stopReason: "toolUse" }),
      fauxAssistantMessage("first done"),
      fauxAssistantMessage("second done"),
    ]);

    let selectionReads = 0;
    const schema = Type.Object({});
    const writeTool: AgentTool<typeof schema, undefined> = {
      name: "write",
      label: "write",
      description: "write a cell",
      parameters: schema,
      execute: async () => {
        // Pi's own write, then the user adds a sheet and edits it while Pi works.
        await workbook.edit("sheet-1", "C1", "ThisLocalAddin");
        if (!workbook.sheets.has("sheet-2")) {
          await workbook.addSheet("sheet-2", "Added");
          await workbook.edit("sheet-2", "B2", "Unknown");
        }
        return { content: [{ type: "text", text: "wrote C1" }], details: undefined };
      },
    };

    const requests: TranscriptContext[] = [];
    const runtime: { agent?: Agent } = {};
    const agent = new Agent({
      initialState: { model: faux.getModel(), messages: [], tools: [writeTool] },
      convertToLlm: createConvertToLlm({ getContextWindow: () => 200_000 }),
      ...createAutoContext({
        changeTracker: tracker,
        getAgent: () => runtime.agent,
        getWorkbookId: () => Promise.resolve("wb-1"),
        getBlueprint: () => Promise.resolve(blueprint),
        getFilesSummary: () => Promise.resolve(null),
        readSelection: () => {
          selectionReads += 1;
          return Promise.resolve({ address: "Data!A1", text: `Current selection: read ${selectionReads}` });
        },
      }),
      streamFn: (model, context, options) => {
        requests.push(structuredClone(context));
        return models.streamSimple(model, context, options);
      },
    });

    runtime.agent = agent;
    await agent.prompt("write C1");
    await workbook.edit("sheet-1", "D4", "Unknown");
    await agent.prompt("again");

    assert.equal(requests.length, 4);
    const [first, second, third, fourth] = requests.map(autoContextTexts);

    // First call: the workbook note, then the selection, both ahead of nothing stale.
    assert.equal(first?.length, 2);
    assert.match(first?.[0] ?? "", /Workbook structure \(initial\)/u);
    assert.doesNotMatch(first?.[0] ?? "", /Added/u);

    // Second call (after the user added a sheet and edited it): a structure
    // note listing the new sheet and an edit note that names the user's edit
    // on it but not Pi's own write.
    const editNote = second?.find((text) => text.includes("Edits made in Excel"));
    assert.ok(editNote, "edit note sent mid-turn");
    assert.match(editNote, /- Added: B2/u);
    assert.doesNotMatch(editNote, /C1/u);
    assert.ok(second?.some((text) => /Workbook structure \(structure changed\)[\s\S]*Added/u.test(text)));

    // Third call, same turn: notes from earlier calls are still there, once each.
    assert.equal(third?.filter((text) => text.includes("- Added: B2")).length, 1);
    assert.equal(third?.filter((text) => text.includes("Workbook structure")).length, 2);

    // Next turn: the user's later edit is reported; nothing is re-sent.
    assert.equal(fourth?.filter((text) => text.includes("Workbook structure")).length, 2);
    assert.equal(fourth?.filter((text) => /- Data: D4/u.test(text)).length, 1);

    // Cache shape: everything except the trailing selection is a prefix of the next request.
    for (let i = 0; i + 1 < requests.length; i += 1) {
      const current = requests[i]?.messages ?? [];
      const next = requests[i + 1]?.messages ?? [];
      const stored = current.slice(0, -1);
      assert.deepEqual(next.slice(0, stored.length), stored, `request ${i + 1} is a prefix of request ${i + 2}`);
    }
    for (const request of requests) {
      const last = request.messages.at(-1);
      assert.ok(last);
      assert.match(textOf(last), /^\[Auto-context\] Current selection: read/u);
    }
    assert.equal(selectionReads, 4);

    // Notes are stored in the transcript (sessions persist them); the selection is not.
    const stored = agent.state.messages.filter((message: AgentMessage) => message.role === "autoContext");
    assert.deepEqual(stored.map((message) => message.role === "autoContext" ? message.kind : ""), ["workbook", "workbook", "changes", "changes"]);
  } finally {
    uninstall();
  }
});

void test("every chat tab hears about the same user edit", async () => {
  const workbook = new FakeWorkbook();
  const uninstall = workbook.install();
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    const tab = () => createAutoContext({
      changeTracker: tracker,
      getAgent: () => undefined,
      getWorkbookId: () => Promise.resolve("wb-1"),
      getBlueprint: () => Promise.resolve(""),
      getFilesSummary: () => Promise.resolve(null),
    });
    const tabA = tab();
    const tabB = tab();
    await workbook.edit("sheet-1", "A1", "Unknown");

    const request = { model: fauxProvider().getModel(), thinkingLevel: "off" as const, context: { messages: [] } };
    for (const autoContext of [tabA, tabB]) {
      const update = await autoContext.prepareRequest(request);
      const notes = (update?.context?.messages ?? []).filter((message) => message.role === "autoContext");
      assert.equal(notes.length, 1);
      assert.match(notes[0]?.role === "autoContext" ? notes[0].text : "", /- Data: A1/u);
    }
    // Reported once per tab, not again.
    assert.equal(await tabA.prepareRequest(request), undefined);
  } finally {
    uninstall();
  }
});

void test("the workbook note is re-sent after compaction removed it, and labelled so", async () => {
  const autoContext = createAutoContext({
    changeTracker: { cursor: () => 0, describeSince: () => ({ text: null, cursor: 0 }) },
    getAgent: () => undefined,
    getWorkbookId: () => Promise.resolve("wb-1"),
    getBlueprint: () => Promise.resolve("1. **Data**"),
    getFilesSummary: () => Promise.resolve(null),
  });
  const update = await autoContext.prepareRequest({
    model: fauxProvider().getModel(),
    thinkingLevel: "off",
    context: {
      messages: [
        { role: "compactionSummary", summary: "earlier work", tokensBefore: 10, timestamp: 1 },
        { role: "user", content: "next", timestamp: 2 },
      ],
    },
  });
  const messages = update?.context?.messages ?? [];
  assert.equal(messages[1]?.role, "autoContext");
  assert.match(messages[1]?.role === "autoContext" ? messages[1].text : "", /re-sent after the conversation was compacted/u);
  assert.equal(messages[2]?.role, "user", "the user's prompt stays after the note");
});

void test("the Files summary is sent again whenever it changes, and to every new tab", async () => {
  let files = { summary: "### Workspace\n- a.csv", relevantSignature: "sig-1", hasRelevantFiles: true };
  const transcript: AgentMessage[] = [];
  const newTab = () => createAutoContext({
    changeTracker: { cursor: () => 0, describeSince: () => ({ text: null, cursor: 0 }) },
    getAgent: () => undefined,
    getWorkbookId: () => Promise.resolve(null),
    getBlueprint: () => Promise.resolve(""),
    getFilesSummary: () => Promise.resolve(files),
  });
  const model = fauxProvider().getModel();
  const send = async (autoContext: ReturnType<typeof createAutoContext>, messages: AgentMessage[]) => {
    const update = await autoContext.prepareRequest({ model, thinkingLevel: "off", context: { messages } });
    return (update?.context?.messages ?? messages).filter((m) => m.role === "autoContext");
  };

  const tab = newTab();
  const first = await send(tab, transcript);
  assert.match(first[0]?.role === "autoContext" ? first[0].text : "", /Files workspace \(initial\)/u);
  transcript.push(...first);
  assert.equal((await send(tab, transcript)).length, 1, "unchanged: not re-sent");

  for (const signature of ["sig-2", "sig-3"]) {
    files = { hasRelevantFiles: true, relevantSignature: signature, summary: `### Workspace\n- ${signature}` };
    const notes = await send(tab, transcript);
    const latest = notes.at(-1);
    assert.match(latest?.role === "autoContext" ? latest.text : "", new RegExp(`files changed[\\s\\S]*${signature}`, "u"));
    transcript.splice(0, transcript.length, ...notes);
  }

  const other = await send(newTab(), []);
  assert.match(other[0]?.role === "autoContext" ? other[0].text : "", /Files workspace \(initial\)[\s\S]*sig-3/u);
});

// ---------------------------------------------------------------------------
// Selection context
// ---------------------------------------------------------------------------

function installSelection(options: {
  address: string;
  rowIndex: number;
  columnIndex: number;
  rowCount: number;
  columnCount: number;
  used: { address: string; rowIndex: number; columnIndex: number; rowCount: number; columnCount: number } | null;
  fail?: { code: string; message: string };
}): () => void {
  const sheet = {
    name: "Data",
    getUsedRangeOrNullObject: () => ({ ...options.used, isNullObject: options.used === null, load: () => {} }),
    getRangeByIndexes: (row: number, col: number, rows: number, cols: number) => {
      // Long text in every cell: the sample must stay bounded regardless.
      const values = Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => `Row ${row + r + 1} col ${col + c} ${"x".repeat(200)}`));
      const formats = values.map((r) => r.map(() => "General"));
      return { values, formulas: values, text: values, numberFormat: formats, load: () => {} };
    },
  };
  const selection = { ...options, worksheet: sheet, load: () => {} };
  const previous: unknown = Reflect.get(globalThis, "Excel");
  Reflect.set(globalThis, "Excel", {
    run: <T>(callback: (ctx: unknown) => Promise<T>): Promise<T> => {
      if (options.fail) return Promise.reject(Object.assign(new Error(options.fail.message), { code: options.fail.code }));
      return callback({ workbook: { getSelectedRange: () => selection }, sync: () => Promise.resolve() });
    },
  });
  return () => Reflect.set(globalThis, "Excel", previous);
}

const USED = { address: "Data!A1:Z5000", rowIndex: 0, columnIndex: 0, rowCount: 5000, columnCount: 26 };

void test("a 5,000-row selection sends its size and a bounded sample, not every row", async () => {
  const uninstall = installSelection({ address: "Data!A1:A5000", rowIndex: 0, columnIndex: 0, rowCount: 5000, columnCount: 1, used: USED });
  try {
    const { address, text } = await readSelectionContext();
    assert.equal(address, "Data!A1:A5000");
    assert.match(text, /Selection: Data!A1:A5000 — 5,000 rows × 1 column/u);
    assert.match(text, /\| 5000 \|/u, "last row shown");
    assert.match(text, /rows not shown/u);
    assert.ok(text.length < 6_000, `selection context is ${text.length} chars`);
  } finally {
    uninstall();
  }
});

void test("a whole-column selection still gives its address, size and a sample", async () => {
  const uninstall = installSelection({ address: "Data!A:A", rowIndex: 0, columnIndex: 0, rowCount: 1_048_576, columnCount: 1, used: USED });
  try {
    const { text } = await readSelectionContext();
    assert.match(text, /Selection: Data!A:A — 1,048,576 rows/u);
    assert.match(text, /\| 1 \| Row 1 /u);
    assert.ok(text.length < 6_000);
  } finally {
    uninstall();
  }
});

void test("a selection that cannot be read says why instead of disappearing", async () => {
  const uninstall = installSelection({
    address: "", rowIndex: 0, columnIndex: 0, rowCount: 0, columnCount: 0, used: null,
    fail: { code: "InvalidOperationInCellEditMode", message: "Excel is in cell-editing mode." },
  });
  try {
    const { address, text } = await readSelectionContext();
    assert.equal(address, null);
    assert.match(text, /Unavailable: Excel is in cell-edit mode/u);
  } finally {
    uninstall();
  }
});
