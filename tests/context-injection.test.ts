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
import { marked, type Tokens } from "marked";

import { ChangeTracker, type PiWriteOrigin, type PiWriteWindow } from "../src/context/change-tracker.ts";
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
    const previous: unknown = Reflect.get(globalThis, "Excel");
    Reflect.set(globalThis, "Excel", {
      // A fresh context per run, as Office.js does.
      run: <T>(callback: (ctx: unknown) => Promise<T>): Promise<T> =>
        callback({ workbook: { worksheets }, sync: () => Promise.resolve() }),
    });
    return () => Reflect.set(globalThis, "Excel", previous);
  }

  private async emit(name: string, event: object): Promise<void> {
    const handler = this.handlers.get(name);
    if (!handler) throw new Error(`no ${name} handler registered`);
    await (handler as (event: object) => Promise<void>)(event);
  }

  edit(
    sheetId: string,
    address: string,
    triggerSource: "Unknown" | "ThisLocalAddin",
    options: { source?: "Local" | "Remote"; changeType?: string } = {},
  ): Promise<void> {
    return this.emit("changed", {
      worksheetId: sheetId,
      address,
      changeType: options.changeType ?? "RangeEdited",
      source: options.source ?? "Local",
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
        tracker.notePiWrite({ toolCallId: "w", startedAt: Date.now(), settledAt: Date.now() + 5, address: "Data!C1" }, { kind: "runtime", runtimeId: "tab-a" });
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
        runtimeId: "tab-a",
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
    // The trailing note says it is background; mid-turn it says it is not a new
    // request, and it never restates the request (that looped a mutation).
    for (const [index, request] of requests.entries()) {
      const last = request.messages.at(-1);
      assert.ok(last);
      assert.match(textOf(last), /^\[Auto-context\] Current selection: read/u);
      assert.doesNotMatch(textOf(last), /write C1|again/u, `request ${index + 1} restates the request`);
      const midTurn = index === 1 || index === 2; // calls 2 and 3 follow tool results
      assert.ok(
        textOf(last).endsWith(midTurn
          ? "not a new request. Continue from the tool results above; do not repeat steps that already succeeded.)"
          : "(Background from Excel, not a message from the user.)"),
        `request ${index + 1}: ${textOf(last)}`,
      );
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
      runtimeId: crypto.randomUUID(),
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
    changeTracker: { cursor: () => 0, describeSince: () => Promise.resolve({ text: null, cursor: 0 }) },
    getAgent: () => undefined,
    runtimeId: "tab-a",
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
    changeTracker: { cursor: () => 0, describeSince: () => Promise.resolve({ text: null, cursor: 0 }) },
    getAgent: () => undefined,
    runtimeId: "tab-a",
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
  values?: string[][];
}): () => void {
  const sheet = {
    name: "Data",
    getUsedRangeOrNullObject: () => ({ ...options.used, isNullObject: options.used === null, load: () => {} }),
    getRangeByIndexes: (row: number, col: number, rows: number, cols: number) => {
      // Long text in every cell: the sample must stay bounded regardless.
      const values = Array.from({ length: rows }, (_, r) =>
        Array.from({ length: cols }, (_, c) => options.values?.[row + r]?.[col + c]
          ?? `Row ${row + r + 1} col ${col + c} ${"x".repeat(200)}`));
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

void test("selection Markdown keeps backslashes, pipes and neighboring cells without changing newline wording", async () => {
  const range = { address: "Data!A1:D1", rowIndex: 0, columnIndex: 0, rowCount: 1, columnCount: 4 };
  const uninstall = installSelection({
    ...range,
    used: range,
    values: [[String.raw`back\|pipe`, "control", "two\nlines", String.raw`two\\|pipes`]],
  });
  try {
    const { text } = await readSelectionContext();
    const table = marked.lexer(text).find((token): token is Tokens.Table => token.type === "table");
    assert.ok(table);
    assert.deepEqual(table.rows[0]?.map((cell) => marked.parseInline(cell.text)), [
      "1", String.raw`back\|pipe`, "control", "two lines", String.raw`two\\|pipes`,
    ]);
  } finally {
    uninstall();
  }
});

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

void test("with a co-author present, Pi's own writes are left out and the co-author's edits are attributed to them", async () => {
  const workbook = new FakeWorkbook();
  const uninstall = workbook.install();
  try {
    let structureChanges = 0;
    const tracker = new ChangeTracker({
      onStructureChange: () => { structureChanges += 1; },
      getPiWriteAddresses: (ids) => Promise.resolve(new Map(ids.map((id) => [id, ["Data!A6:B6"]]))),
    });
    await tracker.start();
    const autoContext = createAutoContext({
      changeTracker: tracker,
      getAgent: () => undefined,
      runtimeId: "tab-a",
      getWorkbookId: () => Promise.resolve("wb-1"),
      getBlueprint: () => Promise.resolve(""),
      getFilesSummary: () => Promise.resolve(null),
    });

    // Once a co-author has joined, Pi's writes arrive as "Unknown", like the user's.
    const window = { toolCallId: "call-1", toolName: "write_cells", startedAt: Date.now(), settledAt: Number.POSITIVE_INFINITY };
    tracker.notePiWrite(window, { kind: "runtime", runtimeId: "tab-a" });
    await workbook.edit("sheet-1", "A6:B6", "Unknown");
    await workbook.edit("sheet-1", "D9", "Unknown"); // the user, while Pi wrote
    await workbook.edit("sheet-1", "3:3", "Unknown", { source: "Remote", changeType: "RowInserted" });
    await workbook.edit("sheet-1", "C2", "Unknown", { source: "Remote" });
    window.settledAt = Date.now();

    const request = { model: fauxProvider().getModel(), thinkingLevel: "off" as const, context: { messages: [] } };
    const update = await autoContext.prepareRequest(request);
    const note = (update?.context?.messages ?? []).find((message) => message.role === "autoContext");
    const text = note?.role === "autoContext" ? note.text : "";

    assert.doesNotMatch(text, /A6:B6/u, "Pi's own write is not reported");
    assert.match(text, /By the user:\n- Data: D9\n/u);
    assert.match(text, /By a co-author[^\n]*:\n- Data: 3:3 \(rows inserted\), C2/u);
    assert.match(text, /A co-author inserted or deleted cells on Data: cells there may have moved/u);
    assert.equal(structureChanges, 1, "the co-author's row insert rebuilds the blueprint");
  } finally {
    uninstall();
  }
});

void test("a backup restore, another tab's write and an extension's write are reported; only this tab's own write is left out", async () => {
  const workbook = new FakeWorkbook();
  const uninstall = workbook.install();
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    const tab = (runtimeId: string) => createAutoContext({
      changeTracker: tracker,
      getAgent: () => undefined,
      runtimeId,
      getWorkbookId: () => Promise.resolve("wb-1"),
      getBlueprint: () => Promise.resolve(""),
      getFilesSummary: () => Promise.resolve(null),
    });
    const tabA = tab("tab-a");
    const tabB = tab("tab-b");

    // Writes are serialized by the workbook lock: one window at a time.
    const write = async (origin: PiWriteOrigin, address: string, eventAddress = address) => {
      await new Promise((resolve) => setTimeout(resolve, 2)); // writes take time
      const window: PiWriteWindow = { toolCallId: crypto.randomUUID(), startedAt: Date.now(), settledAt: Number.POSITIVE_INFINITY, address: address.includes("!") ? address : `Data!${address}` };
      tracker.notePiWrite(window, origin);
      await workbook.edit("sheet-1", eventAddress, "ThisLocalAddin");
      window.settledAt = Date.now();
    };
    await write({ kind: "runtime", runtimeId: "tab-a" }, "F3");
    await write({ kind: "restore" }, "B2:D4");
    // Tab A's write whose event names other cells than the recorded range (a row insert, say).
    await write({ kind: "runtime", runtimeId: "tab-a" }, "Data!J1", "9:9");
    // The user restores tab A's G7 write: listed once, as the restore.
    await write({ kind: "runtime", runtimeId: "tab-a" }, "G7");
    await write({ kind: "restore" }, "G7");
    await new Promise((resolve) => setTimeout(resolve, 5));
    tracker.notePiWrite({ toolCallId: "far", startedAt: 0, settledAt: 1 }, { kind: "runtime", runtimeId: "tab-a" });
    // An add-in write no window covers (an extension). Moves the clock past the grace window.
    const realNow = Date.now;
    Date.now = () => realNow() + 60_000;
    try {
      await workbook.edit("sheet-1", "E1", "ThisLocalAddin");
    } finally {
      Date.now = realNow;
    }

    const request = { model: fauxProvider().getModel(), thinkingLevel: "off" as const, context: { messages: [] } };
    const noteFor = async (autoContext: ReturnType<typeof tab>) => {
      const update = await autoContext.prepareRequest(request);
      const note = (update?.context?.messages ?? []).find((message) => message.role === "autoContext");
      return note?.role === "autoContext" ? note.text : "";
    };
    const a = await noteFor(tabA);
    const b = await noteFor(tabB);

    assert.doesNotMatch(a, /F3|9:9/u, "tab A's own writes are left out of its note");
    assert.match(a, /Restored from a backup by the user[^\n]*\n- Data: B2:D4, G7\n/u);
    assert.match(a, /By Pi outside this chat[^\n]*\n- Data: E1/u);
    assert.doesNotMatch(a, /By the user/u);

    assert.match(b, /By Pi in another chat tab:\n- Data: F3, 9:9\n/u);
    assert.match(b, /Restored from a backup by the user[^\n]*\n- Data: B2:D4, G7\n/u);
    assert.equal(b.match(/G7/gu)?.length, 1, "a cell is listed once, under whoever changed it last");
  } finally {
    uninstall();
  }
});

void test("Pi's own structure change is not sent back as a new workbook note mid-turn; the next turn gets it", async () => {
  const workbook = new FakeWorkbook();
  const uninstall = workbook.install();
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    let blueprint = "1. **Data** (10 rows)";
    const autoContext = createAutoContext({
      changeTracker: tracker,
      getAgent: () => undefined,
      runtimeId: "tab-a",
      getWorkbookId: () => Promise.resolve("wb-1"),
      getBlueprint: () => Promise.resolve(blueprint),
      getFilesSummary: () => Promise.resolve(null),
    });
    const model = fauxProvider().getModel();
    const notesFor = async (messages: AgentMessage[]) => {
      const update = await autoContext.prepareRequest({ model, thinkingLevel: "off", context: { messages } });
      return (update?.context?.messages ?? []).filter((message) => message.role === "autoContext" && !messages.includes(message));
    };

    const prompt: AgentMessage = { role: "user", content: "Insert one row at row 12 of Data.", timestamp: 1 };
    const first = await notesFor([prompt]);
    assert.equal(first.length, 1);
    const history: AgentMessage[] = [prompt, ...first];

    // Pi's insert succeeds and changes the structure.
    blueprint = "1. **Data** (11 rows)";
    autoContext.noteOwnStructureChange();
    history.push(
      fauxAssistantMessage(fauxToolCall("modify_structure", {}, { id: "ins-1" }), { stopReason: "toolUse" }),
      { role: "toolResult", toolCallId: "ins-1", toolName: "modify_structure", content: [{ type: "text", text: "Inserted row 12" }], isError: false, timestamp: 2 },
    );
    assert.deepEqual(await notesFor(history), [], "no fresh structure note after Pi's own insert");

    // The next user prompt gets the new structure.
    history.push(fauxAssistantMessage("Done."), { role: "user", content: "Thanks", timestamp: 3 });
    const next = await notesFor(history);
    assert.equal(next.length, 1);
    assert.match(next[0]?.role === "autoContext" ? next[0].text : "", /structure changed[\s\S]*11 rows/u);
  } finally {
    uninstall();
  }
});
