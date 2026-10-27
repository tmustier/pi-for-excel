/**
 * UI Gallery — scripted live turn.
 *
 * Mounts the real <pi-sidebar>, a real pi-agent-core Agent and the real tool
 * renderers, fed by pi-ai's faux provider instead of a model. The turn
 * streams thinking, text and three parallel read_range calls, then a second
 * response with code fences. Tool calls finish when released: by timer in
 * "Play" mode, or by hand (`window.piTranscriptHarness.release(id)`) so
 * browser contracts can observe each state deterministically.
 *
 * No Office.js dependency; cell links fail here, which exercises the
 * failed-navigation path.
 */

import { Agent, type AgentTool } from "@earendil-works/pi-agent-core";
import {
  createFauxCore,
  fauxAssistantMessage,
  fauxText,
  fauxThinking,
  fauxToolCall,
  Type,
} from "@earendil-works/pi-ai";

import { PiSidebar } from "./ui/pi-sidebar.js";
import "./ui/tool-renderers.js";
import "./ui/message-renderers.js";

export const TRANSCRIPT_TOOL_CALL_IDS = ["read-1", "read-2", "read-3"] as const;

/** Release delays used by the Play button (ms after the tool starts). */
const PLAY_DELAYS_MS: Record<string, number> = {
  "read-1": 1000,
  "read-2": 1500,
  "read-3": 4000,
};

const FIRST_RESPONSE_TEXT = "I'll read the three blocks of the sheet.\nThen I'll summarise them.";
const FINAL_RESPONSE_TEXT = [
  "All three blocks hold monthly totals.",
  "Q1 is highest.",
  "",
  "```c++",
  "int total = a + b;",
  "```",
  "",
  "```f#",
  "let total = a + b",
  "```",
].join("\n");

interface Gate {
  promise: Promise<void>;
  resolve: () => void;
}

function createGate(): Gate {
  let resolve: () => void = () => {};
  const promise = new Promise<void>((done) => { resolve = done; });
  return { promise, resolve };
}

export interface TranscriptHarness {
  readonly agent: Agent;
  readonly sidebar: PiSidebar;
  /** Start the scripted turn. With `autoRelease`, tools finish on PLAY_DELAYS_MS. */
  start(options?: { autoRelease?: boolean }): Promise<void>;
  /** Finish one tool call (manual mode). */
  release(toolCallId: string): void;
  reset(): void;
}

export function mountTranscriptHarness(host: HTMLElement): TranscriptHarness {
  const faux = createFauxCore({
    provider: "gallery",
    models: [{ id: "gallery-model", name: "Gallery model", reasoning: true }],
    tokensPerSecond: 40,
  });

  let gates = new Map<string, Gate>();
  let autoRelease = false;

  const readRangeParameters = Type.Object({ range: Type.String() });
  const readRange: AgentTool<typeof readRangeParameters> = {
    name: "read_range",
    label: "Read range",
    description: "Scripted read_range for the UI gallery.",
    parameters: readRangeParameters,
    async execute(toolCallId, params) {
      const gate = gates.get(toolCallId) ?? createGate();
      gates.set(toolCallId, gate);
      const delay = PLAY_DELAYS_MS[toolCallId];
      if (autoRelease && delay !== undefined) setTimeout(gate.resolve, delay);
      await gate.promise;
      return {
        content: [{ type: "text", text: `Read ${params.range}: 12 rows of monthly totals.` }],
        details: undefined,
      };
    },
  };

  const agent = new Agent({
    initialState: {
      model: faux.getModel(),
      thinkingLevel: "high",
      tools: [readRange],
    },
    streamFn: faux.streamSimple,
    getApiKey: () => "gallery",
  });

  const sidebar = new PiSidebar();
  sidebar.agent = agent;
  host.appendChild(sidebar);

  // Factories, so each response is stamped when it is requested, as a
  // provider stamps it when it starts streaming.
  function script(): void {
    faux.setResponses([
      () => fauxAssistantMessage([
        fauxThinking("The user wants a summary. I should read each block of the sheet before answering, so I will request all three ranges at once."),
        fauxText(FIRST_RESPONSE_TEXT),
        fauxToolCall("read_range", { range: "Sheet1!A1:B12" }, { id: "read-1" }),
        fauxToolCall("read_range", { range: "Sheet1!C1:D12" }, { id: "read-2" }),
        fauxToolCall("read_range", { range: "Sheet1!E1:F12" }, { id: "read-3" }),
      ], { stopReason: "toolUse" }),
      () => fauxAssistantMessage([
        fauxThinking("All three reads came back; Q1 is the largest."),
        fauxText(FINAL_RESPONSE_TEXT),
      ]),
    ]);
  }

  return {
    agent,
    sidebar,
    async start(options = {}) {
      autoRelease = options.autoRelease ?? false;
      gates = new Map();
      for (const id of TRANSCRIPT_TOOL_CALL_IDS) gates.set(id, createGate());
      script();
      await agent.prompt("Summarise Sheet1.\nKeep it short.");
    },
    release(toolCallId) {
      const gate = gates.get(toolCallId) ?? createGate();
      gates.set(toolCallId, gate);
      gate.resolve();
    },
    reset() {
      agent.reset();
      sidebar.syncFromAgent();
    },
  };
}

declare global {
  interface Window {
    piTranscriptHarness?: TranscriptHarness;
  }
}
