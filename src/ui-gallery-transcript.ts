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
  start(options?: { autoRelease?: boolean }): Promise<void>;
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
