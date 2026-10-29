import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import type { Api, AssistantMessageEventStream, Model, StreamOptions, TranscriptContext } from "@earendil-works/pi-ai";
import { normalizeContext } from "@earendil-works/pi-ai";
import { Type } from "typebox";
import { Value } from "typebox/value";
import type { WebSearchToolDetails } from "./web-search.js";

export function supportsNativeWebSearch(model: Model<Api>): boolean {
  switch (model.provider) {
    case "openai": case "xai": return model.api === "openai-responses";
    case "openai-codex": return model.api === "openai-codex-responses";
    case "anthropic": return model.api === "anthropic-messages";
    case "google": return model.api === "google-generative-ai";
    default: return false;
  }
}

const payloadSchema = Type.Object({
  model: Type.String(),
  config: Type.Optional(Type.Object({ maxOutputTokens: Type.Optional(Type.Number()) }, { additionalProperties: true })),
}, { additionalProperties: true });

export async function runNativeWebSearch(args: {
  model: Model<Api>;
  apiKey: string;
  query: string;
  sessionId: string;
  signal: AbortSignal | undefined;
  stream: (model: Model<Api>, context: TranscriptContext, options?: StreamOptions) => Promise<AssistantMessageEventStream>;
}): Promise<AgentToolResult<WebSearchToolDetails>> {
  if (!supportsNativeWebSearch(args.model)) throw new Error("Native search is unavailable for this model.");
  const provider = args.model.provider;
  const stream = await args.stream(args.model, normalizeContext({
    messages: [{ role: "user", content: `Search the web for: ${args.query}\nReturn concise findings with source URLs written explicitly.`, timestamp: Date.now() }],
  }), {
    apiKey: args.apiKey,
    sessionId: `${args.sessionId}::native-search`,
    ...(args.signal ? { signal: args.signal } : {}),
    onPayload: (payload) => {
      if (!Value.Check(payloadSchema, payload)) throw new Error("Invalid native search request payload.");
      if (provider === "google") return { ...payload, config: { ...payload.config, tools: [{ googleSearch: {} }] } };
      if (provider === "anthropic") return {
        ...payload,
        tools: [{ type: "web_search_20250305", name: "web_search", max_uses: 3 }],
      };
      return { ...payload, tools: [{ type: "web_search" }], tool_choice: "required" };
    },
  });
  const result = await stream.result();
  if (result.stopReason === "error" || result.stopReason === "aborted") {
    throw new Error(result.errorMessage ?? "Native web search failed.");
  }
  const text = result.content.filter((block) => block.type === "text").map((block) => block.text).join("\n");
  if (!text.trim()) throw new Error("Native web search returned no findings.");
  return {
    content: [{ type: "text", text: `Web search via ${provider}\n\n${text}` }],
    details: { kind: "web_search", ok: true, provider, query: args.query, sentQuery: args.query, maxResults: 5 },
  };
}
