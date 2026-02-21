import type {
  Agent,
  AgentEvent,
  AgentTool,
  AgentToolResult,
  AgentToolUpdateCallback,
} from "@mariozechner/pi-agent-core";
import type { Static, TSchema } from "@sinclair/typebox";

import type { ExtensionCapability } from "../extensions/permissions.js";
import type { ExtensionWidgetPlacement } from "../extensions/internal/widget-surface.js";

export interface ExtensionCommand {
  description: string;
  /**
   * Allow running this command while the active runtime is busy/streaming.
   * Defaults to true for extension commands.
   */
  busyAllowed?: boolean;
  handler: (args: string) => void | Promise<void>;
}

export type ExtensionCleanup = () => void | Promise<void>;

export interface ExtensionToolDefinition<TParameters extends TSchema = TSchema, TDetails = unknown> {
  description: string;
  parameters: TParameters;
  label?: string;
  execute: (
    params: Static<TParameters>,
    signal?: AbortSignal,
    onUpdate?: AgentToolUpdateCallback<TDetails>,
  ) => Promise<AgentToolResult<TDetails>> | AgentToolResult<TDetails>;
}

export interface OverlayAPI {
  /** Show an HTML element as a full-screen overlay */
  show(el: HTMLElement): void;
  /** Remove the overlay */
  dismiss(): void;
}

export type WidgetPlacement = ExtensionWidgetPlacement;

export interface WidgetUpsertSpec {
  id: string;
  el: HTMLElement;
  title?: string;
  placement?: WidgetPlacement;
  order?: number;
  collapsible?: boolean;
  collapsed?: boolean;
  minHeightPx?: number | null;
  maxHeightPx?: number | null;
}

export interface WidgetAPI {
  /** Show an HTML element as an inline widget above the input area */
  show(el: HTMLElement): void;
  /** Remove the legacy widget */
  dismiss(): void;
  /** Add or update a named widget (Widget API v2; gated by experiment). */
  upsert(spec: WidgetUpsertSpec): void;
  /** Remove a specific named widget (Widget API v2; gated by experiment). */
  remove(id: string): void;
  /** Remove all widgets owned by the extension (Widget API v2; gated by experiment). */
  clear(): void;
}

export interface LlmCompletionMessage {
  role: "user" | "assistant";
  content: string;
}

export interface LlmCompletionRequest {
  model?: string;
  systemPrompt?: string;
  messages: LlmCompletionMessage[];
  maxTokens?: number;
}

export interface LlmCompletionResult {
  content: string;
  model: string;
  usage?: {
    inputTokens: number;
    outputTokens: number;
  };
}

export interface LlmAPI {
  complete(request: LlmCompletionRequest): Promise<LlmCompletionResult>;
}

export interface HttpRequestOptions {
  method?: "GET" | "POST" | "PUT" | "PATCH" | "DELETE" | "HEAD";
  headers?: Record<string, string>;
  body?: string;
  timeoutMs?: number;
}

export interface HttpResponse {
  status: number;
  statusText: string;
  headers: Record<string, string>;
  body: string;
}

export interface HttpAPI {
  fetch(url: string, options?: HttpRequestOptions): Promise<HttpResponse>;
}

export interface StorageAPI {
  get(key: string): Promise<unknown>;
  set(key: string, value: unknown): Promise<void>;
  delete(key: string): Promise<void>;
  keys(): Promise<string[]>;
}

export interface ClipboardAPI {
  writeText(text: string): Promise<void>;
}

export interface ExtensionAgentAPI {
  readonly raw: Agent;
  injectContext(content: string): void;
  steer(content: string): void;
  followUp(content: string): void;
}

export interface SkillSummary {
  name: string;
  description: string;
  sourceKind: string;
}

export interface SkillsAPI {
  list(): Promise<SkillSummary[]>;
  read(name: string): Promise<string>;
  install(name: string, markdown: string): Promise<void>;
  uninstall(name: string): Promise<void>;
}

export interface DownloadAPI {
  download(filename: string, content: string, mimeType?: string): void;
}

export interface ExcelExtensionAPI {
  /** Register a slash command */
  registerCommand(name: string, cmd: ExtensionCommand): void;
  /** Register a custom tool callable by the agent */
  registerTool(name: string, tool: ExtensionToolDefinition): void;
  /** Remove a previously registered custom tool */
  unregisterTool(name: string): void;
  /** Agent access and steering APIs */
  readonly agent: ExtensionAgentAPI;
  /** LLM completion API via host mediation */
  llm: LlmAPI;
  /** HTTP fetch API via host mediation */
  http: HttpAPI;
  /** Persistent extension-scoped key/value storage */
  storage: StorageAPI;
  /** Clipboard operations */
  clipboard: ClipboardAPI;
  /** Skill catalog read/write helpers */
  skills: SkillsAPI;
  /** Trigger browser downloads */
  download: DownloadAPI;
  /** Show/dismiss full-screen overlay UI */
  overlay: OverlayAPI;
  /** Show/dismiss inline widget above input (messages still visible above) */
  widget: WidgetAPI;
  /** Show a toast notification */
  toast(message: string): void;
  /** Subscribe to agent events */
  onAgentEvent(handler: (ev: AgentEvent) => void): () => void;
}

export interface CreateExtensionAPIOptions {
  getAgent: () => Agent;
  registerCommand?: (name: string, cmd: ExtensionCommand) => void;
  registerTool?: (tool: AgentTool) => void;
  unregisterTool?: (name: string) => void;
  subscribeAgentEvents?: (handler: (ev: AgentEvent) => void) => () => void;
  llmComplete?: (request: LlmCompletionRequest) => Promise<LlmCompletionResult>;
  httpFetch?: (url: string, options?: HttpRequestOptions) => Promise<HttpResponse>;
  storageGet?: (key: string) => Promise<unknown>;
  storageSet?: (key: string, value: unknown) => Promise<void>;
  storageDelete?: (key: string) => Promise<void>;
  storageKeys?: () => Promise<string[]>;
  clipboardWriteText?: (text: string) => Promise<void>;
  injectAgentContext?: (content: string) => void;
  steerAgent?: (content: string) => void;
  followUpAgent?: (content: string) => void;
  listSkills?: () => Promise<SkillSummary[]>;
  readSkill?: (name: string) => Promise<string>;
  installSkill?: (name: string, markdown: string) => Promise<void>;
  uninstallSkill?: (name: string) => Promise<void>;
  downloadFile?: (filename: string, content: string, mimeType?: string) => void;
  toast?: (message: string) => void;
  isCapabilityEnabled?: (capability: ExtensionCapability) => boolean;
  formatCapabilityError?: (capability: ExtensionCapability) => string;
  extensionOwnerId?: string;
  widgetApiV2Enabled?: boolean;
}
