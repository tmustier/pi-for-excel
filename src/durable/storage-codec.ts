import { isJsonValue, type JsonValue } from "@earendil-works/chord";
import { assertValidOp, type Op } from "@earendil-works/chord/delta";
import type { StorageWrite } from "@earendil-works/pi-durable";
import { Type } from "typebox";

export const counterSchema = Type.Integer({ minimum: 1, maximum: Number.MAX_SAFE_INTEGER });
const jsonSchema = Type.Refine(Type.Unsafe<JsonValue>({}), isJsonValue);
const objectSchema = Type.Record(Type.String(), jsonSchema);
const opSchema = Type.Refine(Type.Unsafe<Op>({}), (op) => {
  try { assertValidOp(op); return true; } catch { return false; }
});
const optionalId = Type.Optional(counterSchema);
const absent = Type.Optional(Type.Never());
const textSchema = Type.Object({ type: Type.Literal("text"), text: Type.String(), textSignature: Type.Optional(Type.String()) });
const imageSchema = Type.Object({ type: Type.Literal("image"), data: Type.String(), mimeType: Type.String() });
const usageSchema = Type.Object({
  input: Type.Number(), output: Type.Number(), cacheRead: Type.Number(), cacheWrite: Type.Number(),
  cacheWrite1h: Type.Optional(Type.Number()), reasoning: Type.Optional(Type.Number()), totalTokens: Type.Number(),
  cost: Type.Object({ input: Type.Number(), output: Type.Number(), cacheRead: Type.Number(), cacheWrite: Type.Number(), total: Type.Number() }),
});
const errorSchema = Type.Object({ message: Type.String(), detail: Type.Optional(jsonSchema) });
const messageSchema = Type.Union([
  Type.Object({ role: Type.Literal("user"), timestamp: Type.Number(), content: Type.Union([Type.String(), Type.Array(Type.Union([textSchema, imageSchema]))]) }),
  Type.Object({
    role: Type.Literal("system"), timestamp: Type.Number(), content: Type.Union([Type.String(), Type.Array(textSchema)]),
    sections: Type.Optional(Type.Record(Type.String(), Type.Union([Type.String(), Type.Null()]))),
    toolsAdded: Type.Optional(Type.Array(Type.Object({
      name: Type.String(), description: Type.String(), parameters: objectSchema,
      constrainedSampling: Type.Optional(Type.Union([
        Type.Literal(false), Type.Object({ type: Type.Literal("json_schema"), strict: Type.Union([Type.Literal("prefer"), Type.Literal("require")]) }),
        Type.Object({ type: Type.Literal("grammar"), variants: Type.Object({ openai_lark: Type.Optional(Type.String()), openai_regex: Type.Optional(Type.String()) }) }),
      ])),
    }))),
    toolsRemoved: Type.Optional(Type.Array(Type.Object({ name: Type.String() }))),
  }),
  Type.Object({
    role: Type.Literal("assistant"), timestamp: Type.Number(), api: Type.String(), provider: Type.String(), model: Type.String(),
    content: Type.Array(Type.Union([textSchema,
      Type.Object({ type: Type.Literal("thinking"), thinking: Type.String(), thinkingSignature: Type.Optional(Type.String()), redacted: Type.Optional(Type.Boolean()) }),
      Type.Object({ type: Type.Literal("toolCall"), id: Type.String(), name: Type.String(), arguments: objectSchema,
        thoughtSignature: Type.Optional(Type.String()), namespace: Type.Optional(Type.String()), argumentsError: Type.Optional(Type.String()) }),
    ])),
    responseModel: Type.Optional(Type.String()), responseId: Type.Optional(Type.String()), providerThinkingLevel: Type.Optional(Type.String()),
    thinkingLevel: Type.Optional(Type.Union([Type.Literal("off"), Type.Literal("minimal"), Type.Literal("low"), Type.Literal("medium"), Type.Literal("high"), Type.Literal("xhigh"), Type.Literal("max")])),
    usage: usageSchema,
    stopReason: Type.Union([Type.Literal("pending"), Type.Literal("stop"), Type.Literal("length"), Type.Literal("toolUse"), Type.Literal("error"), Type.Literal("aborted"), Type.Literal("deferred")]),
    deferred: Type.Optional(Type.Object({ provider: Type.String(), modelId: Type.String(), api: Type.String(), id: Type.String(),
      expiresAt: Type.Optional(Type.Number()), pollAfterMs: Type.Optional(Type.Number()), data: Type.Optional(jsonSchema) })),
    diagnostics: Type.Optional(Type.Array(Type.Object({ type: Type.String(), timestamp: Type.Number(), details: Type.Optional(objectSchema),
      error: Type.Optional(Type.Object({ message: Type.String(), name: Type.Optional(Type.String()), stack: Type.Optional(Type.String()), code: Type.Optional(Type.Union([Type.String(), Type.Number()])) })) }))),
    errorMessage: Type.Optional(Type.String()), rawStopReason: Type.Optional(Type.String()), endTurn: Type.Optional(Type.Boolean()), durationMs: Type.Optional(Type.Number()),
  }),
  Type.Object({
    role: Type.Literal("toolResult"), timestamp: Type.Number(), toolCallId: Type.String(), toolName: Type.String(),
    content: Type.Array(Type.Union([textSchema, imageSchema])), isError: Type.Boolean(), details: Type.Optional(jsonSchema),
    usage: Type.Optional(usageSchema), durationMs: Type.Optional(Type.Number()),
    nestedCalls: Type.Optional(Type.Object({ complete: Type.Boolean(), calls: Type.Array(Type.Object({
      id: Type.String(), name: Type.String(), arguments: Type.Optional(objectSchema), argumentsBytes: Type.Optional(Type.Number()),
      status: Type.Union([Type.Literal("ok"), Type.Literal("error"), Type.Literal("unfinished")]), durationMs: Type.Optional(Type.Number()), error: Type.Optional(Type.String()),
    })) })),
  }),
]);
export const conversationSchema = Type.Object({ id: counterSchema,
  parent: Type.Optional(Type.Object({ conversationId: counterSchema, at: counterSchema })),
  owner: Type.Optional(Type.Object({ conversationId: counterSchema, taskId: counterSchema })),
});
export const entrySchema = Type.Object({
  id: counterSchema, conversationId: counterSchema, kind: Type.String(), model: Type.Optional(Type.Array(messageSchema)),
  data: Type.Optional(jsonSchema), head: optionalId, byTaskId: optionalId,
  edits: Type.Optional(Type.Array(Type.Union([
    Type.Object({ target: counterSchema, action: Type.Literal("omit"), messages: absent }),
    Type.Object({ target: counterSchema, action: Type.Literal("replace"), messages: Type.Array(messageSchema) }),
  ]))),
});
const outcomeSchema = Type.Union([
  Type.Object({ status: Type.Literal("completed"), result: jsonSchema, error: absent, reason: absent }),
  Type.Object({ status: Type.Literal("failed"), error: errorSchema, result: Type.Optional(jsonSchema), reason: absent }),
  Type.Object({ status: Type.Literal("aborted"), reason: Type.Optional(Type.String()), result: Type.Optional(jsonSchema), error: absent }),
  Type.Object({ status: Type.Literal("orphaned"), reason: Type.String(), error: absent, result: absent }),
  Type.Object({ status: Type.Literal("faulted"), error: errorSchema, result: absent, reason: absent }),
]);
const taskIdentity = {
  id: counterSchema, conversationId: counterSchema, kind: Type.String(), version: Type.Number(), input: jsonSchema,
  owner: optionalId, background: Type.Boolean(), abortRequested: Type.Boolean(), abortReason: Type.Optional(Type.Literal("restart")),
  abandonOnRestart: Type.Optional(Type.Literal(true)), startedAt: Type.Optional(Type.Number()), endedAt: Type.Optional(Type.Number()),
};
export const taskSchema = Type.Union([
  Type.Object({ ...taskIdentity, memos: Type.Optional(objectSchema), state: Type.Union([
    Type.Object({ status: Type.Union([Type.Literal("pending"), Type.Literal("running")]), checkpoint: jsonSchema, outcome: absent }),
    Type.Object({ status: Type.Literal("waiting"), checkpoint: jsonSchema, on: Type.Array(counterSchema), policy: Type.Union([Type.Literal("failFast"), Type.Literal("allSettled")]), outcome: absent }),
  ]) }),
  Type.Object({ ...taskIdentity, memos: absent, state: Type.Object({ status: Type.Union([Type.Literal("completing"), Type.Literal("terminal")]), outcome: outcomeSchema, checkpoint: absent }) }),
]);
const submissionIdentity = { id: counterSchema, conversationId: counterSchema, requestId: Type.Optional(Type.String()) };
export const submissionSchema = Type.Union([
  Type.Object({ ...submissionIdentity, type: Type.Union([Type.Literal("input"), Type.Literal("write")]), status: Type.Literal("queued"), entry: absent, answer: absent, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("input"), status: Type.Literal("placed"), entry: counterSchema, answer: absent, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("input"), status: Type.Literal("done"), entry: counterSchema, answer: counterSchema, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("write"), status: Type.Literal("done"), entry: counterSchema, answer: absent, reason: absent, detail: absent }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("input"), status: Type.Literal("unanswered"), entry: optionalId, answer: absent, reason: Type.String(), detail: Type.Optional(jsonSchema) }),
  Type.Object({ ...submissionIdentity, type: Type.Literal("write"), status: Type.Literal("unanswered"), entry: absent, answer: absent, reason: Type.String(), detail: Type.Optional(jsonSchema) }),
]);
const documentIdentity = { id: counterSchema, kind: Type.String(), key: Type.Optional(Type.String()) };
export const documentSchema = Type.Union([
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("session") }), history: absent, fork: absent }),
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("task"), taskId: counterSchema }), history: absent, fork: absent }),
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("conversation"), conversationId: counterSchema }), history: Type.Literal("latest"), fork: Type.Union([Type.Literal("current"), Type.Literal("initial")]) }),
  Type.Object({ ...documentIdentity, scope: Type.Object({ kind: Type.Literal("conversation"), conversationId: counterSchema }), history: Type.Literal("rewindable"), fork: Type.Union([Type.Literal("asOf"), Type.Literal("current"), Type.Literal("initial")]) }),
]);
const baseSchema = Type.Object({ kind: Type.Literal("base"), version: Type.Number(), value: objectSchema });
export const contentSchema = Type.Union([baseSchema, Type.Object({ kind: Type.Literal("delta"), version: Type.Number(), ops: Type.Array(opSchema) })]);
// The schema checks the 1.3.0 wire contract. Unsafe supplies only erased ID brands
// and readonly TypeScript annotations; it does not bypass runtime validation.
export const writesSchema = Type.Unsafe<readonly StorageWrite[]>(Type.Array(Type.Union([
  Type.Object({ type: Type.Literal("conversation"), value: conversationSchema }),
  Type.Object({ type: Type.Literal("entry"), value: entrySchema }),
  Type.Object({ type: Type.Literal("task"), value: taskSchema }),
  Type.Object({ type: Type.Literal("submission"), value: submissionSchema }),
  Type.Object({ type: Type.Literal("document.create"), record: documentSchema, content: baseSchema }),
  Type.Object({ type: Type.Literal("document.change"), id: counterSchema, content: contentSchema }),
  Type.Object({ type: Type.Literal("document.retire"), id: counterSchema }),
])));
export const journalSchema = Type.Object({
  seq: counterSchema,
  writes: Type.String(),
  digest: Type.String(),
}, { additionalProperties: false });

export function completed(transaction: IDBTransaction): Promise<void> {
  return new Promise((resolve, reject) => {
    transaction.oncomplete = () => resolve();
    transaction.onabort = () => reject(transaction.error ?? new Error("Durable storage transaction aborted"));
    transaction.onerror = () => { /* onabort is the final transaction outcome. */ };
  });
}

export async function digest(text: string): Promise<string> {
  const bytes = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(bytes), (byte) => byte.toString(16).padStart(2, "0")).join("");
}
