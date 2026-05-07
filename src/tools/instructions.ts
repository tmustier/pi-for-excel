/**
 * instructions — update persistent user/workbook rules.
 *
 * The tool name remains "instructions" for backward compatibility with
 * existing sessions. All user-facing text says "rules".
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { getAppStorage } from "@earendil-works/pi-web-ui/dist/storage/app-storage.js";

import {
  applyRuleAction,
  getUserRules,
  getWorkbookRules,
  setUserRules,
  setWorkbookRules,
  USER_RULES_SOFT_LIMIT,
  WORKBOOK_RULES_SOFT_LIMIT,
} from "../rules/store.js";
import { getWorkbookContext } from "../workbook/context.js";
import { getErrorMessage } from "../utils/errors.js";

const schema = Type.Object({
  action: Type.Union([
    Type.Literal("append"),
    Type.Literal("replace"),
  ], {
    description: "append = add to existing rules, replace = rewrite the full rules text",
  }),
  level: Type.Union([
    Type.Literal("user"),
    Type.Literal("workbook"),
  ], {
    description: "Target rule scope.",
  }),
  content: Type.String({
    description:
      "Rule text to save. For append, this is the new line/note to add. For replace, this becomes the full text.",
  }),
});

type Params = Static<typeof schema>;

function getSoftLimit(level: Params["level"]): number {
  return level === "user" ? USER_RULES_SOFT_LIMIT : WORKBOOK_RULES_SOFT_LIMIT;
}

function emitRulesUpdatedEvent(): void {
  if (typeof document === "undefined") return;

  document.dispatchEvent(new CustomEvent("pi:rules-updated"));
  document.dispatchEvent(new CustomEvent("pi:status-update"));
}

export function createInstructionsTool(): AgentTool<typeof schema, undefined> {
  return {
    name: "instructions",
    label: "Rules",
    description:
      "Update persistent rules for the agent. " +
      "Use level=user for personal preferences (all files) and level=workbook for workbook-specific notes (this file).",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const storage = getAppStorage();
        const settings = storage.settings;

        if (params.action === "append" && params.content.trim().length === 0) {
          return {
            content: [{ type: "text", text: "Error: content is required for append." }],
            details: undefined,
          };
        }

        if (params.level === "user") {
          const current = await getUserRules(settings);
          const updated = applyRuleAction({
            currentValue: current,
            action: params.action,
            content: params.content,
          });

          const saved = await setUserRules(settings, updated);
          emitRulesUpdatedEvent();

          const body = saved ?? "(No user rules set.)";
          const warning =
            saved && saved.length > USER_RULES_SOFT_LIMIT
              ? `\n\n⚠️ User rules are above the ${USER_RULES_SOFT_LIMIT}-char soft limit.`
              : "";

          return {
            content: [
              {
                type: "text",
                text: `Updated user rules (${saved?.length ?? 0}/${USER_RULES_SOFT_LIMIT} chars):\n\n${body}${warning}`,
              },
            ],
            details: undefined,
          };
        }

        const workbookContext = await getWorkbookContext();
        const workbookId = workbookContext.workbookId;

        if (!workbookId) {
          return {
            content: [{
              type: "text",
              text: "Error: workbook identity unavailable. Can't update workbook rules right now.",
            }],
            details: undefined,
          };
        }

        const current = await getWorkbookRules(settings, workbookId);
        const updated = applyRuleAction({
          currentValue: current,
          action: params.action,
          content: params.content,
        });

        const saved = await setWorkbookRules(settings, workbookId, updated);
        emitRulesUpdatedEvent();

        const limit = getSoftLimit(params.level);
        const body = saved ?? "(No workbook rules set.)";
        const warning =
          saved && saved.length > limit
            ? `\n\n⚠️ Workbook rules are above the ${limit}-char soft limit.`
            : "";

        return {
          content: [
            {
              type: "text",
              text: `Updated workbook rules (${saved?.length ?? 0}/${limit} chars):\n\n${body}${warning}`,
            },
          ],
          details: undefined,
        };
      } catch (error: unknown) {
        return {
          content: [{ type: "text", text: `Error updating rules: ${getErrorMessage(error)}` }],
          details: undefined,
        };
      }
    },
  };
}
