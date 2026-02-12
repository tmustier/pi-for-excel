/**
 * view_settings — Control worksheet display settings.
 *
 * Scope: on-screen worksheet view/navigation only (not print/page layout).
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@mariozechner/pi-agent-core";
import { excelRun, qualifiedAddress } from "../excel/helpers.js";
import {
  getWorkbookChangeAuditLog,
  type AppendWorkbookChangeAuditEntryArgs,
} from "../audit/workbook-change-audit.js";
import { getErrorMessage } from "../utils/errors.js";
import {
  NON_CHECKPOINTED_MUTATION_NOTE,
  NON_CHECKPOINTED_MUTATION_REASON,
  recoveryCheckpointUnavailable,
} from "./recovery-metadata.js";
import type { ViewSettingsDetails } from "./tool-details.js";

function StringEnum<T extends string[]>(values: [...T], opts?: { description?: string }) {
  return Type.Union(
    values.map((v) => Type.Literal(v)),
    opts,
  );
}

const schema = Type.Object({
  action: StringEnum(
    [
      "get",
      "show_gridlines",
      "hide_gridlines",
      "show_headings",
      "hide_headings",
      "freeze_rows",
      "freeze_columns",
      "freeze_at",
      "unfreeze",
      "set_tab_color",
      "hide_sheet",
      "show_sheet",
      "very_hide_sheet",
      "set_standard_width",
      "activate",
    ],
    { description: "The view setting to read or change." },
  ),
  sheet: Type.Optional(
    Type.String({
      description:
        "Target sheet name. Defaults to the active sheet for most actions. " +
        "Required for hide/show/very_hide and activate.",
    }),
  ),
  count: Type.Optional(
    Type.Number({
      description: "Number of rows or columns to freeze. Required for freeze_rows/freeze_columns.",
    }),
  ),
  range: Type.Optional(
    Type.String({
      description:
        "Cell range for freeze_at (e.g. \"B3\"). Everything above and to the left of " +
        "this cell will be frozen.",
    }),
  ),
  color: Type.Optional(
    Type.String({
      description: "Tab color in #RRGGBB format (e.g. \"#FF6600\"). Use \"\" to clear.",
    }),
  ),
  width: Type.Optional(
    Type.Number({
      description:
        "Standard (default) column width for the worksheet, in Excel character-width units. " +
        "Required for set_standard_width.",
    }),
  ),
});

type Params = Static<typeof schema>;
type ViewSettingsAction = Params["action"];

interface ViewSettingsActionResult {
  text: string;
  outputAddress?: string;
  changedCount?: number;
  summary?: string;
}

interface ViewSettingsToolDependencies {
  executeAction: (params: Params) => Promise<ViewSettingsActionResult>;
  appendAuditEntry: (entry: AppendWorkbookChangeAuditEntryArgs) => Promise<void>;
}

function requireSheetName(action: string, sheet: string | undefined): string {
  if (!sheet) {
    throw new Error(`sheet is required for ${action}`);
  }
  return sheet;
}

function isMutatingViewSettingsAction(action: ViewSettingsAction): boolean {
  return action !== "get";
}

function buildMutationDetails(args: {
  action: ViewSettingsAction;
  address?: string;
}): ViewSettingsDetails {
  return {
    kind: "view_settings",
    action: args.action,
    address: args.address,
    recovery: recoveryCheckpointUnavailable(NON_CHECKPOINTED_MUTATION_REASON),
  };
}

const defaultDependencies: ViewSettingsToolDependencies = {
  executeAction: executeViewSettingsAction,
  appendAuditEntry: (entry) => getWorkbookChangeAuditLog().append(entry),
};

export function createViewSettingsTool(
  dependencies: Partial<ViewSettingsToolDependencies> = {},
): AgentTool<typeof schema> {
  const resolvedDependencies: ViewSettingsToolDependencies = {
    executeAction: dependencies.executeAction ?? defaultDependencies.executeAction,
    appendAuditEntry: dependencies.appendAuditEntry ?? defaultDependencies.appendAuditEntry,
  };

  return {
    name: "view_settings",
    label: "View Settings",
    description:
      "Read or change worksheet view/navigation settings: gridlines, row/column headings, " +
      "freeze panes, tab color, sheet visibility, sheet activation, and standard width. " +
      "Use \"get\" to inspect the current state first.",
    parameters: schema,
    execute: async (
      toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ViewSettingsDetails | undefined>> => {
      const isMutation = isMutatingViewSettingsAction(params.action);

      try {
        const result = await resolvedDependencies.executeAction(params);
        const outputAddress = result.outputAddress ?? params.range ?? params.sheet;

        if (isMutation) {
          await resolvedDependencies.appendAuditEntry({
            toolName: "view_settings",
            toolCallId,
            blocked: false,
            outputAddress,
            changedCount: result.changedCount ?? 1,
            changes: [],
            summary: result.summary ?? `${params.action} view setting`,
          });
        }

        const details = isMutation
          ? buildMutationDetails({
              action: params.action,
              address: outputAddress,
            })
          : undefined;

        const note = isMutation
          ? `${result.text}\n\n${NON_CHECKPOINTED_MUTATION_NOTE}`
          : result.text;

        return {
          content: [{ type: "text", text: note }],
          details,
        };
      } catch (e: unknown) {
        const message = getErrorMessage(e);
        const outputAddress = params.range ?? params.sheet;

        if (isMutation) {
          await resolvedDependencies.appendAuditEntry({
            toolName: "view_settings",
            toolCallId,
            blocked: true,
            outputAddress,
            changedCount: 0,
            changes: [],
            summary: `error: ${message}`,
          });
        }

        const details = isMutation
          ? buildMutationDetails({
              action: params.action,
              address: outputAddress,
            })
          : undefined;

        const errorText = isMutation
          ? `Error: ${message}\n\n${NON_CHECKPOINTED_MUTATION_NOTE}`
          : `Error: ${message}`;

        return {
          content: [{ type: "text", text: errorText }],
          details,
        };
      }
    },
  };
}

async function executeViewSettingsAction(params: Params): Promise<ViewSettingsActionResult> {
  return excelRun(async (context) => {
    const sheet = params.sheet
      ? context.workbook.worksheets.getItem(params.sheet)
      : context.workbook.worksheets.getActiveWorksheet();

    switch (params.action) {
      case "get": {
        sheet.load("name, showGridlines, showHeadings, tabColor, visibility, standardWidth");
        const frozen = sheet.freezePanes.getLocationOrNullObject();
        frozen.load("address");
        await context.sync();

        const lines: string[] = [
          `Sheet: "${sheet.name}"`,
          `Visibility: ${sheet.visibility}`,
          `Gridlines: ${sheet.showGridlines ? "visible" : "hidden"}`,
          `Headings: ${sheet.showHeadings ? "visible" : "hidden"}`,
          `Tab color: ${sheet.tabColor || "(none)"}`,
          `Standard width: ${sheet.standardWidth}`,
          `Frozen panes: ${frozen.isNullObject ? "none" : frozen.address}`,
        ];
        return { text: lines.join("\n") };
      }

      case "show_gridlines": {
        sheet.showGridlines = true;
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Gridlines visible on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `showed gridlines on ${sheet.name}`,
        };
      }

      case "hide_gridlines": {
        sheet.showGridlines = false;
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Gridlines hidden on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `hid gridlines on ${sheet.name}`,
        };
      }

      case "show_headings": {
        sheet.showHeadings = true;
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Headings visible on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `showed headings on ${sheet.name}`,
        };
      }

      case "hide_headings": {
        sheet.showHeadings = false;
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Headings hidden on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `hid headings on ${sheet.name}`,
        };
      }

      case "freeze_rows": {
        if (params.count === undefined) throw new Error("count is required for freeze_rows");
        sheet.freezePanes.freezeRows(params.count);
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Froze top ${params.count} row(s) on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `froze ${params.count} row(s) on ${sheet.name}`,
        };
      }

      case "freeze_columns": {
        if (params.count === undefined) throw new Error("count is required for freeze_columns");
        sheet.freezePanes.freezeColumns(params.count);
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Froze first ${params.count} column(s) on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `froze ${params.count} column(s) on ${sheet.name}`,
        };
      }

      case "freeze_at": {
        if (!params.range) throw new Error("range is required for freeze_at");
        const ref = params.sheet ? `${params.sheet}!${params.range}` : params.range;
        const freezeRange = params.sheet
          ? sheet.getRange(params.range)
          : context.workbook.worksheets.getActiveWorksheet().getRange(params.range);
        sheet.freezePanes.freezeAt(freezeRange);
        await context.sync();
        sheet.load("name");
        await context.sync();

        const fullAddr = qualifiedAddress(sheet.name, params.range);
        return {
          text: `Froze panes at ${ref} on "${sheet.name}".`,
          outputAddress: fullAddr,
          changedCount: 1,
          summary: `froze panes at ${fullAddr}`,
        };
      }

      case "unfreeze": {
        sheet.freezePanes.unfreeze();
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: `Unfroze all panes on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `unfroze panes on ${sheet.name}`,
        };
      }

      case "set_tab_color": {
        if (params.color === undefined) throw new Error("color is required for set_tab_color");
        sheet.tabColor = params.color;
        await context.sync();
        sheet.load("name");
        await context.sync();
        return {
          text: params.color
            ? `Set tab color to ${params.color} on "${sheet.name}".`
            : `Cleared tab color on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: params.color
            ? `set tab color ${params.color} on ${sheet.name}`
            : `cleared tab color on ${sheet.name}`,
        };
      }

      case "hide_sheet": {
        const targetName = requireSheetName("hide_sheet", params.sheet);
        const target = context.workbook.worksheets.getItem(targetName);
        target.visibility = "Hidden";
        await context.sync();
        return {
          text: `Set sheet "${targetName}" visibility to Hidden.`,
          outputAddress: targetName,
          changedCount: 1,
          summary: `hid sheet ${targetName}`,
        };
      }

      case "show_sheet": {
        const targetName = requireSheetName("show_sheet", params.sheet);
        const target = context.workbook.worksheets.getItem(targetName);
        target.visibility = "Visible";
        await context.sync();
        return {
          text: `Set sheet "${targetName}" visibility to Visible.`,
          outputAddress: targetName,
          changedCount: 1,
          summary: `showed sheet ${targetName}`,
        };
      }

      case "very_hide_sheet": {
        const targetName = requireSheetName("very_hide_sheet", params.sheet);
        const target = context.workbook.worksheets.getItem(targetName);
        target.visibility = "VeryHidden";
        await context.sync();
        return {
          text: `Set sheet "${targetName}" visibility to VeryHidden.`,
          outputAddress: targetName,
          changedCount: 1,
          summary: `set sheet ${targetName} to VeryHidden`,
        };
      }

      case "set_standard_width": {
        if (params.width === undefined) {
          throw new Error("width is required for set_standard_width");
        }
        sheet.standardWidth = params.width;
        await context.sync();
        sheet.load("name,standardWidth");
        await context.sync();
        return {
          text: `Set standard width to ${sheet.standardWidth} on "${sheet.name}".`,
          outputAddress: sheet.name,
          changedCount: 1,
          summary: `set standard width ${sheet.standardWidth} on ${sheet.name}`,
        };
      }

      case "activate": {
        const targetName = requireSheetName("activate", params.sheet);
        const target = context.workbook.worksheets.getItem(targetName);
        target.activate();
        await context.sync();
        return {
          text: `Activated sheet "${targetName}".`,
          outputAddress: targetName,
          changedCount: 1,
          summary: `activated sheet ${targetName}`,
        };
      }

      default:
        throw new Error(`Unknown action: ${String(params.action)}`);
    }
  });
}
