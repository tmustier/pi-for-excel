/**
 * explain_formula — provide a concise natural-language explanation for a formula cell.
 */

import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { cellAddress, excelRun, getRange, parseCell, parseRangeRef, qualifiedAddress } from "../excel/helpers.js";
import { getErrorMessage } from "../utils/errors.js";
import type { ExplainFormulaDetails, ExplainFormulaReferenceDetail } from "./tool-details.js";
import {
  buildExplainFormulaNarrative,
  extractFormulaFunctionNames,
  previewCellValue,
} from "./explain-formula-logic.js";
import { extractFormulaReferences, type ParsedFormulaReference } from "./trace-dependencies-logic.js";

const DEFAULT_MAX_REFERENCES = 8;
const MAX_REFERENCES_LIMIT = 20;

const schema = Type.Object({
  cell: Type.String({
    description: 'Single formula cell to explain, e.g. "D10" or "Sheet2!F5".',
  }),
  max_references: Type.Optional(
    Type.Number({
      description: "Max number of direct references to preview. Default: 8. Max: 20.",
    }),
  ),
});

type Params = Static<typeof schema>;

function toQualifiedReferenceAddress(reference: ParsedFormulaReference): string {
  const start = cellAddress(reference.startCol, reference.startRow);
  const end = cellAddress(reference.endCol, reference.endRow);
  const localAddress = start === end ? start : `${start}:${end}`;
  return qualifiedAddress(reference.sheet, localAddress);
}

function clampMaxReferences(value: number | undefined): number {
  if (typeof value !== "number" || Number.isNaN(value)) {
    return DEFAULT_MAX_REFERENCES;
  }

  const rounded = Math.floor(value);
  if (rounded < 1) return 1;
  if (rounded > MAX_REFERENCES_LIMIT) return MAX_REFERENCES_LIMIT;
  return rounded;
}

export function isSingleCellReference(reference: string): boolean {
  try {
    const parsed = parseRangeRef(reference);
    const localAddress = parsed.address.trim();

    if (localAddress.includes(":") || localAddress.includes(",")) {
      return false;
    }

    parseCell(localAddress);
    return true;
  } catch {
    return false;
  }
}

export function createExplainFormulaTool(): AgentTool<typeof schema, ExplainFormulaDetails> {
  return {
    name: "explain_formula",
    label: "Explain Formula",
    description:
      "Explain what a formula cell is doing in plain language, including direct input references and current values.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<ExplainFormulaDetails>> => {
      try {
        if (!isSingleCellReference(params.cell)) {
          return {
            content: [{ type: "text", text: "Error: explain_formula expects a single cell, not a range." }],
            details: {
              kind: "explain_formula",
              cell: params.cell,
              hasFormula: false,
              explanation: "The request was rejected because the input is not a single cell.",
              references: [],
              truncated: false,
            },
          };
        }

        const maxReferences = clampMaxReferences(params.max_references);

        const details = await excelRun(async (context) => {
          const { sheet, range } = getRange(context, params.cell);
          sheet.load("name");
          range.load("address,values,formulas");
          await context.sync();

          const resolvedCell = qualifiedAddress(sheet.name, range.address);
          const valuePreview = previewCellValue(range.values[0][0]);
          const rawFormula: unknown = range.formulas[0][0];
          const formula = typeof rawFormula === "string" && rawFormula.startsWith("=")
            ? rawFormula
            : undefined;

          if (!formula) {
            return {
              kind: "explain_formula",
              cell: resolvedCell,
              hasFormula: false,
              valuePreview,
              explanation: "This cell currently contains a static value, not a formula.",
              references: [],
              truncated: false,
            } satisfies ExplainFormulaDetails;
          }

          const parsedReferences = extractFormulaReferences(formula, sheet.name);
          const truncated = parsedReferences.length > maxReferences;
          const referencesToLoad = parsedReferences.slice(0, maxReferences);

          const loadedReferences: Array<{
            address: string;
            range: Excel.Range;
          }> = [];

          for (const reference of referencesToLoad) {
            const address = toQualifiedReferenceAddress(reference);
            const { range: refRange } = getRange(context, reference.startAddress);
            refRange.load("values,formulas");
            loadedReferences.push({ address, range: refRange });
          }

          if (loadedReferences.length > 0) {
            await context.sync();
          }

          const referenceDetails: ExplainFormulaReferenceDetail[] = loadedReferences.map((reference) => {
            const preview = previewCellValue(reference.range.values[0][0]);
            const rawRefFormula: unknown = reference.range.formulas[0][0];

            return {
              address: reference.address,
              valuePreview: preview,
              formulaPreview: typeof rawRefFormula === "string" && rawRefFormula.startsWith("=")
                ? rawRefFormula
                : undefined,
            };
          });

          const functionNames = extractFormulaFunctionNames(formula);
          const explanation = buildExplainFormulaNarrative({
            valuePreview,
            functionNames,
            referenceCount: parsedReferences.length,
            truncated,
          });

          return {
            kind: "explain_formula",
            cell: resolvedCell,
            hasFormula: true,
            formula,
            valuePreview,
            explanation,
            references: referenceDetails,
            truncated,
          } satisfies ExplainFormulaDetails;
        });

        if (!details.hasFormula) {
          return {
            content: [{
              type: "text",
              text: `**Formula explanation for ${details.cell}**\n\n${details.explanation}`,
            }],
            details,
          };
        }

        const referenceLines = details.references.length > 0
          ? details.references.map((reference) => {
            const preview = reference.valuePreview ? ` → ${reference.valuePreview}` : "";
            const formulaPreview = reference.formulaPreview ? ` (formula: \`${reference.formulaPreview}\`)` : "";
            return `- ${reference.address}${preview}${formulaPreview}`;
          })
          : ["- (No direct references detected)"];

        const lines = [
          `**Formula explanation for ${details.cell}**`,
          "",
          `- Current value: ${details.valuePreview ?? "(blank)"}`,
          `- Formula: \`${details.formula}\``,
          "",
          details.explanation,
          "",
          `Direct references (${details.references.length} shown):`,
          ...referenceLines,
        ];

        if (details.truncated) {
          lines.push("", `_Showing first ${details.references.length} reference(s)._`);
        }

        return {
          content: [{ type: "text", text: lines.join("\n") }],
          details,
        };
      } catch (error: unknown) {
        return {
          content: [{ type: "text", text: `Error explaining formula: ${getErrorMessage(error)}` }],
          details: {
            kind: "explain_formula",
            cell: params.cell,
            hasFormula: false,
            explanation: `Failed to explain formula: ${getErrorMessage(error)}`,
            references: [],
            truncated: false,
          },
        };
      }
    },
  };
}
