import { defineRule } from "@oxlint/plugins";
import type { ESTree } from "@oxlint/plugins";

const allowedBoundaryTypes = new Set(["unknown", "JsonValue"]);

function unwrap(node: ESTree.Node): ESTree.Node {
  let current = node;
  while (current.type === "ParenthesizedExpression" || current.type === "ChainExpression") {
    current = current.expression;
  }
  return current;
}

function calledMethod(node: ESTree.Node): ESTree.MemberExpression | null {
  const call = unwrap(node);
  if (call.type !== "CallExpression") return null;
  const callee = unwrap(call.callee);
  if (callee.type !== "MemberExpression" || callee.computed || callee.property.type !== "Identifier") return null;
  return callee;
}

function boundaryKind(node: ESTree.Node, sourceText: (node: ESTree.Node) => string): string | null {
  const expression = unwrap(node);
  const jsonParse = calledMethod(expression);
  if (jsonParse?.property.type === "Identifier" && jsonParse.property.name === "parse" && sourceText(jsonParse.object) === "JSON") {
    return "JSON.parse";
  }
  const call = expression.type === "AwaitExpression" ? expression.argument : expression;
  const responseJson = calledMethod(call);
  if (responseJson?.property.type === "Identifier" && responseJson.property.name === "json") return "Response.json";
  return null;
}

/** Decoded JSON and fetch payloads land as unknown, then a concrete parser refines them. */
export const noTypedJsonBoundaryCastRule = defineRule({
  meta: {
    type: "problem",
    docs: { description: "Disallow casting JSON.parse or Response.json results to anything but unknown or JsonValue." },
    messages: {
      typedCast: "{{kind}} cast to {{castType}}: decoded JSON/fetch payloads must first land as unknown, then be parsed/refined by a concrete boundary parser.",
    },
  },
  createOnce(context) {
    return {
      TSAsExpression(node) {
        const castType = context.sourceCode.getText(node.typeAnnotation).replace(/\s+/g, " ").trim();
        if (allowedBoundaryTypes.has(castType)) return;
        const kind = boundaryKind(node.expression, (child) => context.sourceCode.getText(child));
        if (kind !== null) context.report({ node, messageId: "typedCast", data: { kind, castType } });
      },
    };
  },
});
