import { defineRule } from "@oxlint/plugins";
import type { ESTree } from "@oxlint/plugins";

function runsAfterModuleEvaluation(node: ESTree.Node): boolean {
  let current: ESTree.Node | null = node.parent;
  while (current !== null && current.type !== "Program") {
    if (
      current.type === "ArrowFunctionExpression"
      || current.type === "FunctionDeclaration"
      || current.type === "FunctionExpression"
      || (current.type === "PropertyDefinition" && !current.static)
    ) {
      return true;
    }
    current = current.parent;
  }
  return false;
}

export const noModuleScopeTranslationRule = defineRule({
  meta: {
    type: "problem",
    docs: { description: "Disallow t() calls that run while the module is evaluated, before the language is initialised." },
    messages: {
      moduleScope: "t() runs at module scope, before initLanguage() sets the language, so this string freezes to English; resolve it lazily inside a function.",
    },
  },
  createOnce(context) {
    return {
      CallExpression(node) {
        if (node.callee.type !== "Identifier" || node.callee.name !== "t") return;
        if (runsAfterModuleEvaluation(node)) return;
        context.report({ node, messageId: "moduleScope" });
      },
    };
  },
});
