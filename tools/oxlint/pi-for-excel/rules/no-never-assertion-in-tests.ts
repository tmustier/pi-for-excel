import { defineRule, type ESTree } from "@oxlint/plugins";

export const noNeverAssertionInTestsRule = defineRule({
  meta: {
    type: "problem",
    docs: { description: "Reject assertions to never in tests." },
    messages: {
      never: "Do not manufacture an impossible state with a never assertion. Test invalid input at the public parsing boundary or use a faithful typed fake.",
    },
  },
  createOnce(context) {
    const check = (node: ESTree.TSAsExpression | ESTree.TSTypeAssertion) => {
      if (node.typeAnnotation.type !== "TSNeverKeyword") return;
      if (!/(?:\.(?:test|spec)\.[cm]?tsx?$|\.browser-test\.ts$|[/\\]__tests__[/\\])/.test(context.filename)) return;
      context.report({ node, messageId: "never" });
    };
    return { TSAsExpression: check, TSTypeAssertion: check };
  },
});
