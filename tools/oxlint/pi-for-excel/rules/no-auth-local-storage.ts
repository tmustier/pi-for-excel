import { defineRule } from "@oxlint/plugins";

export const noAuthLocalStorageRule = defineRule({
  meta: {
    type: "problem",
    docs: { description: "Disallow any reference to localStorage in auth code." },
    messages: {
      localStorage: "OAuth credentials must not use localStorage; use the encrypted OAuth storage boundary.",
    },
  },
  createOnce(context) {
    return {
      Identifier(node) {
        if (node.name === "localStorage") context.report({ node, messageId: "localStorage" });
      },
    };
  },
});
