/**
 * Stub module for `@earendil-works/pi-web-ui/dist/tools/javascript-repl.js`.
 *
 * The Excel add-in does not expose the JavaScript REPL tool.
 */

type JavaScriptReplTool = {
  label: string;
  name: "javascript_repl";
  description: string;
  parameters: unknown;
  execute: (...args: unknown[]) => Promise<never>;
};

export function createJavaScriptReplTool(): JavaScriptReplTool {
  return {
    label: "JavaScript REPL",
    name: "javascript_repl",
    description: "javascript_repl is not available in this build.",
    parameters: {},
    execute: () => Promise.reject(new Error("javascript_repl is not available in this build")),
  };
}

export const javascriptReplTool: JavaScriptReplTool = createJavaScriptReplTool();
