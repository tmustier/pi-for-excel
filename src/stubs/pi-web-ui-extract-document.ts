/**
 * Stub module for `@earendil-works/pi-web-ui/dist/tools/extract-document.js`.
 *
 * pi-web-ui's public entrypoint re-exports this tool, and the real
 * implementation pulls in large dependencies via `loadAttachment()`:
 * - pdfjs-dist (+ worker)
 * - docx-preview
 * - xlsx
 *
 * The Excel add-in does not ship the extract_document tool.
 */

type ExtractDocumentTool = {
  label: string;
  name: "extract_document";
  description: string;
  parameters: unknown;
  execute: (...args: unknown[]) => Promise<never>;
};

export function createExtractDocumentTool(): ExtractDocumentTool {
  return {
    label: "Extract Document",
    name: "extract_document",
    description: "extract_document is not available in this build.",
    parameters: {},
    execute: () => Promise.reject(new Error("extract_document is not available in this build")),
  };
}

export const extractDocumentTool: ExtractDocumentTool = createExtractDocumentTool();
