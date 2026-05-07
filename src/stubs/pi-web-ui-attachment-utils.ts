/**
 * Stub module for `@earendil-works/pi-web-ui/dist/utils/attachment-utils.js`.
 *
 * The real module bundles heavy document-parsing dependencies.
 * Pi for Excel currently does not support attachments in the sidebar UI.
 */

export function loadAttachment(_source: unknown, _fileName?: string): Promise<never> {
  return Promise.reject(new Error("Attachments are not supported in this build"));
}
