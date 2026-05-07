/**
 * Stub module for `@earendil-works/pi-web-ui/dist/tools/artifacts/artifacts-tool-renderer.js`.
 *
 * The Excel add-in does not render artifacts.
 */

type ToolRenderResult = {
  content: unknown;
  isCustom: boolean;
};

export class ArtifactsToolRenderer {
  constructor(_panel: unknown) {}

  render(_params: unknown, _result: unknown, _isStreaming: boolean): ToolRenderResult {
    return { content: null, isCustom: false };
  }
}
