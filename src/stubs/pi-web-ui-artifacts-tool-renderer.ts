/**
 * Stub module for `@earendil-works/pi-web-ui/dist/tools/artifacts/artifacts-tool-renderer.js`.
 *
 * The Excel add-in does not render artifacts.
 */

type ToolRenderResult = {
  content: DynamicValue;
  isCustom: boolean;
};

export class ArtifactsToolRenderer {
  constructor(_panel: DynamicValue) {}

  render(_params: DynamicValue, _result: DynamicValue, _isStreaming: boolean): ToolRenderResult {
    return { content: null, isCustom: false };
  }
}
