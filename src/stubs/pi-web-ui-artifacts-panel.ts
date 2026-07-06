/**
 * Stub module for `@earendil-works/pi-web-ui/dist/tools/artifacts/artifacts.js`.
 *
 * The real ArtifactsPanel pulls in document preview/rendering dependencies.
 * Pi for Excel does not use artifacts yet.
 */

type ArtifactsTool = {
  label: string;
  name: "artifacts";
  description: string;
  parameters: DynamicValue;
  execute: (...args: DynamicValue[]) => Promise<never>;
};

export class ArtifactsPanel extends HTMLElement {
  // API surface used by pi-web-ui's ChatPanel (if it ever gets used here).
  agent: DynamicValue;
  sandboxUrlProvider: DynamicValue;

  onArtifactsChange: (() => void) | undefined;
  onClose: (() => void) | undefined;
  onOpen: (() => void) | undefined;

  collapsed = false;
  overlay = false;

  artifacts: Map<string, DynamicValue> = new Map();

  tool: ArtifactsTool;

  constructor() {
    super();
    this.tool = {
      label: "Artifacts",
      name: "artifacts",
      description: "Artifacts are not available in this build.",
      parameters: {},
      execute: () => Promise.reject(new Error("Artifacts are not available in this build")),
    };
  }

  async reconstructFromMessages(_messages: DynamicValue): Promise<void> {
    // no-op
  }
}

if (!customElements.get("artifacts-panel")) {
  customElements.define("artifacts-panel", ArtifactsPanel);
}
