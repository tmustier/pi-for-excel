import type { BorderWeight } from "../conventions/index.js";

export type BorderEdgeIndex =
  | "EdgeTop"
  | "EdgeBottom"
  | "EdgeLeft"
  | "EdgeRight"
  | "InsideHorizontal"
  | "InsideVertical";

export interface NormalizedBorderParams {
  shorthand?: BorderWeight;
  top?: BorderWeight;
  bottom?: BorderWeight;
  left?: BorderWeight;
  right?: BorderWeight;
}

export interface BorderInstructions {
  operations: Array<{ edge: BorderEdgeIndex; weight: BorderWeight }>;
  appliedText: string;
}

const BORDER_WEIGHT_ALIASES = new Map<string, BorderWeight>([
  ["thin", "thin"],
  ["medium", "medium"],
  ["thick", "thick"],
  ["none", "none"],
  ["borderthin", "thin"],
  ["bordersthin", "thin"],
  ["bordermedium", "medium"],
  ["bordersmedium", "medium"],
  ["borderthick", "thick"],
  ["bordersthick", "thick"],
  ["bordernone", "none"],
  ["bordersnone", "none"],
]);

const ALL_BORDER_EDGES: BorderEdgeIndex[] = [
  "EdgeTop",
  "EdgeBottom",
  "EdgeLeft",
  "EdgeRight",
  "InsideHorizontal",
  "InsideVertical",
];

function normalizeBorderWeight(value: DynamicValue, fieldName: string): BorderWeight | undefined {
  if (value === undefined) {
    return undefined;
  }

  if (typeof value !== "string") {
    throw new Error(`Invalid ${fieldName} value. Use thin, medium, thick, or none.`);
  }

  const normalized = value.trim().toLowerCase().replace(/[\s_-]/gu, "");
  const mapped = BORDER_WEIGHT_ALIASES.get(normalized);
  if (mapped !== undefined) {
    return mapped;
  }

  throw new Error(`Invalid ${fieldName} "${value}". Use thin, medium, thick, or none.`);
}

export function normalizeBorderParams(params: {
  borders?: DynamicValue;
  border_top?: DynamicValue;
  border_bottom?: DynamicValue;
  border_left?: DynamicValue;
  border_right?: DynamicValue;
}): NormalizedBorderParams {
  const shorthand = normalizeBorderWeight(params.borders, "borders");
  const top = normalizeBorderWeight(params.border_top, "border_top");
  const bottom = normalizeBorderWeight(params.border_bottom, "border_bottom");
  const left = normalizeBorderWeight(params.border_left, "border_left");
  const right = normalizeBorderWeight(params.border_right, "border_right");
  return {
    ...(shorthand !== undefined ? { shorthand } : {}),
    ...(top !== undefined ? { top } : {}),
    ...(bottom !== undefined ? { bottom } : {}),
    ...(left !== undefined ? { left } : {}),
    ...(right !== undefined ? { right } : {}),
  };
}

export function buildBorderInstructions(
  borderParams: NormalizedBorderParams,
  props: { borderTop?: BorderWeight; borderBottom?: BorderWeight; borderLeft?: BorderWeight; borderRight?: BorderWeight },
  color?: string,
): BorderInstructions | null {
  const shorthand = borderParams.shorthand;
  const hasShorthand = shorthand !== undefined;
  const hasEdges = borderParams.top !== undefined || borderParams.bottom !== undefined ||
    borderParams.left !== undefined || borderParams.right !== undefined;
  const hasStyleEdges = props.borderTop !== undefined || props.borderBottom !== undefined ||
    props.borderLeft !== undefined || props.borderRight !== undefined;

  if (!hasShorthand && !hasEdges && !hasStyleEdges) {
    return null;
  }

  if (hasShorthand && !hasEdges && !hasStyleEdges) {
    return {
      operations: ALL_BORDER_EDGES.map((edge) => ({ edge, weight: shorthand })),
      appliedText: `${shorthand} borders${color ? ` (${color})` : ""}`,
    };
  }

  const edges: Array<{ edge: "EdgeTop" | "EdgeBottom" | "EdgeLeft" | "EdgeRight"; param: BorderWeight | undefined; styleProp: BorderWeight | undefined; label: string }> = [
    { edge: "EdgeTop", param: borderParams.top, styleProp: props.borderTop, label: "top" },
    { edge: "EdgeBottom", param: borderParams.bottom, styleProp: props.borderBottom, label: "bottom" },
    { edge: "EdgeLeft", param: borderParams.left, styleProp: props.borderLeft, label: "left" },
    { edge: "EdgeRight", param: borderParams.right, styleProp: props.borderRight, label: "right" },
  ];

  const operations: BorderInstructions["operations"] = [];
  const appliedEdges: string[] = [];

  for (const { edge, param, styleProp, label } of edges) {
    const weight = param ?? styleProp;
    if (weight !== undefined) {
      operations.push({ edge, weight });
      appliedEdges.push(`${label}:${weight}`);
    }
  }

  if (operations.length === 0) {
    return null;
  }

  return {
    operations,
    appliedText: `borders ${appliedEdges.join(", ")}${color ? ` (${color})` : ""}`,
  };
}
