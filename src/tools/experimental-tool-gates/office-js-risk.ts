const BARE_RISK_IDENTIFIERS: readonly string[] = [
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "EventSource",
  "localStorage",
  "sessionStorage",
  "indexedDB",
  "caches",
  "window",
  "globalThis",
  "self",
  "top",
  "parent",
  "frames",
  "document",
  "navigator",
  "location",
  "open",
  "postMessage",
  "Office",
  "Worker",
  "SharedWorker",
  "importScripts",
  "eval",
  "Function",
  "import",
];

const MEMBER_RISK_IDENTIFIERS: readonly string[] = [
  "constructor",
];

interface CompiledRiskPattern {
  identifier: string;
  pattern: RegExp;
}

function compileBarePattern(identifier: string): RegExp {
  return new RegExp(`(?<![.\\w$])${identifier}(?![\\w$])`, "u");
}

function compileMemberPattern(identifier: string): RegExp {
  return new RegExp(`(?<![\\w$])${identifier}(?![\\w$])`, "u");
}

const COMPILED_RISK_PATTERNS: readonly CompiledRiskPattern[] = [
  ...BARE_RISK_IDENTIFIERS.map((identifier) => ({
    identifier,
    pattern: compileBarePattern(identifier),
  })),
  ...MEMBER_RISK_IDENTIFIERS.map((identifier) => ({
    identifier,
    pattern: compileMemberPattern(identifier),
  })),
];

const UNICODE_ESCAPE_PATTERN = /\\u\{([0-9a-fA-F]{1,6})\}|\\u([0-9a-fA-F]{4})/gu;

const MAX_ESCAPE_DECODE_PASSES = 5;

function decodeUnicodeEscapesOnce(code: string): string {
  return code.replace(
    UNICODE_ESCAPE_PATTERN,
    (match: string, braced: string | undefined, plain: string | undefined) => {
      const hex = braced ?? plain;
      if (hex === undefined) {
        return match;
      }

      const codePoint = Number.parseInt(hex, 16);
      if (!Number.isSafeInteger(codePoint) || codePoint > 0x10ffff) {
        return match;
      }

      return String.fromCodePoint(codePoint);
    },
  );
}

function enumerateScanTargets(code: string): string[] {
  const targets = [code];

  let current = code;
  for (let pass = 0; pass < MAX_ESCAPE_DECODE_PASSES; pass += 1) {
    const decoded = decodeUnicodeEscapesOnce(current);
    if (decoded === current) {
      break;
    }

    targets.push(decoded);
    current = decoded;
  }

  return targets;
}

export interface OfficeJsCodeRiskAssessment {
  flagged: boolean;
  identifiers: string[];
  destructive: boolean;
}

export function assessOfficeJsCodeRisk(code: string): OfficeJsCodeRiskAssessment {
  const scanTargets = enumerateScanTargets(code);
  const identifiers: string[] = [];

  for (const { identifier, pattern } of COMPILED_RISK_PATTERNS) {
    if (scanTargets.some((target) => pattern.test(target))) {
      identifiers.push(identifier);
    }
  }

  return {
    flagged: identifiers.length > 0,
    identifiers,
    destructive: scanTargets.some((target) => /\b(?:delete|clear|deleteAll|remove|removeAll|Delete|Clear|DeleteAll)\b/u.test(target)),
  };
}
