/**
 * Static ambient-authority risk lint for `execute_office_js` code.
 *
 * Submitted code is compiled to a blob module and imported into the taskpane
 * page realm, so it has ambient access to browser globals — network, storage
 * (where provider credentials live), the DOM, and realm-escape primitives.
 * Auto mode skips the per-call approval prompt for pure Excel API code; this
 * lint keeps the prompt — in any execution mode — for code that references
 * ambient authority.
 *
 * Design notes:
 * - This is a heuristic tripwire, not a sandbox. It deliberately scans the
 *   raw code text (including strings and comments), which also catches simple
 *   obfuscation staged through `eval` / `Function` string payloads. False
 *   positives cost one approval dialog, never a hard block.
 * - Identifiers are matched as bare references (not preceded by `.`), so
 *   harmless member access like `chart.top` or `range.format` never trips the
 *   lint. `constructor` is the exception: the dangerous form *is* the member
 *   access (`({}).constructor.constructor(...)`), so it is flagged anywhere.
 * - JavaScript accepts Unicode escapes inside identifiers (`\u0066etch`
 *   executes as `fetch`), so the scan also runs against a version of the code
 *   with `\uXXXX` / `\u{...}` escapes decoded to a fixpoint. String-literal
 *   escapes like `"caf\u00e9"` decode to harmless text and do not trip it.
 * - The durable fix is an isolated-realm runner with only a guarded workbook
 *   API bridged in (#605); until then this lint plus the CSP allowlist is the
 *   boundary between workbook-content prompt injection and the page realm.
 */

/** Bare-reference identifiers that reach ambient browser authority. */
const BARE_RISK_IDENTIFIERS: readonly string[] = [
  // Network egress
  "fetch",
  "XMLHttpRequest",
  "WebSocket",
  "EventSource",
  // Storage (settings, credentials, caches)
  "localStorage",
  "sessionStorage",
  "indexedDB",
  "caches",
  // Global / DOM handles (paths back to everything above)
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
  // Office host surface beyond the provided Excel context
  "Office",
  // Workers and script loading
  "Worker",
  "SharedWorker",
  "importScripts",
  // Dynamic code evaluation
  "eval",
  "Function",
  "import",
];

/** Identifiers flagged even as member access (realm-escape primitives). */
const MEMBER_RISK_IDENTIFIERS: readonly string[] = [
  "constructor",
];

interface CompiledRiskPattern {
  identifier: string;
  pattern: RegExp;
}

function compileBarePattern(identifier: string): RegExp {
  // Not preceded by `.`, identifier chars, or `$` (excludes member access and
  // longer identifiers); not followed by identifier chars.
  return new RegExp(`(?<![.\\w$])${identifier}(?![\\w$])`, "u");
}

function compileMemberPattern(identifier: string): RegExp {
  // Preceding `.` allowed: member access is the dangerous form.
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

/** Max decode passes when chasing nested escape encodings to a fixpoint. */
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

/**
 * Produce every distinct decoding of the code, chasing `\u` escapes to a
 * fixpoint (bounded) so escape-encoded identifiers cannot slip past the scan.
 */
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
  /** True when the code references ambient authority beyond the Excel API. */
  flagged: boolean;
  /** Catalog-ordered unique identifiers that triggered the assessment. */
  identifiers: string[];
}

/**
 * Assess `execute_office_js` code for references to ambient browser
 * authority beyond the provided `context: Excel.RequestContext`.
 */
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
  };
}
