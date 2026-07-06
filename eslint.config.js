import tseslint from "typescript-eslint";

const NO_UNKNOWN_MESSAGE =
  "Do not introduce `unknown` here. If you think you need `unknown`, fix the upstream type first: the agent should not have an unknown at this point in the code at all.";

const NO_GENERIC_OBJECT_GUARD_MESSAGE =
  "Do not define or use generic object/record guards such as `isRecord`, `isObjectValue`, or `isPlainObject`. This hides an upstream `unknown`; fix the real typed boundary first because the agent should not have an unknown at this point in the code at all.";

const GENERIC_OBJECT_GUARD_NAME_PATTERN =
  "/^(?:is.*[Rr]ecord.*|is.*Object(?:Value|Map|Like|Payload)?|isObject(?:Value|Map|Like)?|isPlainObject)$/";

const BAN_GENERIC_OBJECT_GUARD_SELECTORS = [
  `FunctionDeclaration[id.name=${GENERIC_OBJECT_GUARD_NAME_PATTERN}]`,
  `VariableDeclarator[id.name=${GENERIC_OBJECT_GUARD_NAME_PATTERN}]`,
  `ImportSpecifier[imported.name=${GENERIC_OBJECT_GUARD_NAME_PATTERN}]`,
  `CallExpression[callee.name=${GENERIC_OBJECT_GUARD_NAME_PATTERN}]`,
].map((selector) => ({ selector, message: NO_GENERIC_OBJECT_GUARD_MESSAGE }));

export default tseslint.config(
  {
    ignores: [
      "dist/**",
      "node_modules/**",
      "poc/**",
      ".research/**",
      "research/**",
    ],
  },

  // Baseline: recommended + type-checked rules for all TS files.
  ...tseslint.configs.recommendedTypeChecked,

  {
    languageOptions: {
      parserOptions: {
        project: "./tsconfig.eslint.json",
        tsconfigRootDir: import.meta.dirname,
      },
    },

    rules: {
      // ── Type-system hygiene (Python-typing spirit) ─────────────────────────

      // Ban ts-ignore (force fixing the real type issue). Allow ts-expect-error
      // but require an explanation.
      "@typescript-eslint/ban-ts-comment": [
        "error",
        {
          "ts-ignore": true,
          "ts-nocheck": true,
          "ts-expect-error": "allow-with-description",
          minimumDescriptionLength: 10,
        },
      ],

      // Any defeats type checking. Warn for now; tighten to "error" once clean.
      "@typescript-eslint/no-explicit-any": "warn",

      // Non-null assertion is a common escape hatch; prefer runtime checks.
      "@typescript-eslint/no-non-null-assertion": "warn",

      // Type assertions should be rare; prefer narrowing/guards.
      "@typescript-eslint/consistent-type-assertions": [
        "warn",
        {
          assertionStyle: "as",
          objectLiteralTypeAssertions: "never",
        },
      ],

      // Dynamic values must be normalized at typed boundaries rather than
      // leaking `unknown` and generic record probes through application code.
      "no-restricted-syntax": [
        "error",
        { selector: "TSUnknownKeyword", message: NO_UNKNOWN_MESSAGE },
        ...BAN_GENERIC_OBJECT_GUARD_SELECTORS,
      ],

      // ── Async safety ───────────────────────────────────────────────────────
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      "@typescript-eslint/await-thenable": "error",

      // ── Dead-code hygiene ──────────────────────────────────────────────────
      "@typescript-eslint/no-unused-vars": [
        "warn",
        {
          argsIgnorePattern: "^_",
          varsIgnorePattern: "^_",
          caughtErrorsIgnorePattern: "^_",
        },
      ],

      // ── Relaxations for legitimate patterns ────────────────────────────────

      // We use `require()` in a few CJS-compat spots (install-githooks, etc.).
      "@typescript-eslint/no-require-imports": "off",

      // Empty catch blocks are fine when intentional (e.g. best-effort cleanup).
      "@typescript-eslint/no-empty-function": "off",

      // `void` operator is used to deliberately discard promises in fire-and-forget.
      "no-void": "off",

      // Allow `${expr}` in template literals even when expr is non-string.
      "@typescript-eslint/restrict-template-expressions": [
        "warn",
        {
          allowNumber: true,
          allowBoolean: true,
          allowNullish: false,
        },
      ],

      // Unbound methods show up in event-handler patterns with Lit;
      // too noisy to be useful here.
      "@typescript-eslint/unbound-method": "off",
    },
  },

  {
    files: ["src/types/dynamic-values.d.ts"],
    rules: {
      // The single sanctioned untyped boundary marker. All other explicit
      // `unknown` spellings remain banned by the main rule above.
      "no-restricted-syntax": [
        "error",
        ...BAN_GENERIC_OBJECT_GUARD_SELECTORS,
      ],
    },
  },
);
