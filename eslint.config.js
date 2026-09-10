import tseslint from "typescript-eslint";


export default tseslint.config(
  {
    linterOptions: {
      reportUnusedDisableDirectives: "error",
    },
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

      // Any defeats type checking.
      "@typescript-eslint/no-explicit-any": "error",

      // Non-null assertion is a common escape hatch; prefer runtime checks.
      "@typescript-eslint/no-non-null-assertion": "error",

      // Type assertions should be rare; prefer narrowing/guards.
      "@typescript-eslint/consistent-type-assertions": [
        "error",
        {
          assertionStyle: "as",
          objectLiteralTypeAssertions: "never",
        },
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
    files: ["src/auth/**/*.ts"],
    rules: {
      "no-restricted-syntax": [
        "error",
        {
          selector: "Identifier[name='localStorage']",
          message: "OAuth credentials must not use localStorage; use the encrypted OAuth storage boundary.",
        },
      ],
    },
  },
);
