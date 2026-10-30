module.exports = {
  root: true,
  // Scratch/dev-artifact trees are NOT production workspaces: `eslint .`
  // traversed 174k+ files (scratch/ alone: 275k files, 10.8k lint-targets)
  // and never finished (timed out at 90 s). Keep lint on committed
  // workspaces only: packages, scripts,
  // tests. (node_modules and .git
  // are ignored by default; every dir below is also gitignored except
  // .kilo, which holds nested worktree checkouts.)
  ignorePatterns: [
    'scratch/',
    '.astra/',
    'dist/',
    'release/',
    'coverage/',
    'test-results/',
    'plans/',
    'documents/',
    'pi-source/',
    '.specstory/',
    '.kilo/',
    // The website is its own repository (nested, gitignored); it is linted there.
    'website-project/',
  ],
  env: {
    browser: true,
    es2021: true,
    node: true,
  },
  extends: [
    'eslint:recommended',
    'plugin:@typescript-eslint/recommended',
    'prettier',
  ],
  parser: '@typescript-eslint/parser',
  parserOptions: {
    ecmaFeatures: {
      jsx: false,
    },
    ecmaVersion: 12,
    sourceType: 'module',
  },
  plugins: ['@typescript-eslint'],
  rules: {
    // no-explicit-any in warn mode (large surface: 500+ hits; error mode
    // would overwhelm. Warn first, ratchet to error as clusters land).
    '@typescript-eslint/no-explicit-any': 'warn',
  },
  overrides: [
    {
      // Ban raw adm-zip in production. The wrapper is fail-open, so
      // any direct import bypasses the CVE-2026-76845 instance hardening.
      // NOTE: keep in parity with scripts/check-adm-zip-imports.mjs SCAN_EXTS
      // (.ts/.tsx/.js/.jsx/.mjs/.cjs) and the grep-in-test ext list — any prod
      // ext the gates scan must also be linted here or raw imports slip past.
      files: [
        'packages/engine/src/**/*.{js,jsx,ts,tsx,mjs,cjs}',
        'packages/cli/src/**/*.{js,jsx,ts,tsx,mjs,cjs}',
        'packages/mcp-server/src/**/*.{js,jsx,ts,tsx,mjs,cjs}',
      ],
      rules: {
        'no-restricted-imports': [
          'error',
          {
            paths: [
              {
                name: 'adm-zip',
                message:
                  "Raw 'adm-zip' is banned (CVE-2026-76845, fail-open): import { AdmZip } from '@attentionform/astra-engine' (or the local './adm-zip-safe' wrapper) instead.",
              },
            ],
            patterns: [
              {
                group: ['adm-zip/*'],
                message:
                  "Raw 'adm-zip/*' subpaths are banned (CVE-2026-76845, fail-open): use the '@attentionform/astra-engine' AdmZip / './adm-zip-safe' wrapper instead.",
              },
            ],
          },
        ],
        // no-restricted-imports misses dynamic import() of adm-zip: ban it here.
        // TemplateLiteral forms (dynamic import of `adm-zip` / require of
        // `adm-zip`) bypass the Literal-only check above, so ban the
        // adm-zip-prefixed first quasi explicitly (prefix regex has no slash:
        // esquery-safe).
        'no-restricted-syntax': [
          'error',
          {
            selector: 'ImportExpression[source.value=/^adm-zip/]',
            message:
              "Dynamic import() of adm-zip is banned (CVE-2026-76845, fail-open): import { AdmZip } from '@attentionform/astra-engine' (or the local './adm-zip-safe' wrapper) instead.",
          },
          {
            selector: 'ImportExpression[source.quasis.0.value.raw=/^adm-zip/]',
            message:
              "Dynamic import() of adm-zip is banned (CVE-2026-76845, fail-open): import { AdmZip } from '@attentionform/astra-engine' (or the local './adm-zip-safe' wrapper) instead.",
          },
          {
            selector:
              'CallExpression[callee.name="require"][arguments.0.quasis.0.value.raw=/^adm-zip/]',
            message:
              "require() of adm-zip is banned (CVE-2026-76845, fail-open): import { AdmZip } from '@attentionform/astra-engine' (or the local './adm-zip-safe' wrapper) instead.",
          },
        ],
      },
    },
    {
      // The hardened wrapper is the ONLY production file allowed to import raw adm-zip.
      files: ['**/adm-zip-safe.ts'],
      rules: {
        'no-restricted-imports': 'off',
      },
    },
    {
      // Lock the 6 converted tests to the wrapper so they cannot regress to raw.
      files: [
        'packages/mcp-server/tests/unit/compose-document.test.ts',
        'packages/mcp-server/tests/unit/mcp-normalized-envelope.test.ts',
        'packages/cli/tests/cli-batch-safety.test.ts',
        'packages/cli/tests/cli-smoke.test.ts',
        'packages/cli/tests/cli-normalize.test.ts',
        'packages/cli/tests/cli-dx-wave.test.ts',
      ],
      rules: {
        'no-restricted-imports': [
          'error',
          {
            paths: [
              {
                name: 'adm-zip',
                message:
                  "This test must use the hardened wrapper (CVE-2026-76845): import { AdmZip } from '@attentionform/astra-engine' instead of raw 'adm-zip'.",
              },
            ],
            patterns: [
              {
                group: ['adm-zip/*'],
                message:
                  "This test must use the hardened wrapper (CVE-2026-76845): import { AdmZip } from '@attentionform/astra-engine' instead of raw 'adm-zip/*'.",
              },
            ],
          },
        ],
        // no-restricted-imports misses dynamic import() of adm-zip: ban it here.
        // TemplateLiteral forms (dynamic import of `adm-zip` / require of
        // `adm-zip`) bypass the Literal-only check above, so ban the
        // adm-zip-prefixed first quasi explicitly (prefix regex has no slash:
        // esquery-safe).
        'no-restricted-syntax': [
          'error',
          {
            selector: 'ImportExpression[source.value=/^adm-zip/]',
            message:
              "This test must use the hardened wrapper (CVE-2026-76845): dynamic import() of adm-zip is banned, import { AdmZip } from '@attentionform/astra-engine' instead.",
          },
          {
            selector: 'ImportExpression[source.quasis.0.value.raw=/^adm-zip/]',
            message:
              "This test must use the hardened wrapper (CVE-2026-76845): dynamic import() of adm-zip is banned, import { AdmZip } from '@attentionform/astra-engine' instead.",
          },
          {
            selector:
              'CallExpression[callee.name="require"][arguments.0.quasis.0.value.raw=/^adm-zip/]',
            message:
              "This test must use the hardened wrapper (CVE-2026-76845): require() of adm-zip is banned, import { AdmZip } from '@attentionform/astra-engine' instead.",
          },
        ],
      },
    },
  ],
};
