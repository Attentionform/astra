#!/usr/bin/env node
/**
 * @attentionform/astra — Track-B command line (`astra` bin).
 *
 * Commands (all fail-closed — any operational failure exits non-zero):
 *   inspect  <docx> [--fidelity compact|standard|full] [--section <id>] [--max-text <n>] [--width <n>] [--wrap] [--visual] [--format png|jpeg] [--max-ratio <n>] [--max-uncompressed-bytes <n>]
 *   validate <docx> [--max-ratio <n>] [--max-uncompressed-bytes <n>] [--style-guide <name|path>]
 *   sanitize <in.docx> <out.docx> [--client-safe] [--redact-pii] [--zero-retention] [--replacement-author <name>] [--accept-all-revisions] [--strip-comments] [--metadata-only]
 *     (bare sanitize = metadata preset: zero-retention + accept-all-revisions + strip-comments + neutral author, body text kept; --redact-pii/--client-safe also redact body PII)
 *   batch    <docx> [<ops.json|->] [--target-text <needle> --append <text>] [--append <text>] [--out <docx>] [--sequential|--no-sequential] [--no-validate] [--schema] [--max-ratio <n>] [--max-uncompressed-bytes <n>] [--dry-run] [--backup[=<path>]] [--matrix <vars.csv>] [--reproducible|--no-reproducible] [--stream [--on-error abort|continue] [--max-errors <n>]] [--style-guide <name|path>]
 *     (ops path `-` reads ops JSON from stdin; see `batch --help`)
 *     (ops support CLI-level `insert type:heading level:1..3` and `{action:'clearDoc'}`; see `batch --schema`)
 *   normalize <docx> [--out <docx>] [--dry-run] [--backup[=<path>]] (repair-as-action for legacy files; `validate` stays read-only, no --fix alias)
 *   export   <docx> [--format html|markdown|md|pdf] [-o <outPath>]
 *   search   <docx> <query> [--limit <n>]
 *   create   <out.docx> [--template report|memo|blank] [--force] [--from-md <file.md>] [--allowed-dir <dir>] [--dollar-math] [--style-guide <name|path>]
 *   import   <file.md> [-o <out.docx>] [--template report|memo|blank] [--force] [--allowed-dir <dir>] [--dollar-math]
 *   stats    <docx>
 *
 * Exit codes: 0 ok · 1 operational failure (invalid docx, bad ops, batch
 * failure/rollback, failed validation — never writes corruption) · 2 usage.
 *
 * Engine wiring:
 *   imports resolve to the `@attentionform/astra-engine` package.
 *   CLI command names, flags and exit-code contract stay identical
 *   across refactors (scripts pin them).
 */

import { run } from './cli/run';
import { CLI_NAME, EXIT_FAIL, printStderr } from './cli/output';

// MS-Q2a PART1_QUALITY Q2.1: cli.ts is now a thin dispatcher + barrel.
// All logic moved verbatim to ./cli/* modules (output/errors/usage/io +
// per-command modules + run). This file re-exports the old paths so
// `import { run, resolveInspectWidth } from '../src/cli'` keeps working
// with zero test changes. Hand-rolled argv preserved (split ≠ swap).
export * from './cli/output';
export * from './cli/errors';
export * from './cli/usage';
export * from './cli/io';
export * from './cli/inspect-cmd';
export * from './cli/validate-cmd';
export * from './cli/sanitize-cmd';
export * from './cli/batch-cmd';
export * from './cli/normalize-cmd';
export * from './cli/export-cmd';
export * from './cli/search-cmd';
export * from './cli/create-cmd';
export * from './cli/import-cmd';
export * from './cli/stats-cmd';
export * from './cli/visual-cmd';
export * from './cli/history-cmd';
export * from './cli/run';

function isDirectEntry(): boolean {
  const entry = process.argv?.[1] ?? '';
  return (
    entry.endsWith('packages/cli/src/cli.ts') ||
    entry.includes('packages/cli/dist/') ||
    /(^|[/\\])astra$/.test(entry)
  );
}

if (isDirectEntry()) {
  run(process.argv.slice(2)).then(
    (code) => {
      process.exitCode = code;
    },
    (err: unknown) => {
      const msg = err instanceof Error ? err.message : String(err);
      printStderr(`${CLI_NAME}: fatal: ${msg}`);
      process.exitCode = EXIT_FAIL;
    }
  );
}
