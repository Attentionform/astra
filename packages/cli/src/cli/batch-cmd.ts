import * as fsp from 'fs/promises';
import * as path from 'path';
import { BatchOperationsSchema, DocxParser, OperationEngine, chartHoleSizeSchemaRemedy, chartTypeSchemaRemedy, getEngineRemedy, resolveStyleGuide, summarizeStyleGuide, type DryRunPreview, type StyleGuideConfig } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_FAIL, EXIT_OK, type GlobalOpts, formatNormalizedInline, mergeNormalized, previewNormalizedForDisplay, print, printJson, printStderr, readEngineNormalized, readParserNormalizations, stringifyJsonPayload } from './output';
import { classifyFileReadError, cleanCliMessage, extractEngineRemedy, failWithCode, inferBatchCode, readErrRemedyProp, remedyForPostBatchValidation, rewriteMcpRemedies, withGatedEngineOutput, type KnownCliCode } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { type ZipCeilingOverrides, clearDocumentBody, escapeXmlText, macroFlagsUsageError, macroOptsFromFlags, parseMaxRatioFlag, parseMaxUncompressedBytesFlag, readDocx, readStdinText, withReproducibleEnv, withAuthorEnv, withZipCeilingEnv, writeDocx } from './io';

/** Canonical per-action ops examples (`batch --schema`). */
export function batchSchemaPayload(): Record<string, unknown> {
  return {
    $comment: 'Replace <nodeId> with ids from `astra inspect`. CLI-only ops are marked [cli].',
    replace: { action: 'replace', target: '<nodeId>', text: 'New text' },
    insert_paragraph: { action: 'insert', type: 'paragraph', anchor: '<nodeId>', position: 'after', text: 'New paragraph' },
    insert_heading_cli: { action: 'insert', type: 'heading', anchor: '<nodeId>', position: 'after', level: 1, text: 'Section title' },
    insert_pageBreak: { action: 'insert', type: 'pageBreak', anchor: '<nodeId>', position: 'after' },
    insert_table: {
      action: 'insert',
      type: 'table',
      anchor: '<nodeId>',
      position: 'after',
      headers: ['A', 'B'],
      rows: [['1', '2']],
    },
    style: { action: 'style', target: '<nodeId>', bold: true },
    namedStyle_apply: { action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: '<nodeId>' },
    delete: { action: 'delete', target: '<nodeId>' },
    move: { action: 'move', target: '<nodeId>', anchor: '<nodeId>', position: 'after' },
    table_updateCell: { action: 'table', operation: 'updateCell', target: '<tableNodeId>', rowIndex: 1, columnIndex: 1, text: 'Cell' },
    table_deleteRow: { action: 'table', operation: 'deleteRow', target: '<tableNodeId>', index: 2 },
    chart_insert: { action: 'chart', operation: 'insert', anchor: '<nodeId>', chartType: 'column', title: 'Quarterly revenue', categories: ['Q1', 'Q2'], series: [{ name: 'Revenue', values: [10, 20] }] },
    chart_update: { action: 'chart', operation: 'update', target: '<chartNodeId>', title: 'Quarterly revenue (revised)', categories: ['Q1', 'Q2'], series: [{ name: 'Revenue', values: [12, 22] }] },
    list_create: { action: 'list', operation: 'create', target: '<nodeId>', listType: 'bullet' },
    comment_insert: { action: 'comment', operation: 'insert', target: '<nodeId>', text: 'Review note' },
    comment_reply: { action: 'comment', operation: 'reply', target: '<commentId>', text: 'Agreed' },
    callout_insert: { action: 'callout', anchor: '<nodeId>', position: 'after', theme: 'amber', title: 'Field note', text: 'Callout body' },
    footnote_insert: { action: 'footnote', anchor: '<nodeId>', text: 'Note text' },
    endnote_insert: { action: 'endnote', anchor: '<nodeId>', text: 'Note text' },
    clearDoc_cli: { action: 'clearDoc' },
  };
}

export interface OpsFile {
  operations: unknown[];
  sequential?: boolean;
  rollbackOnError?: boolean;
  maxSteps?: number;
  maxRatio?: number;
  maxUncompressedBytes?: number;
}

export function parseOpsFile(raw: string): OpsFile {
  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw new Error(
      `E_VALIDATION: ops file is not valid JSON: ${msg}. Remedy: pass a bare operations array or { "operations": [ ... ] } as valid JSON.`
    );
  }
  if (Array.isArray(parsed)) return { operations: parsed };
  if (typeof parsed === 'object' && parsed !== null && Array.isArray((parsed as { operations?: unknown }).operations)) {
    return parsed as OpsFile;
  }
  throw new Error(
    'E_VALIDATION: ops file must be a bare operations array or { "operations": [ ... ] }. Remedy: pass { "operations": [{ "action": "replace", "target": "<nodeId>", "text": "..." }], "sequential": true }.'
  );
}

export function printBatchSchema(out: GlobalOpts): number {
  const payload = batchSchemaPayload();
  if (!out.json) console.log(stringifyJsonPayload(payload, 2));
  if (out.json) console.log(stringifyJsonPayload({ ok: true, schema: payload }));
  return EXIT_OK;
}

/**
 * Expand CLI-level op extensions before the engine schema gate:
 * - `insert type:heading level:1..3` -> `insert type:xml` carrying a
 *   Heading1..3 paragraph (the engine insert schema has no heading type, and
 *   this CLI must not grow engine-side edges — the rewrite keeps one op, one
 *   anchor, no extra ids).
 * - `{action:'clearDoc'}` (alias `clear-doc`) -> counted separately; the body
 *   reset itself runs via clearDocumentBody() before the remaining ops.
 */
export function expandCliOps(operations: unknown[]): { ops: unknown[]; clearCount: number } {
  const ops: unknown[] = [];
  let clearCount = 0;
  for (const raw of operations) {
    if (raw !== null && typeof raw === 'object' && !Array.isArray(raw)) {
      const rec = raw as Record<string, unknown>;
      if (rec.action === 'clearDoc' || rec.action === 'clear-doc') {
        const keys = Object.keys(rec);
        if (keys.length !== 1) {
          throw new Error(
            `E_VALIDATION: clearDoc takes no fields (got ${keys.join(',')}). Remedy: pass exactly { "action": "clearDoc" }.`
          );
        }
        clearCount += 1;
        continue;
      }
      if (rec.action === 'insert' && rec.type === 'heading') {
        const levelRaw = rec.level === undefined ? 1 : Number(rec.level);
        if (!Number.isInteger(levelRaw) || levelRaw < 1 || levelRaw > 3) {
          throw new Error(
            `E_VALIDATION: insert type:heading needs level 1..3 (got ${String(rec.level)}). Remedy: pass { "action": "insert", "type": "heading", "level": 1|2|3, "anchor": "<nodeId>", "position": "after", "text": "..." }.`
          );
        }
        if (typeof rec.anchor !== 'string' || rec.anchor.length === 0) {
          throw new Error(
            'E_VALIDATION: insert type:heading needs a string anchor. Remedy: re-inspect ids via `astra inspect` and pass { "action": "insert", "type": "heading", "anchor": "<nodeId>", "position": "after", "text": "..." }.'
          );
        }
        const position = rec.position === undefined ? 'after' : rec.position;
        if (position !== 'before' && position !== 'after') {
          throw new Error(
            `E_VALIDATION: insert type:heading needs position before|after (got ${String(rec.position)}). Remedy: pass position "before" or "after".`
          );
        }
        const text = rec.text === undefined ? '' : String(rec.text);
        const xml =
          `<w:p><w:pPr><w:pStyle w:val="Heading${levelRaw}"/></w:pPr>` +
          `<w:r><w:t xml:space="preserve">${escapeXmlText(text)}</w:t></w:r></w:p>`;
        const rewritten: Record<string, unknown> = {
          action: 'insert',
          type: 'xml',
          anchor: rec.anchor,
          position,
          xml,
        };
        if (typeof rec.id === 'string' && rec.id.length > 0) rewritten.id = rec.id;
        ops.push(rewritten);
        continue;
      }
    }
    ops.push(raw);
  }
  return { ops, clearCount };
}

// ── Batch safety: --dry-run / --backup / --matrix ───────────────────────────
// Undo choice: a failed batch already rolls back in memory, but a successful
// batch keeps no persistent history across CLI invocations (that would need
// on-disk history — a redesign). Undo across runs is therefore the --backup
// snapshot: restore the .bak copy. Documented in `batch --help`.

export interface BackupState {
  /** Resolved backup path, or undefined when --backup was not given or this is a dry run. */
  path: string | undefined;
  done: boolean;
}

export interface BatchJob {
  opsFile: OpsFile;
  target: string;
  rowLabel: string | undefined;
}

export interface RunBatchJobCtx {
  file: string;
  docx: Buffer;
  opsFile: OpsFile;
  target: string;
  sequential: boolean | undefined;
  skipValidate: boolean;
  flagCeilings: ZipCeilingOverrides;
  dryRun: boolean;
  backupState: BackupState;
  rowLabel: string | undefined;
  reproducible: boolean | undefined;
  macroOpts: { allowMacros: true } | { stripMacros: true } | undefined;
  allowedDir?: string | undefined;
  /** Resolved --style-guide config (omitted = built-in defaults, byte-identical). */
  styleGuide?: StyleGuideConfig | undefined;
  /** Parser diagnostics of the resolved --style-guide (ride JSON styleWarnings). */
  styleGuideWarnings?: string[] | undefined;
}

export function resolveBackupPath(file: string, raw: string | undefined): string {
  if (raw !== undefined && raw.length > 0) return raw;
  return `${file}.bak.docx`;
}

/**
 * Snapshot the pre-batch input bytes before the first write. Returns an
 * exit code when the backup itself fails closed (target left untouched),
 * or undefined when the backup is done (or not requested / already done).
 */
export async function ensureBackupCopy(
  out: GlobalOpts,
  file: string,
  docx: Buffer,
  target: string,
  backupState: BackupState
): Promise<number | undefined> {
  const backupPath = backupState.path;
  if (backupPath === undefined || backupState.done) return undefined;
  const absBackup = path.resolve(backupPath);
  if (absBackup === path.resolve(file) || absBackup === path.resolve(target)) {
    return failWithCode(
      out,
      `E_VALIDATION: --backup path "${backupPath}" must differ from the input and the output (refusing to overwrite either).`,
      'E_VALIDATION',
      undefined,
      'pass a distinct backup path or bare --backup for the default <input>.bak.docx',
      backupPath
    );
  }
  try {
    await writeDocx(backupPath, docx);
  } catch (err: unknown) {
    const rawMsg = err instanceof Error ? err.message : String(err);
    return failWithCode(
      out,
      rawMsg,
      'E_IO',
      undefined,
      'check the backup directory is writable, then retry (nothing was written)',
      err
    );
  }
  backupState.done = true;
  print(out, `BACKUP ${file} -> ${backupPath}`);
  return undefined;
}

/** Split one CSV line (double-quote escaping, "" = literal quote). */
export function splitCsvLine(line: string, source: string): string[] {
  const fields: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (inQuotes) {
      if (c === '"') {
        if (line[i + 1] === '"') {
          cur += '"';
          i++;
        } else {
          inQuotes = false;
        }
      } else {
        cur += c;
      }
    } else if (c === '"') {
      inQuotes = true;
    } else if (c === ',') {
      fields.push(cur);
      cur = '';
    } else {
      cur += c;
    }
  }
  fields.push(cur);
  if (inQuotes) {
    throw new Error(
      `E_VALIDATION: matrix CSV "${source}" has an unterminated quote. Remedy: balance the double quotes on every line and retry.`
    );
  }
  return fields;
}

/**
 * Minimal CSV reader for --matrix: first non-blank line is the header,
 * every later non-blank line is one data row (one row = one --out).
 */
export function parseMatrixCsv(raw: string, source: string): { headers: string[]; rows: Array<Record<string, string>> } {
  const lines = raw.replace(/\r\n?/g, '\n').split('\n');
  const records: string[][] = [];
  for (const line of lines) {
    if (line.trim() === '') continue;
    records.push(splitCsvLine(line, source));
  }
  if (records.length === 0) {
    throw new Error(
      `E_VALIDATION: matrix CSV "${source}" is empty. Remedy: pass a CSV whose first line is a header row (e.g. name,company) followed by one data row per output.`
    );
  }
  const headers = (records[0] as string[]).map((h) => h.trim());
  headers.forEach((h, i) => {
    if (h.length === 0) {
      throw new Error(
        `E_VALIDATION: matrix CSV "${source}" header column ${i + 1} is empty. Remedy: name every header column (e.g. name,company) and retry.`
      );
    }
  });
  const dupes = headers.filter((h, i) => headers.indexOf(h) !== i);
  if (dupes.length > 0) {
    throw new Error(
      `E_VALIDATION: matrix CSV "${source}" has a duplicate header "${dupes[0]}". Remedy: keep header names unique and retry.`
    );
  }
  if (records.length === 1) {
    throw new Error(
      `E_VALIDATION: matrix CSV "${source}" has a header but no data rows. Remedy: add one data row per output and retry.`
    );
  }
  const rows: Array<Record<string, string>> = [];
  for (let r = 1; r < records.length; r++) {
    const fields = records[r] as string[];
    if (fields.length !== headers.length) {
      throw new Error(
        `E_VALIDATION: matrix CSV "${source}" row ${r} has ${fields.length} fields but the header has ${headers.length}. Remedy: give every row exactly ${headers.length} comma-separated values and retry.`
      );
    }
    const row: Record<string, string> = {};
    headers.forEach((h, i) => {
      row[h] = fields[i] as string;
    });
    rows.push(row);
  }
  return { headers, rows };
}

/**
 * Literal {{var}} substitution: header names only, single pass over the
 * original string (values are never re-scanned), no expressions,
 * conditionals, loops, or filters. Unknown names fail closed.
 */
export const MATRIX_VAR_RE = /\{\{\s*([A-Za-z0-9_.-]+)\s*\}\}/g;

export function substituteMatrixString(template: string, row: Record<string, string>, ctx: string): string {
  return template.replace(MATRIX_VAR_RE, (_match, key: string) => {
    if (!Object.prototype.hasOwnProperty.call(row, key)) {
      throw new Error(
        `E_VALIDATION: ${ctx} references {{${key}}} but the CSV has no column "${key}". Remedy: add the column to the header row or drop the placeholder (literal {{var}} scope: header names only).`
      );
    }
    return row[key] as string;
  });
}

export function substituteMatrixVars(value: unknown, row: Record<string, string>, ctx: string): unknown {
  if (typeof value === 'string') return substituteMatrixString(value, row, ctx);
  if (Array.isArray(value)) return value.map((v) => substituteMatrixVars(v, row, ctx));
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>).map(([k, v]) => [k, substituteMatrixVars(v, row, ctx)])
    );
  }
  return value;
}

export function truncLine(text: string, max: number): string {
  const flat = String(text).replace(/\s+/g, ' ');
  return flat.length > max ? `${flat.slice(0, max - 3)}...` : flat;
}

/** Unified diff of the affected nodes (changed nodes only). */
export function renderUnifiedDiff(preview: DryRunPreview[]): string {
  const lines: string[] = [];
  for (const p of preview) {
    if (p.before === p.after) continue;
    lines.push(`--- a/${p.nodeId} (before)`);
    lines.push(`+++ b/${p.nodeId} (after)`);
    if (p.before !== null) {
      for (const ln of p.before.split('\n')) lines.push(`- ${truncLine(ln, 1000)}`);
    }
    if (p.after !== null) {
      for (const ln of p.after.split('\n')) lines.push(`+ ${truncLine(ln, 1000)}`);
    }
  }
  return lines.join('\n');
}

/** Human + --json report for a successful dry run (writes nothing). */
export function reportDryRun(
  out: GlobalOpts,
  args: {
    file: string;
    target: string;
    appliedCount: number;
    cleared: number;
    preview: DryRunPreview[];
    normalized: string[];
    rowLabel: string | undefined;
    possiblyStaleBodyRefs?: { nodeId: string; preview: string }[];
  }
): void {
  const { file, target, appliedCount, cleared, preview, normalized, rowLabel } = args;
  const staleRefs = Array.isArray(args.possiblyStaleBodyRefs)
    ? args.possiblyStaleBodyRefs.filter(
        (v): v is { nodeId: string; preview: string } =>
          typeof v === 'object' &&
          v !== null &&
          typeof (v as { nodeId?: unknown }).nodeId === 'string' &&
          typeof (v as { preview?: unknown }).preview === 'string'
      )
    : [];
  const rowSuffix = rowLabel !== undefined ? ` [${rowLabel}]` : '';
  const changed = preview.filter((p) => p.before !== p.after);
  print(
    out,
    `DRY-RUN OK ${file} -> ${target} (would apply=${appliedCount}${cleared > 0 ? ` cleared=${cleared}` : ''} affected=${preview.length} would-be ${formatNormalizedInline(normalized)}${rowSuffix}, writes nothing)`
  );
  for (const p of preview) {
    if (p.before === null) print(out, `affected ${p.nodeId}: (new) "${truncLine(p.after ?? '', 120)}"`);
    else if (p.after === null) print(out, `affected ${p.nodeId}: "${truncLine(p.before, 120)}" -> (deleted)`);
    else if (p.before !== p.after) {
      print(out, `affected ${p.nodeId}: "${truncLine(p.before, 120)}" -> "${truncLine(p.after, 120)}"`);
    } else print(out, `affected ${p.nodeId}: (no text change)`);
  }
  for (const n of normalized) print(out, `would-be normalize: ${truncLine(n, 500)}`);
  for (const s of staleRefs) print(out, `possibly-stale ${s.nodeId}: "${truncLine(s.preview, 120)}"`);
  const diff = renderUnifiedDiff(preview);
  print(out, changed.length > 0 ? diff : '(no text changes)');
  printJson(out, {
    ok: true,
    dryRun: true,
    file,
    output: target,
    appliedCount,
    cleared,
    affected: preview,
    normalized,
    diff,
    ...(staleRefs.length > 0 ? { possiblyStaleBodyRefs: staleRefs } : {}),
    ...(rowLabel !== undefined ? { row: rowLabel } : {}),
  });
}

export async function cmdBatch(out: GlobalOpts, args: string[]): Promise<number> {
  if (args.includes('--schema')) return printBatchSchema(out);
  if (hasHelpFlag(args)) return printCommandHelp(out, 'batch');
  let file: string | undefined;
  let opsPath: string | undefined;
  let outPath: string | undefined;
  let sequential: boolean | undefined;
  let skipValidate = false;
  let maxRatioRaw: string | undefined;
  let maxBytesRaw: string | undefined;
  let targetTextRaw: string | undefined;
  let appendRaw: string | undefined;
  let dryRun = false;
  let force = false;
  let backupGiven = false;
  let backupRaw: string | undefined;
  let matrixRaw: string | undefined;
  let reproducible: boolean | undefined;
  let noReproducible = false;
  let author: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  let allowedDirRaw: string | undefined;
  let styleGuideRaw: string | undefined;
  let stream = false;
  let onErrorRaw: string | undefined;
  let maxErrorsRaw: string | undefined;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--out') outPath = args[++i];
    else if (a === '--sequential') sequential = true;
    else if (a === '--no-sequential') sequential = false;
    else if (a === '--no-validate') skipValidate = true;
    else if (a === '--dry-run') dryRun = true;
    else if (a === '--force') force = true;
    else if (a === '--reproducible') reproducible = true;
    else if (a === '--no-reproducible') noReproducible = true;
    else if (a === '--author' || a.startsWith('--author=')) {
      const value = a === '--author' ? args[++i] : a.slice('--author='.length);
      if (value === undefined || value.trim().length === 0 || (a === '--author' && value.startsWith('-')) || value.trim().length > 128) {
        return usage(out, 'batch --author requires a name of 1 to 128 characters (the default author for tracked changes and comments that do not name one)');
      }
      author = value.trim();
    }
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a === '--allowed-dir') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'batch --allowed-dir requires a value (directory holding image files for filePath inserts)');
      }
      allowedDirRaw = args[++i];
    } else if (a.startsWith('--allowed-dir=')) {
      const value = a.slice('--allowed-dir='.length);
      if (value.length === 0) {
        return usage(out, 'batch --allowed-dir requires a value (directory holding image files for filePath inserts)');
      }
      allowedDirRaw = value;
    }
    else if (a === '--style-guide') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'batch --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
      }
      styleGuideRaw = args[++i];
    } else if (a.startsWith('--style-guide=')) {
      const value = a.slice('--style-guide='.length);
      if (value.length === 0) {
        return usage(out, 'batch --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
      }
      styleGuideRaw = value;
    }
    else if (a === '--backup') {
      backupGiven = true;
      const nextBackup = args[i + 1];
      if (nextBackup !== undefined && nextBackup.length > 0 && !nextBackup.startsWith('-')) backupRaw = args[++i];
    } else if (a.startsWith('--backup=')) {
      backupGiven = true;
      backupRaw = a.slice('--backup='.length);
      if (backupRaw.length === 0) {
        return usage(out, 'batch --backup requires a path value or bare --backup for the default <input>.bak.docx');
      }
    } else if (a === '--matrix') {
      const nextMatrix = args[i + 1];
      if (nextMatrix === undefined || nextMatrix.length === 0 || nextMatrix.startsWith('-')) {
        return usage(out, 'batch --matrix requires a vars.csv value (e.g. --matrix vars.csv with --out "out-{{name}}.docx")');
      }
      matrixRaw = args[++i];
    } else if (a.startsWith('--matrix=')) {
      matrixRaw = a.slice('--matrix='.length);
      if (matrixRaw.length === 0) return usage(out, 'batch --matrix requires a vars.csv value (e.g. --matrix vars.csv)');
    } else if (a === '--target-text') {
      const next = args[i + 1];
      if (next === undefined) return usage(out, 'batch --target-text requires a value (text identifying exactly one node)');
      targetTextRaw = args[++i];
    } else if (a.startsWith('--target-text=')) targetTextRaw = a.slice('--target-text='.length);
    else if (a === '--append') {
      const next = args[i + 1];
      if (next === undefined) return usage(out, 'batch --append requires a value (text to write)');
      appendRaw = args[++i];
    } else if (a.startsWith('--append=')) appendRaw = a.slice('--append='.length);
    else if (a === '--max-ratio') maxRatioRaw = args[++i];
    else if (a.startsWith('--max-ratio=')) maxRatioRaw = a.slice('--max-ratio='.length);
    else if (a === '--stream') stream = true;
    else if (a === '--on-error') {
      const next = args[i + 1];
      if (next === undefined) return usage(out, 'batch --on-error requires abort|continue (default abort)');
      onErrorRaw = args[++i];
    } else if (a.startsWith('--on-error=')) onErrorRaw = a.slice('--on-error='.length);
    else if (a === '--max-errors') {
      const next = args[i + 1];
      if (next === undefined) return usage(out, 'batch --max-errors requires a non-negative integer (default 10; 0 disables the breaker)');
      maxErrorsRaw = args[++i];
    } else if (a.startsWith('--max-errors=')) maxErrorsRaw = a.slice('--max-errors='.length);
    else if (a === '--max-uncompressed-bytes') maxBytesRaw = args[++i];
    else if (a.startsWith('--max-uncompressed-bytes=')) maxBytesRaw = a.slice('--max-uncompressed-bytes='.length);
    else if (a.startsWith('--')) return usage(out, `batch: unknown flag "${a}"`);
    else if (!file) file = a;
    else if (!opsPath) opsPath = a;
    else return usage(out, 'batch takes exactly <docx> <ops.json>');
  }
  // Shortcut mode: text-addressed single edit, no ops file, no node ids.
  // Both shortcuts synthesize canonical ops that flow through the identical
  // downstream path below (CLI extensions, schema gate, engine rollback,
  // post-batch validation, exit 0/1/2) — only ops acquisition differs.
  const shortcut = targetTextRaw !== undefined || appendRaw !== undefined;
  if (targetTextRaw !== undefined && targetTextRaw.trim() === '') {
    return usage(out, 'batch --target-text requires a non-empty value (text identifying exactly one node)');
  }
  if (appendRaw === undefined && targetTextRaw !== undefined) {
    return usage(out, 'batch --target-text requires --append <replacement text> (e.g. `astra batch doc.docx --target-text "teh wrod" --append "the word"`)');
  }
  if (shortcut && opsPath !== undefined) {
    return usage(out, 'batch --target-text/--append cannot be combined with an <ops.json|-> argument (pass one or the other)');
  }
  if (!file || (!opsPath && !shortcut)) return usage(out, 'batch requires <docx> <ops.json>');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  if (reproducible === true && noReproducible === true) return usage(out, 'batch --reproducible and --no-reproducible cannot be combined (deterministic is the default; use --no-reproducible to restore wall-clock)');
  const effectiveReproducible: boolean | undefined = noReproducible ? false : reproducible;
  const macroOpts = macroOptsFromFlags({ allowMacros, stripMacros });
  // Jail for image `filePath` inserts: explicit --allowed-dir wins, else the
  // input document's own directory (never bare process.cwd()). Fail closed
  // when the jail is unresolvable; batches without `filePath` ops pass it
  // through inertly.
  let batchAllowedDir: string | undefined;
  {
    const candidate = allowedDirRaw ?? path.dirname(path.resolve(file as string));
    let candidateStat;
    try {
      candidateStat = await fsp.stat(candidate);
    } catch (err: unknown) {
      return failWithCode(
        out,
        `E_VALIDATION: batch --allowed-dir "${candidate}" does not exist or cannot be read.`,
        'E_VALIDATION',
        undefined,
        'pass an existing directory holding the image files (or omit --allowed-dir to use the input document folder)',
        err
      );
    }
    if (!candidateStat.isDirectory()) {
      return failWithCode(
        out,
        `E_VALIDATION: batch --allowed-dir "${candidate}" is not a directory.`,
        'E_VALIDATION',
        undefined,
        'pass an existing directory holding the image files (or omit --allowed-dir to use the input document folder)',
        `batch --allowed-dir not a directory: ${candidate}`
      );
    }
    try {
      batchAllowedDir = await fsp.realpath(candidate);
    } catch {
      batchAllowedDir = candidate;
    }
  }
  // Guide attach (fail-closed before any read/apply): unknown names, jail
  // escapes, and remote URLs refuse with E_VALIDATION/E_SANDBOX and write
  // nothing. File refs jail to --allowed-dir when one is in play, else the
  // input document folder. One guide per invocation; --stream applies it to
  // the whole stream.
  let styleGuide: StyleGuideConfig | undefined;
  let styleGuideWarnings: string[] = [];
  if (styleGuideRaw !== undefined) {
    try {
      const resolved = resolveStyleGuide(styleGuideRaw, {
        allowedDir: allowedDirRaw ?? path.dirname(path.resolve(file as string)),
      });
      styleGuide = resolved.config;
      styleGuideWarnings = [...resolved.warnings];
      for (const w of resolved.warnings) printStderr(`${CLI_NAME}: WARNING: ${w}`);
    } catch (err: unknown) {
      const rawMsg = err instanceof Error ? err.message : String(err);
      return failWithCode(out, rawMsg, inferBatchCode(err, 'E_VALIDATION'), undefined, undefined, err);
    }
  }
  if (matrixRaw !== undefined && outPath === undefined) {
    return usage(out, 'batch --matrix requires --out <path template with {{column}}> (one row = one --out; in-place writes are refused for matrix runs)');
  }
  const flagCeilings: ZipCeilingOverrides = {};
  try {
    if (maxRatioRaw !== undefined) flagCeilings.maxRatio = parseMaxRatioFlag(maxRatioRaw);
    if (maxBytesRaw !== undefined) flagCeilings.maxUncompressedBytes = parseMaxUncompressedBytesFlag(maxBytesRaw);
  } catch (err: unknown) {
    return usage(out, err instanceof Error ? err.message : String(err));
  }
  if (!stream && onErrorRaw !== undefined) {
    return usage(out, 'batch --on-error requires --stream (stream-only line-failure policy, e.g. `astra batch <docx> - --stream --on-error=continue`)');
  }
  if (!stream && maxErrorsRaw !== undefined) {
    return usage(out, 'batch --max-errors requires --stream (stream-only circuit breaker, e.g. `astra batch <docx> - --stream --max-errors=10`)');
  }
  if (stream) {
    return await withReproducibleEnv(effectiveReproducible, () => withAuthorEnv(author, () =>
      cmdBatchStream(out, {
        file,
        opsPath,
        outPath,
        sequential,
        skipValidate,
        flagCeilings,
        dryRun,
        backupGiven,
        backupRaw,
        matrixRaw,
        targetTextRaw,
        appendRaw,
        onErrorRaw,
        maxErrorsRaw,
        force,
        reproducible: effectiveReproducible,
        macroOpts,
        allowedDir: batchAllowedDir,
        styleGuide,
        styleGuideWarnings,
      })
    ));
  }

  let docx: Buffer;
  let opsFile: OpsFile;
  if (shortcut) {
    // Shortcut acquisition: no ops file. --append synthesizes an
    // anchor:end paragraph insert (no id needed); --target-text resolves via
    // the existing engine search to one node id and synthesizes a replace.
    // Everything below (extensions, schema gate, rollback, validation,
    // exit codes) is the shared ops-file path unchanged.
    try {
      docx = await readDocx(file);
    } catch (err: unknown) {
      const read = classifyFileReadError(err, file);
      if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, read.remedy, err);
      const rawMsg = err instanceof Error ? err.message : String(err);
      const code = /ENOENT|no such file/i.test(rawMsg) ? 'E_IO' : inferBatchCode(err, 'E_VALIDATION');
      return failWithCode(
        out,
        rawMsg,
        code,
        undefined,
        'check the file paths exist and are readable, then retry',
        err
      );
    }
    const appendText = appendRaw as string;
    if (targetTextRaw !== undefined) {
      const needle = targetTextRaw;
      let hits: Array<{ id: string }>;
      try {
        // Same parser + ceilings + gating as the batch load below; this AST
        // is resolution-only (search never mutates) and discarded — the
        // shared path re-parses for the apply.
        const resolver = await withZipCeilingEnv(flagCeilings, () => new DocxParser(docx, macroOpts));
        const resolverAst = await withGatedEngineOutput(out, () => resolver.getAST());
        hits = resolver.searchNodesInAST(resolverAst, needle);
      } catch (err: unknown) {
        return failWithCode(
          out,
          err instanceof Error ? err.message : String(err),
          inferBatchCode(err, 'E_VALIDATION'),
          undefined,
          undefined,
          err
        );
      }
      if (hits.length === 0) {
        return failWithCode(
          out,
          `E_VALIDATION: batch --target-text ${JSON.stringify(needle)} matched 0 nodes, nothing written.`,
          'E_VALIDATION',
          undefined,
          'run `astra search <docx> <text>` to list matches, then retry with text matching exactly one node (or use explicit node ids via `astra inspect` + an ops file)',
          `target-text matched 0 nodes for ${needle}`
        );
      }
      if (hits.length > 1 || hits[0] === undefined) {
        const ids = hits.slice(0, 3).map((h) => h.id).join(',');
        return failWithCode(
          out,
          `E_VALIDATION: batch --target-text ${JSON.stringify(needle)} matched ${hits.length} node${hits.length === 1 ? '' : 's'} (${ids}${hits.length > 3 ? ',...' : ''}), refusing to guess which to edit, nothing written.`,
          'E_VALIDATION',
          undefined,
          'refine --target-text to match exactly one node (see `astra search <docx> <text>`), or use explicit node ids via `astra inspect` + an ops file',
          `target-text matched ${hits.length} nodes for ${needle}`
        );
      }
      opsFile = { operations: [{ action: 'replace', target: hits[0].id, text: appendText }] };
    } else {
      opsFile = {
        operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: appendText }],
      };
    }
  } else {
    if (opsPath === undefined) return usage(out, 'batch requires <docx> <ops.json>');
    const opsArg = opsPath;
    let opsRaw: string;
    // R-44: ops path `-` reads ops JSON from stdin (same parseOpsFile path,
    // same errors/exit codes as a file). TTY + empty guards return usage
    // errors (exit 2) so a missing pipe fails fast instead of hanging.
    if (opsArg === '-' && process.stdin.isTTY) {
      return usage(
        out,
        'E_VALIDATION: batch ops path "-" reads ops JSON from stdin, but stdin is a terminal (nothing piped). Remedy: pipe ops JSON, e.g. `cat ops.json | astra batch <docx> -`, or pass an ops file path.'
      );
    }
    try {
      if (opsArg === '-') {
        const [docBuf, stdinRaw] = await Promise.all([readDocx(file), readStdinText()]);
        if (stdinRaw.trim() === '') {
          return usage(
            out,
            'E_VALIDATION: batch ops path "-" read empty stdin. Remedy: pipe a bare operations array or { "operations": [ ... ] } as JSON, e.g. `cat ops.json | astra batch <docx> -`.'
          );
        }
        docx = docBuf;
        opsRaw = stdinRaw;
      } else {
        [docx, opsRaw] = await Promise.all([readDocx(file), fsp.readFile(opsArg, 'utf8')]);
      }
    } catch (err: unknown) {
      // F5-CLI: IO failures also carry a machine code (human + --json).
      // Missing files are E_IO; engine-thrown codes pass through unchanged.
      // Read failures (missing/corrupt/directory, docx or ops side) use the
      // single normalized shape; anything else keeps the existing path.
      const rawMsg = err instanceof Error ? err.message : String(err);
      const namesOps = opsArg !== '-' && rawMsg.includes(opsArg);
      const read = classifyFileReadError(err, namesOps ? opsArg : file);
      if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, read.remedy, err);
      const code = /ENOENT|no such file/i.test(rawMsg) ? 'E_IO' : inferBatchCode(err, 'E_VALIDATION');
      return failWithCode(
        out,
        rawMsg,
        code,
        undefined,
        'check the file paths exist and are readable, then retry',
        err
      );
    }
    try {
      opsFile = parseOpsFile(opsRaw);
    } catch (err: unknown) {
      return failWithCode(
        out,
        err instanceof Error ? err.message : String(err),
        inferBatchCode(err, 'E_VALIDATION'),
        undefined,
        undefined,
        err
      );
    }
  }
  // Build jobs: a plain run is one job; --matrix fans out to one job per
  // CSV data row (ops strings + --out substituted literally per row, each
  // row starting from the pristine input). All row outputs resolve before
  // any write so a duplicate/guessing collision fails closed with nothing
  // written. A row failure stops the run (earlier row outputs stay — each
  // row is atomic, the matrix as a whole is not).
  let jobs: BatchJob[];
  if (matrixRaw !== undefined) {
    const csvSource = matrixRaw;
    let csvRaw: string;
    try {
      csvRaw = await fsp.readFile(csvSource, 'utf8');
    } catch (err: unknown) {
      const read = classifyFileReadError(err, csvSource);
      if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, 'check the vars.csv path exists and is readable, then retry', err);
      const rawMsg = err instanceof Error ? err.message : String(err);
      return failWithCode(out, rawMsg, 'E_IO', undefined, 'check the vars.csv path exists and is readable, then retry', err);
    }
    let headers: string[];
    let rows: Array<Record<string, string>>;
    try {
      ({ headers, rows } = parseMatrixCsv(csvRaw, csvSource));
    } catch (err: unknown) {
      return failWithCode(out, err instanceof Error ? err.message : String(err), 'E_VALIDATION', undefined, undefined, err);
    }
    void headers;
    const outTemplate = outPath as string;
    const seenOuts = new Set<string>();
    jobs = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r] as Record<string, string>;
      const ctx = `matrix row ${r + 1}`;
      let rowOut: string;
      let rowOps: unknown;
      try {
        rowOut = substituteMatrixString(outTemplate, row, `--out ${ctx}`);
        rowOps = substituteMatrixVars(opsFile.operations, row, ctx);
      } catch (err: unknown) {
        return failWithCode(out, err instanceof Error ? err.message : String(err), 'E_VALIDATION', undefined, undefined, err);
      }
      if (rowOut.trim() === '') {
        return failWithCode(
          out,
          `E_VALIDATION: ${ctx} produced an empty --out path.`,
          'E_VALIDATION',
          undefined,
          'give --out a template with literal {{column}} placeholders that resolve to a non-empty path per row',
          ctx
        );
      }
      const absOut = path.resolve(rowOut);
      if (seenOuts.has(absOut)) {
        return failWithCode(
          out,
          `E_VALIDATION: ${ctx} resolves to the same --out as an earlier row ("${rowOut}").`,
          'E_VALIDATION',
          undefined,
          'make the --out template unique per row (e.g. include a {{column}} that differs per row)',
          rowOut
        );
      }
      seenOuts.add(absOut);
      jobs.push({ opsFile: { ...opsFile, operations: rowOps as unknown[] }, target: rowOut, rowLabel: ctx });
    }
  } else {
    jobs = [{ opsFile, target: (outPath ?? file) as string, rowLabel: undefined }];
  }
  // An explicit --out target that exists refuses without --force (create/
  // import parity); in-place writes and dry runs never gate. Matrix row
  // targets all resolve before any write so a collision fails closed first.
  if (!dryRun && !force && outPath !== undefined) {
    for (const job of jobs) {
      try {
        await fsp.access(job.target);
      } catch {
        continue;
      }
      return usage(out, `batch: --out target "${path.basename(job.target)}" already exists (refusing to overwrite without --force)`);
    }
  }
  const backupState: BackupState = {
    path: backupGiven && !dryRun ? resolveBackupPath(file, backupRaw) : undefined,
    done: false,
  };
  if (backupGiven && dryRun) {
    printStderr(`${CLI_NAME}: note: --backup skipped: --dry-run writes nothing`);
  }
  for (const job of jobs) {
    const code = await withReproducibleEnv(effectiveReproducible, () => withAuthorEnv(author, () => runBatchJob(out, {
      file,
      docx,
      opsFile: job.opsFile,
      target: job.target,
      sequential,
      skipValidate,
      flagCeilings,
      dryRun,
      backupState,
      rowLabel: job.rowLabel,
      reproducible: effectiveReproducible,
      macroOpts,
      allowedDir: batchAllowedDir,
      styleGuide,
      styleGuideWarnings,
    })));
    if (code !== EXIT_OK) return code;
  }
  if (matrixRaw !== undefined) {
    print(out, `MATRIX OK ${file} (rows=${jobs.length}${dryRun ? ', dry run, writes nothing' : ''})`);
    for (const job of jobs) print(out, `  ${job.target}`);
    printJson(out, { ok: true, matrix: true, file, rows: jobs.map((j) => j.target), dryRun });
  }
  return EXIT_OK;
}

export async function runBatchJob(out: GlobalOpts, ctx: RunBatchJobCtx): Promise<number> {
  const file = ctx.file;
  const docx = ctx.docx;
  const opsFile = ctx.opsFile;
  const target = ctx.target;
  const sequential = ctx.sequential;
  const skipValidate = ctx.skipValidate;
  const flagCeilings = ctx.flagCeilings;
  const rowSuffix = ctx.rowLabel !== undefined ? ` [${ctx.rowLabel}]` : '';
  // CLI-level op extensions (heading/clearDoc) expand before the engine gate.
  let engineOps: unknown[];
  let clearCount = 0;
  try {
    ({ ops: engineOps, clearCount } = expandCliOps(opsFile.operations));
  } catch (err: unknown) {
    return failWithCode(
      out,
      err instanceof Error ? err.message : String(err),
      inferBatchCode(err, 'E_VALIDATION'),
      undefined,
      undefined,
      err
    );
  }
  // Clear-only batches (only `{action:'clearDoc'}` ops) skip the engine
  // schema gate (its min(1) would reject the now-empty op list); the reset
  // itself is applied + validated below.
  const clearOnly = engineOps.length === 0 && clearCount > 0;
  const checked = clearOnly
    ? null
    : BatchOperationsSchema.safeParse({
        operations: engineOps,
        ...(opsFile.sequential !== undefined ? { sequential: opsFile.sequential } : {}),
        ...(opsFile.rollbackOnError !== undefined ? { rollbackOnError: opsFile.rollbackOnError } : {}),
        ...(opsFile.maxSteps !== undefined ? { maxSteps: opsFile.maxSteps } : {}),
        ...(opsFile.maxRatio !== undefined ? { maxRatio: opsFile.maxRatio } : {}),
        ...(opsFile.maxUncompressedBytes !== undefined ? { maxUncompressedBytes: opsFile.maxUncompressedBytes } : {}),
        ...(sequential !== undefined ? { sequential } : {}),
      });
  if (checked !== null && !checked.success) {
    const detail = checked.error.issues
      .slice(0, 3)
      .map((iss) => `${iss.path.join('.') || '(root)'}: ${iss.message}`)
      .join('; ');
    // F5-CLI: cap violations (too_big on operations/maxSteps) are E_QUOTA,
    // shape violations are E_VALIDATION — both in human + --json.
    // Array-level overflow only (path length 1): a field-length violation
    // (e.g. operations.0.match too long) is E_VALIDATION, never quota.
    const isQuota = checked.error.issues.some(
      (iss) =>
        iss.code === 'too_big' &&
        iss.path.length === 1 &&
        (String(iss.path[0] ?? '') === 'operations' || String(iss.path[0] ?? '') === 'maxSteps')
    );
    const code: 'E_VALIDATION' | 'E_QUOTA' = isQuota ? 'E_QUOTA' : 'E_VALIDATION';
    const remedy =
      code === 'E_QUOTA'
        ? 'split into batches of <= 500 operations (sequential:true) with an optional positive-integer maxSteps (omit for unbounded) and retry'
        : chartTypeSchemaRemedy(detail) ?? chartHoleSizeSchemaRemedy(detail) ?? 'single edit = 1-item batch with canonical shapes';
    return failWithCode(out, `invalid operations: ${detail}.`, code, { issues: detail }, remedy, checked.error);
  }
  // R-8: effective ceilings — explicit CLI flags win over ops-file values,
  // either wins over env/default (env applies inside the DocxParser gate).
  const fileCeilings: ZipCeilingOverrides =
    checked !== null
      ? {
          ...(checked.data.maxRatio !== undefined ? { maxRatio: checked.data.maxRatio } : {}),
          ...(checked.data.maxUncompressedBytes !== undefined
            ? { maxUncompressedBytes: checked.data.maxUncompressedBytes }
            : {}),
        }
      : {};
  const ceilings: ZipCeilingOverrides = {
    ...(fileCeilings.maxRatio !== undefined ? { maxRatio: fileCeilings.maxRatio } : {}),
    ...(fileCeilings.maxUncompressedBytes !== undefined
      ? { maxUncompressedBytes: fileCeilings.maxUncompressedBytes }
      : {}),
    ...(flagCeilings.maxRatio !== undefined ? { maxRatio: flagCeilings.maxRatio } : {}),
    ...(flagCeilings.maxUncompressedBytes !== undefined
      ? { maxUncompressedBytes: flagCeilings.maxUncompressedBytes }
      : {}),
  };
  try {
    let working = docx;
    if (clearCount > 0) {
      try {
        working = clearDocumentBody(working);
      } catch (err: unknown) {
        return failWithCode(
          out,
          err instanceof Error ? err.message : String(err),
          'E_VALIDATION',
          undefined,
          'pick a valid .docx and retry',
          err
        );
      }
    }
    if (engineOps.length === 0) {
      // Clear-only batch: the body reset IS the batch. Gate + write.
      // R-8: flag/file ceilings scope the DocxParser zip-bomb gate.
      const cleared = await withZipCeilingEnv(ceilings, () => new DocxParser(working, ctx.macroOpts));
      if (!skipValidate) {
        const validation = await withGatedEngineOutput(out, () => cleared.validateOOXML());
        if (!validation.ok) {
          return failWithCode(
            out,
            'clearDoc applied but post-batch validation failed, nothing written.',
            'E_VALIDATION',
            undefined,
            'pick a valid .docx and retry',
            JSON.stringify(validation.errors ?? 'post-batch validation failed')
          );
        }
      }
      if (ctx.dryRun) {
        const beforeParser = await withZipCeilingEnv(ceilings, () => new DocxParser(docx, ctx.macroOpts));
        const beforeAst = await withGatedEngineOutput(out, () => beforeParser.getAST());
        const preview: DryRunPreview[] = [];
        for (const sec of beforeAst.sections ?? []) {
          for (const n of sec.nodes ?? []) {
            preview.push({ nodeId: n.id, before: n.metadata?.text ?? '', after: null });
          }
        }
        preview.push({ nodeId: '(empty paragraph)', before: null, after: '' });
        reportDryRun(out, { file, target, appliedCount: 0, cleared: clearCount, preview, normalized: previewNormalizedForDisplay(working, ctx.macroOpts), rowLabel: ctx.rowLabel });
        return EXIT_OK;
      }
      const clearBackupCode = await ensureBackupCopy(out, file, docx, target, ctx.backupState);
      if (clearBackupCode !== undefined) return clearBackupCode;
      // Display-only: the gate above already rezipped through the core
      // (healing + logging); capture that log before the save rezip below
      // replaces it, then merge both so every rewrite is reported once.
      const clearGateNormalized = skipValidate ? [] : readParserNormalizations(cleared);
      const clearBytes = cleared.rezip({ reproducible: ctx.reproducible });
      const clearNormalized = mergeNormalized(clearGateNormalized, readParserNormalizations(cleared));
      await writeDocx(target, clearBytes);
      print(out, `BATCH OK ${file} -> ${target} (applied=${clearCount} cleared=${clearCount} ${formatNormalizedInline(clearNormalized)})${rowSuffix}`);
      printJson(out, {
        ok: true,
        file,
        output: target,
        appliedCount: clearCount,
        cleared: clearCount,
        skippedIds: [],
        normalized: clearNormalized,
        ...(ctx.rowLabel !== undefined ? { row: ctx.rowLabel } : {}),
        ...(ctx.backupState.path !== undefined ? { backup: ctx.backupState.path } : {}),
        ...(ctx.reproducible ? { reproducible: true } : {}),
      });
      return EXIT_OK;
    }
    // R-8: flag/file ceilings scope the DocxParser zip-bomb gate for this load.
    const parser = await withZipCeilingEnv(ceilings, () => new DocxParser(working, ctx.macroOpts));
    const ast = await withGatedEngineOutput(out, () => parser.getAST());
    const engine = new OperationEngine(ast, parser, ctx.styleGuide !== undefined ? { styleGuide: ctx.styleGuide } : undefined);
    if (checked === null) {
      // Unreachable: clear-only batches return above (their op list is empty).
      return failWithCode(
        out,
        'E_VALIDATION: no operations to apply.',
        'E_VALIDATION',
        undefined,
        'pass a non-empty operations array; see `astra batch --schema`',
        'E_VALIDATION: no operations to apply.'
      );
    }
    const batchInput = checked.data;
    const res = await withGatedEngineOutput(out, () =>
      engine.executeBatch({
        // Strict schema already gated the payload; engine re-validates deep values.
        operations: batchInput.operations as unknown as Parameters<OperationEngine['executeBatch']>[0]['operations'],
        sequential: batchInput.sequential,
        rollbackOnError: batchInput.rollbackOnError,
        maxSteps: batchInput.maxSteps,
        dryRun: ctx.dryRun,
        ...(ctx.allowedDir !== undefined ? { allowedDir: ctx.allowedDir } : {}),
      })
    );
    if (!res.success) {
      const first = res.errors[0];
      // Flatten + clean the engine payload BEFORE composing: errors[0].error
      // is a multi-line stack, so composing first and cleaning after would cut
      // the self-correction hint off at the first stack frame. Clean parts,
      // then join into the single human line.
      const flatErr = String(first?.error ?? 'unknown').replace(/\s+/g, ' ').slice(0, 500);
      const cleanErrPart = rewriteMcpRemedies(cleanCliMessage(flatErr));
      const rawHint = first && 'hint' in first && typeof first.hint === 'string' ? first.hint : '';
      const cleanHintPart = rawHint ? rewriteMcpRemedies(cleanCliMessage(rawHint)).replace(/[.]+$/, '') : '';
      // F5-CLI: surface the backend machine code (E_QUOTA for caps, else E_VALIDATION).
      const code = inferBatchCode(first?.error ?? flatErr, 'E_VALIDATION');
      // Prefer the engine's own remedy when it carried one (authored refusal
      // text rides on the failure entry; jail + allowed-dir remedies ride in
      // the message); the generic re-inspect text stays the fallback.
      const specificRemedy = engineRemedyValue(first?.remedy) ?? engineRemedyOneLine(first?.error);
      // Hygiene: engine error payloads carry full stacks — the --json envelope
      // gets the same cleaned one-liners as the human output (never frames).
      const cleanErrors = res.errors.slice(0, 3).map((e) => ({
        operationIndex: (e as { operationIndex?: unknown }).operationIndex,
        error: rewriteMcpRemedies(cleanCliMessage(String((e as { error?: unknown }).error ?? 'unknown'))),
        hint:
          'hint' in (e as Record<string, unknown>) && typeof (e as { hint?: unknown }).hint === 'string'
            ? rewriteMcpRemedies(cleanCliMessage(String((e as { hint?: unknown }).hint)))
            : undefined,
      }));
      return failWithCode(
        out,
        cleanHintPart
          ? `batch failed (rollback=${res.rollbackTriggered}): ${cleanErrPart}. Hint: ${cleanHintPart}`
          : `batch failed (rollback=${res.rollbackTriggered}): ${cleanErrPart}.`,
        code,
        { rollbackTriggered: res.rollbackTriggered, errors: cleanErrors },
        specificRemedy ??
          're-inspect ids via `astra inspect` (or `astra search <docx> <text>` to locate text, `anchor:end` to append) and retry sequentially. Nothing was written',
        String(first?.error ?? flatErr)
      );
    }
    if (ctx.dryRun) {
      // Preview mode: the engine already applied, validated cheaply, and
      // rewound. Run the same full post-batch gate over the would-be bytes
      // so a dry-run OK stays honest, then report and write nothing.
      if (!skipValidate && res.dryRunBytes) {
        let postParser: DocxParser;
        try {
          postParser = new DocxParser(res.dryRunBytes, ctx.macroOpts);
        } catch (err: unknown) {
          return failWithCode(
            out,
            'dry run applied but the result package is unreadable, nothing written.',
            'E_VALIDATION',
            undefined,
            'pick a valid .docx and retry',
            err
          );
        }
        const validation = await withGatedEngineOutput(out, () => postParser.validateOOXML());
        if (!validation.ok) {
          const firstGate = validation.errors[0] as
            | { description?: unknown; message?: unknown; xpath?: unknown; xPath?: unknown; path?: unknown }
            | undefined;
          const firstGateDesc =
            (typeof firstGate?.description === 'string' && firstGate.description) ||
            (typeof firstGate?.message === 'string' && firstGate.message) ||
            (firstGate ? JSON.stringify(firstGate) : 'OOXML invalid');
          return failWithCode(
            out,
            `dry run: batch would apply but post-batch validation fails, nothing written: ${String(firstGateDesc).slice(0, 500)}.`,
            'E_VALIDATION',
            undefined,
            remedyForPostBatchValidation(firstGate?.xPath ?? firstGate?.xpath ?? firstGate?.path, firstGateDesc),
            JSON.stringify(validation.errors.slice(0, 3))
          );
        }
      }
      reportDryRun(out, {
        file,
        target,
        appliedCount: res.appliedCount,
        cleared: clearCount,
        preview: res.preview ?? [],
        // Would-be rewrites: the engine dry-run captured its bytes through
        // the core rezip (which logged the normalizer pass) before rewinding,
        // so the live parser log is exactly what a real save would report.
        normalized: readParserNormalizations(parser),
        rowLabel: ctx.rowLabel,
        possiblyStaleBodyRefs: Array.isArray((res as { possiblyStaleBodyRefs?: unknown }).possiblyStaleBodyRefs)
          ? ((res as { possiblyStaleBodyRefs: unknown[] }).possiblyStaleBodyRefs.filter(
              (v): v is { nodeId: string; preview: string } =>
                typeof v === 'object' &&
                v !== null &&
                typeof (v as Record<string, unknown>)['nodeId'] === 'string' &&
                typeof (v as Record<string, unknown>)['preview'] === 'string'
            ) as { nodeId: string; preview: string }[])
          : [],
      });
      return EXIT_OK;
    }
    if (!skipValidate) {
      const validation = await withGatedEngineOutput(out, () => parser.validateOOXML());
      if (!validation.ok) {
        const first = validation.errors[0] as
          | { description?: unknown; message?: unknown; xpath?: unknown; xPath?: unknown; path?: unknown }
          | undefined;
        const firstDesc =
          (typeof first?.description === 'string' && first.description) ||
          (typeof first?.message === 'string' && first.message) ||
          (first ? JSON.stringify(first) : 'OOXML invalid');
        return failWithCode(
          out,
          `batch applied but post-batch validation failed, nothing written: ${String(firstDesc).slice(0, 500)}.`,
          'E_VALIDATION',
          undefined,
          remedyForPostBatchValidation(first?.xPath ?? first?.xpath ?? first?.path, firstDesc),
          JSON.stringify(validation.errors.slice(0, 3))
        );
      }
    }
    const skipped = res.skippedIds.length > 0 ? ` skipped=${res.skippedIds.length}` : '';
    const cleared = clearCount > 0 ? ` cleared=${clearCount}` : '';
    const batchWarnings: string[] = Array.isArray((res as { warnings?: unknown }).warnings)
      ? ((res as { warnings: unknown[] }).warnings.filter((v): v is string => typeof v === 'string'))
      : [];
    const staleRefs: { nodeId: string; preview: string }[] = Array.isArray(
      (res as { possiblyStaleBodyRefs?: unknown }).possiblyStaleBodyRefs
    )
      ? ((res as { possiblyStaleBodyRefs: unknown[] }).possiblyStaleBodyRefs.filter(
          (v): v is { nodeId: string; preview: string } =>
            typeof v === 'object' &&
            v !== null &&
            typeof (v as Record<string, unknown>)['nodeId'] === 'string' &&
            typeof (v as Record<string, unknown>)['preview'] === 'string'
        ) as { nodeId: string; preview: string }[])
      : [];
    const backupCode = await ensureBackupCopy(out, file, docx, target, ctx.backupState);
    if (backupCode !== undefined) return backupCode;
    // Display-only: the post-batch gate above already rezipped through the
    // core (healing + logging); capture that log before the save rezip below
    // replaces it, then merge both so every rewrite is reported once. This
    // line never heals, only reports.
    const gateNormalized = skipValidate ? [] : readParserNormalizations(parser);
    const finalBytes = parser.rezip({ reproducible: ctx.reproducible });
    const batchNormalized = mergeNormalized(readEngineNormalized(res), gateNormalized, readParserNormalizations(parser));
    await writeDocx(target, finalBytes);
    print(out, `BATCH OK ${file} -> ${target} (applied=${res.appliedCount}${cleared}${skipped} ${formatNormalizedInline(batchNormalized)})${rowSuffix}`);
    for (const w of batchWarnings) {
      printStderr(`${CLI_NAME}: WARNING: ${w}`);
    }
    for (const s of staleRefs) print(out, `possibly-stale ${s.nodeId}: "${truncLine(s.preview, 120)}"`);
    printJson(out, {
      ok: true,
      file,
      output: target,
      appliedCount: res.appliedCount,
      cleared: clearCount,
      skippedIds: res.skippedIds,
      normalized: batchNormalized,
      ...(batchWarnings.length > 0 ? { warnings: batchWarnings } : {}),
      ...(staleRefs.length > 0 ? { possiblyStaleBodyRefs: staleRefs } : {}),
      ...(ctx.styleGuide !== undefined ? { styleGuide: summarizeStyleGuide(ctx.styleGuide) } : {}),
      ...(ctx.styleGuideWarnings !== undefined && ctx.styleGuideWarnings.length > 0 ? { styleWarnings: ctx.styleGuideWarnings } : {}),
      ...(ctx.rowLabel !== undefined ? { row: ctx.rowLabel } : {}),
      ...(ctx.backupState.path !== undefined ? { backup: ctx.backupState.path } : {}),
      ...(ctx.reproducible ? { reproducible: true } : {}),
    });
    return EXIT_OK;
  } catch (err: unknown) {
    // F5-CLI: quota throws ([E_QUOTA] caps) + validation throws both land here —
    // preserve the code in human + --json instead of bare text.
    // R-8: E_LIMIT keeps its knob remedy (observed vs ceiling), not the id hint.
    // E_SANDBOX likewise keeps its engine remedy (e.g. the $ASTRA_TMPDIR
    // symlink/owner hint, which rides on the `.remedy` prop) — same preserve
    // pattern, never the re-inspect-ids text.
    // Unreadable-container throws (raw zip/buffer/dir text from the parser
    // load) take the single normalized shape instead of leaking raw text.
    const read = classifyFileReadError(err, file);
    if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, read.remedy, err);
    const rawMsg = err instanceof Error ? err.message : String(err);
    const code = inferBatchCode(err, 'E_VALIDATION');
    const remedy =
      code === 'E_LIMIT'
        ? (extractEngineRemedy(rawMsg) ??
          'raise ASTRA_MAX_RATIO / --max-ratio or ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes above the observed value, or use a smaller document')
        : code === 'E_SANDBOX'
          ? (extractEngineRemedy(rawMsg) ?? readErrRemedyProp(err) ?? getEngineRemedy('E_SANDBOX'))
          : rawMsg.includes('--allow-macros')
            ? undefined
            : 're-inspect ids via `astra inspect` (or `astra search <docx> <text>` to locate text, `anchor:end` to append) and retry. Nothing was written';
    return failWithCode(out, rawMsg, code, undefined, remedy, err);
  }
}

// ── Streaming batches over stdin (--stream) ───────────────────────────────
// One process, one resident parser + engine. Each non-blank stdin line is one
// batch routed through the same CLI gate (op extensions, schema, quotas) and
// the same engine executeBatch (per-op validation, cheap structural scan,
// integrity check, snapshot rollback) as a single batch call. The full
// validator runs only at save time (checkpoint / EOF), never per batch.
// stdout carries exactly one JSON object per processed line plus one closing
// summary; all prose goes to stderr. The single-batch path above is
// untouched: without --stream every byte behaves as before.

export type StreamOnError = 'abort' | 'continue';

export const STREAM_DEFAULT_ON_ERROR: StreamOnError = 'abort';
export const STREAM_DEFAULT_MAX_ERRORS = 10;

export interface StreamCmdArgs {
  file: string | undefined;
  opsPath: string | undefined;
  outPath: string | undefined;
  sequential: boolean | undefined;
  skipValidate: boolean;
  flagCeilings: ZipCeilingOverrides;
  dryRun: boolean;
  backupGiven: boolean;
  backupRaw: string | undefined;
  matrixRaw: string | undefined;
  targetTextRaw: string | undefined;
  appendRaw: string | undefined;
  onErrorRaw: string | undefined;
  maxErrorsRaw: string | undefined;
  force: boolean;
  reproducible: boolean | undefined;
  macroOpts: { allowMacros: true } | { stripMacros: true } | undefined;
  allowedDir?: string | undefined;
  /** Resolved --style-guide config (one guide for the invocation; streams apply it wholesale). */
  styleGuide?: StyleGuideConfig | undefined;
  /** Parser diagnostics of the resolved --style-guide (ride the stream eof summary). */
  styleGuideWarnings?: string[] | undefined;
}

export interface StreamSessionOpts {
  file: string;
  target: string;
  sequential: boolean | undefined;
  skipValidate: boolean;
  flagCeilings: ZipCeilingOverrides;
  backupState: BackupState;
  backupSource: Buffer;
  reproducible: boolean | undefined;
  macroOpts: { allowMacros: true } | { stripMacros: true } | undefined;
  onError: StreamOnError;
  maxErrors: number;
  allowedDir?: string | undefined;
  /** Resolved --style-guide config (one guide for the whole stream). */
  styleGuide?: StyleGuideConfig | undefined;
  /** Parser diagnostics of the resolved --style-guide (ride the stream eof summary). */
  styleGuideWarnings?: string[] | undefined;
}

interface StreamBatchFailure {
  code: string;
  message: string;
  remedy?: string;
  rollbackTriggered: boolean;
}

/** Split piped stdin into lines (handles split chunks and trailing \r). */
async function* readStdinLines(): AsyncGenerator<string> {
  let buf = '';
  const flushLine = (raw: string): string => (raw.endsWith('\r') ? raw.slice(0, -1) : raw);
  for await (const chunk of process.stdin as unknown as AsyncIterable<string | Buffer>) {
    buf += typeof chunk === 'string' ? chunk : chunk.toString('utf8');
    let idx = buf.indexOf('\n');
    while (idx >= 0) {
      yield flushLine(buf.slice(0, idx));
      buf = buf.slice(idx + 1);
      idx = buf.indexOf('\n');
    }
  }
  if (buf.length > 0) yield flushLine(buf);
}

/** Clean an engine batch failure into the trimmed ack fields (hints kept). */
function engineRemedyValue(source: unknown): string | undefined {
  // A failure entry's authored `.remedy` is already the remedy text (no
  // `Remedy:` label); tolerate a labelled value from older engine builds.
  if (typeof source !== 'string') return undefined;
  const flat = source.replace(/\s+/g, ' ');
  const idx = flat.indexOf('Remedy:');
  const tail = idx >= 0 ? flat.slice(idx + 'Remedy:'.length) : flat;
  const clean = rewriteMcpRemedies(cleanCliMessage(tail)).replace(/^Remedy:\s*/i, '').replace(/[.]+$/, '').trim();
  return clean === '' ? undefined : clean;
}

function engineRemedyOneLine(source: unknown): string | undefined {
  // Backend-remedy preference: the failure's own `Remedy:` sentence when
  // present (e.g. jail + allowed-dir names on sandbox refusals), else
  // undefined for the generic re-inspect fallback. Single-line, stack-cut,
  // CLI path-hygienic (absolute paths neutralize to <path>, as everywhere).
  const flat = String(source ?? '').replace(/\s+/g, ' ');
  const idx = flat.indexOf('Remedy:');
  if (idx < 0) return undefined;
  return engineRemedyValue(flat.slice(idx));
}

/** Clean an engine batch failure into the trimmed ack fields (hints kept). */
function describeStreamEngineFailure(res: {
  errors: Array<{ error?: unknown; hint?: unknown; remedy?: unknown; operationIndex?: unknown }>;
  rollbackTriggered: boolean;
}): StreamBatchFailure {
  const first = res.errors[0] as { error?: unknown; hint?: unknown; remedy?: unknown } | undefined;
  const flatErr = String(first?.error ?? 'unknown').replace(/\s+/g, ' ').slice(0, 500);
  const cleanErrPart = rewriteMcpRemedies(cleanCliMessage(flatErr));
  const rawHint = typeof first?.hint === 'string' ? (first.hint as string) : '';
  const cleanHintPart = rawHint ? rewriteMcpRemedies(cleanCliMessage(rawHint)).replace(/[.]+$/, '') : '';
  const code = inferBatchCode(first?.error ?? flatErr, 'E_VALIDATION');
  const message = cleanHintPart
    ? `batch failed (rollback=${res.rollbackTriggered}): ${cleanErrPart}. Hint: ${cleanHintPart}`
    : `batch failed (rollback=${res.rollbackTriggered}): ${cleanErrPart}.`;
  // Same engine-remedy preference as the single-batch path (authored refusal
  // text, then jail names, over the generic re-inspect fallback).
  const specificRemedy = engineRemedyValue(first?.remedy) ?? engineRemedyOneLine(first?.error);
  return {
    code,
    message,
    remedy:
      specificRemedy ??
      're-inspect ids via `astra inspect` (or `astra search <docx> <text>` to locate text, `anchor:end` to append) and retry sequentially. Nothing was written',
    rollbackTriggered: res.rollbackTriggered,
  };
}

/**
 * Flag/arg validation for --stream plus session wiring. Pre-session failures
 * use the standard fail path (nothing streamed yet, so no acks are owed).
 */
export async function cmdBatchStream(out: GlobalOpts, args: StreamCmdArgs): Promise<number> {
  const { file, opsPath } = args;
  if (!file) return usage(out, 'batch requires <docx> <ops.json>');
  if (opsPath !== '-') {
    return usage(
      out,
      'batch --stream requires ops path `-` (NDJSON batches on stdin, e.g. `cat batches.ndjson | astra batch <docx> - --stream`)'
    );
  }
  if (args.targetTextRaw !== undefined || args.appendRaw !== undefined) {
    return usage(out, 'batch --stream cannot be combined with --target-text/--append (one resident session, stdin batches only)');
  }
  if (args.matrixRaw !== undefined) {
    return usage(out, 'batch --stream cannot be combined with --matrix (stream lines are the fan-out)');
  }
  if (args.dryRun) {
    return usage(out, 'batch --stream cannot be combined with --dry-run (checkpoints write; use a single batch --dry-run for previews)');
  }
  let onError: StreamOnError = STREAM_DEFAULT_ON_ERROR;
  if (args.onErrorRaw !== undefined) {
    if (args.onErrorRaw !== 'abort' && args.onErrorRaw !== 'continue') {
      return usage(out, 'batch --on-error requires abort|continue (default abort)');
    }
    onError = args.onErrorRaw;
  }
  let maxErrors = STREAM_DEFAULT_MAX_ERRORS;
  if (args.maxErrorsRaw !== undefined) {
    if (!/^\d+$/.test(args.maxErrorsRaw) || !Number.isSafeInteger(Number(args.maxErrorsRaw))) {
      return usage(out, 'batch --max-errors requires a non-negative integer (default 10; 0 disables the breaker)');
    }
    maxErrors = Number(args.maxErrorsRaw);
  }
  const target = (args.outPath ?? file) as string;
  if (args.backupGiven) {
    const resolved = resolveBackupPath(file, args.backupRaw);
    const absBackup = path.resolve(resolved);
    if (absBackup === path.resolve(file) || absBackup === path.resolve(target)) {
      return usage(out, `batch --backup path "${resolved}" must differ from the input and the output (refusing to overwrite either)`);
    }
  }
  if (args.outPath !== undefined && !args.force) {
    try {
      await fsp.access(target);
      return usage(out, `batch: --out target "${path.basename(target)}" already exists (refusing to overwrite without --force)`);
    } catch {
      /* destination is free — proceed */
    }
  }
  if (process.stdin.isTTY) {
    return usage(
      out,
      'batch --stream reads NDJSON batches from stdin, but stdin is a terminal (nothing piped). Remedy: pipe batches, e.g. `cat batches.ndjson | astra batch <docx> - --stream`.'
    );
  }
  let docx: Buffer;
  try {
    docx = await readDocx(file);
  } catch (err: unknown) {
    const read = classifyFileReadError(err, file);
    if (read !== undefined) return failWithCode(out, read.message, read.code, undefined, read.remedy, err);
    const rawMsg = err instanceof Error ? err.message : String(err);
    const code = /ENOENT|no such file/i.test(rawMsg) ? 'E_IO' : inferBatchCode(err, 'E_VALIDATION');
    return failWithCode(out, rawMsg, code, undefined, 'check the file paths exist and are readable, then retry', err);
  }
  const sessionOpts: StreamSessionOpts = {
    file,
    target,
    sequential: args.sequential,
    skipValidate: args.skipValidate,
    flagCeilings: args.flagCeilings,
    backupState: {
      path: args.backupGiven ? resolveBackupPath(file, args.backupRaw) : undefined,
      done: false,
    },
    backupSource: docx,
    reproducible: args.reproducible,
    macroOpts: args.macroOpts,
    onError,
    maxErrors,
    allowedDir: args.allowedDir,
    styleGuide: args.styleGuide,
    styleGuideWarnings: args.styleGuideWarnings,
  };
  const emitLine = (ack: Record<string, unknown>): void => {
    console.log(JSON.stringify(ack));
  };
  let code: number;
  try {
    code = await withZipCeilingEnv(args.flagCeilings, () =>
      runStreamSession(out, docx, sessionOpts, readStdinLines(), emitLine)
    );
  } catch (err: unknown) {
    const rawMsg = err instanceof Error ? err.message : String(err);
    return failWithCode(out, rawMsg, inferBatchCode(err, 'E_VALIDATION'), undefined, undefined, err);
  } finally {
    // Early stops (abort / breaker) leave the pipe open; destroying stdin
    // lets the process exit instead of idling on an unread pipe.
    try {
      const stdin = process.stdin as unknown as { destroy?: () => void };
      if (typeof stdin.destroy === 'function') stdin.destroy();
    } catch {
      /* best-effort stdin teardown */
    }
  }
  return code;
}

/**
 * Resident stream session over an explicit line source. Emits exactly one ack
 * per non-blank line plus one closing summary, then returns the exit code.
 * Exported so tests can drive it with array lines (no stdin mocking).
 */
export async function runStreamSession(
  out: GlobalOpts,
  docx: Buffer,
  opts: StreamSessionOpts,
  lines: Iterable<string> | AsyncIterable<string>,
  emit: (ack: Record<string, unknown>) => void
): Promise<number> {
  const { file, target, skipValidate, flagCeilings, backupState, backupSource, reproducible, macroOpts, onError, maxErrors } = opts;
  let parser = await withZipCeilingEnv(flagCeilings, () => new DocxParser(docx, macroOpts));
  const openAst = await withGatedEngineOutput(out, () => parser.getAST());
  let engine = new OperationEngine(openAst, parser, opts.styleGuide !== undefined ? { styleGuide: opts.styleGuide } : undefined);
  let batchesApplied = 0;
  let batchesFailed = 0;
  let consecutiveFailures = 0;
  let revision = 0;
  let lineIndex = -1;
  const breakerArmed = maxErrors > 0;
  const breakerTripped = (): boolean => breakerArmed && consecutiveFailures >= maxErrors;

  const stderrLine = (code: string, text: string, remedy?: string): void => {
    const cleanText = text.replace(/[.]+$/, '').trim() || text;
    printStderr(
      remedy !== undefined && remedy !== ''
        ? `${CLI_NAME}: error: [${code}] ${cleanText}. Remedy: ${remedy.replace(/[.]+$/, '')}`
        : `${CLI_NAME}: error: [${code}] ${cleanText}`
    );
  };

  /** Flush resident state to disk: full gate + backup-once + rezip + write. */
  const flushToDisk = async (): Promise<
    { ok: true } | { ok: false; code: KnownCliCode; message: string; remedy: string }
  > => {
    try {
      if (!skipValidate) {
        const validation = await withGatedEngineOutput(out, () => parser.validateOOXML());
        if (!validation.ok) {
          const first = validation.errors[0] as
            | { description?: unknown; message?: unknown; xpath?: unknown; xPath?: unknown; path?: unknown }
            | undefined;
          const firstDesc =
            (typeof first?.description === 'string' && first.description) ||
            (typeof first?.message === 'string' && first.message) ||
            (first ? JSON.stringify(first) : 'OOXML invalid');
          return {
            ok: false,
            code: 'E_VALIDATION',
            message: `stream save gated: post-batch validation failed, nothing new written: ${String(firstDesc).slice(0, 500)}.`,
            remedy: remedyForPostBatchValidation(first?.xPath ?? first?.xpath ?? first?.path, firstDesc),
          };
        }
      }
      if (backupState.path !== undefined && !backupState.done) {
        await writeDocx(backupState.path, backupSource);
        backupState.done = true;
        printStderr(`${CLI_NAME}: note: stream backup ${file} -> ${backupState.path}`);
      }
      const finalBytes = parser.rezip({ reproducible });
      await writeDocx(target, finalBytes);
      revision += 1;
      return { ok: true };
    } catch (err: unknown) {
      const rawMsg = err instanceof Error ? err.message : String(err);
      const code = inferBatchCode(err, 'E_VALIDATION');
      return {
        ok: false,
        code,
        message: rewriteMcpRemedies(cleanCliMessage(rawMsg)),
        remedy: 'check the output path is writable, then retry',
      };
    }
  };

  /** Route one batch line through the CLI gate + engine (no disk write). */
  const applyBatchLine = async (rawLine: string): Promise<{ ok: true; appliedCount: number } | StreamBatchFailure> => {
    let opsFile: OpsFile;
    try {
      opsFile = parseOpsFile(rawLine);
    } catch (err: unknown) {
      return {
        code: inferBatchCode(err, 'E_VALIDATION'),
        message: rewriteMcpRemedies(cleanCliMessage(err instanceof Error ? err.message : String(err))),
        remedy: 'pass one batch per line: a bare operations array or { "operations": [ ... ] } as JSON',
        rollbackTriggered: false,
      };
    }
    let engineOps: unknown[];
    let clearCount = 0;
    try {
      ({ ops: engineOps, clearCount } = expandCliOps(opsFile.operations));
    } catch (err: unknown) {
      return {
        code: inferBatchCode(err, 'E_VALIDATION'),
        message: rewriteMcpRemedies(cleanCliMessage(err instanceof Error ? err.message : String(err))),
        rollbackTriggered: false,
      };
    }
    const clearOnly = engineOps.length === 0 && clearCount > 0;
    let batchData:
      | { operations: unknown[]; sequential?: boolean; rollbackOnError?: boolean; maxSteps?: number; maxRatio?: number; maxUncompressedBytes?: number }
      | null = null;
    if (!clearOnly) {
      const checked = BatchOperationsSchema.safeParse({
        operations: engineOps,
        ...(opsFile.sequential !== undefined ? { sequential: opsFile.sequential } : {}),
        ...(opsFile.rollbackOnError !== undefined ? { rollbackOnError: opsFile.rollbackOnError } : {}),
        ...(opsFile.maxSteps !== undefined ? { maxSteps: opsFile.maxSteps } : {}),
        ...(opsFile.maxRatio !== undefined ? { maxRatio: opsFile.maxRatio } : {}),
        ...(opsFile.maxUncompressedBytes !== undefined ? { maxUncompressedBytes: opsFile.maxUncompressedBytes } : {}),
        ...(opts.sequential !== undefined ? { sequential: opts.sequential } : {}),
      });
      if (!checked.success) {
        const detail = checked.error.issues
          .slice(0, 3)
          .map((iss) => `${iss.path.join('.') || '(root)'}: ${iss.message}`)
          .join('; ');
        const isQuota = checked.error.issues.some(
          (iss) =>
            iss.code === 'too_big' &&
            iss.path.length === 1 &&
            (String(iss.path[0] ?? '') === 'operations' || String(iss.path[0] ?? '') === 'maxSteps')
        );
        const code: 'E_VALIDATION' | 'E_QUOTA' = isQuota ? 'E_QUOTA' : 'E_VALIDATION';
        return {
          code,
          message: `invalid operations: ${detail}.`,
          remedy:
            code === 'E_QUOTA'
              ? 'split into batches of <= 500 operations (sequential:true) with an optional positive-integer maxSteps (omit for unbounded) and retry'
              : chartTypeSchemaRemedy(detail) ?? chartHoleSizeSchemaRemedy(detail) ?? 'single edit = 1-item batch with canonical shapes',
          rollbackTriggered: false,
        };
      }
      batchData = checked.data;
    }
    const fileCeilings: ZipCeilingOverrides =
      batchData !== null
        ? {
            ...(batchData.maxRatio !== undefined ? { maxRatio: batchData.maxRatio } : {}),
            ...(batchData.maxUncompressedBytes !== undefined ? { maxUncompressedBytes: batchData.maxUncompressedBytes } : {}),
          }
        : {};
    const mergedCeilings: ZipCeilingOverrides = {
      ...(fileCeilings.maxRatio !== undefined ? { maxRatio: fileCeilings.maxRatio } : {}),
      ...(fileCeilings.maxUncompressedBytes !== undefined ? { maxUncompressedBytes: fileCeilings.maxUncompressedBytes } : {}),
      ...(flagCeilings.maxRatio !== undefined ? { maxRatio: flagCeilings.maxRatio } : {}),
      ...(flagCeilings.maxUncompressedBytes !== undefined ? { maxUncompressedBytes: flagCeilings.maxUncompressedBytes } : {}),
    };
    try {
      if (clearCount > 0) {
        // Body reset runs against current resident bytes, then the session
        // is rebuilt from the cleared package (same reset, resident form).
        const clearedBytes = clearDocumentBody(parser.rezip({ reproducible }));
        const fresh = await withZipCeilingEnv(mergedCeilings, () => new DocxParser(clearedBytes, macroOpts));
        const freshAst = await withGatedEngineOutput(out, () => fresh.getAST());
        parser = fresh;
        engine = new OperationEngine(freshAst, fresh, opts.styleGuide !== undefined ? { styleGuide: opts.styleGuide } : undefined);
      }
      if (engineOps.length === 0) {
        return { ok: true, appliedCount: clearCount };
      }
      const input = batchData as {
        operations: unknown[];
        sequential?: boolean;
        rollbackOnError?: boolean;
        maxSteps?: number;
      };
      const res = await withZipCeilingEnv(mergedCeilings, () =>
        withGatedEngineOutput(out, () =>
          engine.executeBatch({
            operations: input.operations as unknown as Parameters<OperationEngine['executeBatch']>[0]['operations'],
            sequential: input.sequential,
            rollbackOnError: input.rollbackOnError,
            maxSteps: input.maxSteps,
            ...(opts.allowedDir !== undefined ? { allowedDir: opts.allowedDir } : {}),
          })
        )
      );
      if (!res.success) return { ...describeStreamEngineFailure(res), rollbackTriggered: res.rollbackTriggered };
      return { ok: true, appliedCount: res.appliedCount + clearCount };
    } catch (err: unknown) {
      const rawMsg = err instanceof Error ? err.message : String(err);
      const code = inferBatchCode(err, 'E_VALIDATION');
      const remedy =
        code === 'E_LIMIT'
          ? (extractEngineRemedy(rawMsg) ??
            'raise ASTRA_MAX_RATIO / --max-ratio or ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes above the observed value, or use a smaller document')
          : code === 'E_SANDBOX'
            ? (extractEngineRemedy(rawMsg) ?? readErrRemedyProp(err) ?? 'check the sandbox temp dir is usable, then retry')
            : rawMsg.includes('--allow-macros')
              ? undefined
              : 're-inspect ids via `astra inspect` (or `astra search <docx> <text>` to locate text, `anchor:end` to append) and retry. Nothing was written';
      return {
        code,
        message: rewriteMcpRemedies(cleanCliMessage(rawMsg)),
        ...(remedy !== undefined ? { remedy } : {}),
        rollbackTriggered: false,
      };
    }
  };

  const emitSummary = (exitCode: number, reason?: string): void => {
    emit({
      ok: exitCode === EXIT_OK,
      control: 'eof',
      batchesApplied,
      batchesFailed,
      revision,
      exitCode,
      ...(reason !== undefined ? { reason } : {}),
      ...(opts.styleGuideWarnings !== undefined && opts.styleGuideWarnings.length > 0
        ? { styleWarnings: opts.styleGuideWarnings }
        : {}),
    });
  };

  /** Record a failed line; returns true when the session must stop reading. */
  const noteLineFailure = (index: number, failure: StreamBatchFailure): boolean => {
    batchesFailed += 1;
    consecutiveFailures += 1;
    const reason = breakerTripped() ? 'max_errors_exceeded' : undefined;
    emit({
      ok: false,
      index,
      code: failure.code,
      error: failure.message,
      ...(failure.remedy !== undefined ? { remedy: failure.remedy } : {}),
      rollbackTriggered: failure.rollbackTriggered,
      ...(reason !== undefined ? { reason, consecutiveFailures } : {}),
    });
    stderrLine(failure.code, failure.message, failure.remedy);
    return reason !== undefined || onError === 'abort';
  };

  let earlyStop: 'aborted' | 'max_errors_exceeded' | 'save_failed' | null = null;
  for await (const raw of lines) {
    if (raw.trim() === '') continue;
    lineIndex += 1;
    const index = lineIndex;
    let parsed: unknown;
    try {
      parsed = JSON.parse(raw);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      if (noteLineFailure(index, {
        code: 'E_VALIDATION',
        message: `line is not valid JSON: ${rewriteMcpRemedies(cleanCliMessage(msg))}`,
        remedy: 'pass one batch per line: a bare operations array or { "operations": [ ... ] } as JSON',
        rollbackTriggered: false,
      })) {
        earlyStop = breakerTripped() ? 'max_errors_exceeded' : 'aborted';
        break;
      }
      continue;
    }
    // Controls live in their own namespace: only {"control":"save"} is a
    // control. Anything action-shaped (including {"action":"save"}) is an
    // op batch and fails through the normal op path, never a checkpoint.
    if (parsed !== null && typeof parsed === 'object' && !Array.isArray(parsed) && 'control' in parsed) {
      const rec = parsed as Record<string, unknown>;
      if (rec['control'] !== 'save' || Object.keys(rec).length !== 1) {
        if (noteLineFailure(index, {
          code: 'E_VALIDATION',
          message: `unknown stream control ${JSON.stringify(String(rec['control']))}: only {"control":"save"} checkpoints.`,
          remedy: 'send {"control":"save"} to flush to disk and stay resident; EOF saves once and exits',
          rollbackTriggered: false,
        })) {
          earlyStop = breakerTripped() ? 'max_errors_exceeded' : 'aborted';
          break;
        }
        continue;
      }
      const saved = await flushToDisk();
      if (!saved.ok) {
        emit({ ok: false, control: 'save', index, code: saved.code, error: saved.message, remedy: saved.remedy });
        stderrLine(saved.code, saved.message, saved.remedy);
        earlyStop = 'save_failed';
        break;
      }
      emit({ ok: true, control: 'save', index, revision });
      continue;
    }
    const applied = await applyBatchLine(raw);
    if (!('ok' in applied && applied.ok === true)) {
      if (noteLineFailure(index, applied as StreamBatchFailure)) {
        earlyStop = breakerTripped() ? 'max_errors_exceeded' : 'aborted';
        break;
      }
      continue;
    }
    batchesApplied += 1;
    consecutiveFailures = 0;
    emit({ ok: true, index, appliedCount: (applied as { appliedCount: number }).appliedCount, revision });
  }

  if (earlyStop !== null) {
    // Abort / breaker / failed checkpoint: resident state is discarded by
    // simply not flushing (no panic-flush). Disk stays at the last
    // checkpoint, byte-identical to open when there was none.
    emitSummary(EXIT_FAIL, earlyStop);
    return EXIT_FAIL;
  }
  const saved = await flushToDisk();
  if (!saved.ok) {
    emit({
      ok: false,
      control: 'eof',
      code: saved.code,
      error: saved.message,
      remedy: saved.remedy,
      batchesApplied,
      batchesFailed,
      revision,
      exitCode: EXIT_FAIL,
    });
    stderrLine(saved.code, saved.message, saved.remedy);
    return EXIT_FAIL;
  }
  const exitCode = batchesFailed === 0 ? EXIT_OK : EXIT_FAIL;
  emitSummary(exitCode);
  return exitCode;
}
