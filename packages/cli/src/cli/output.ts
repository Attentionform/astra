import { DocxParser } from '@attentionform/astra-engine';
import packageJson from '../../package.json';

export const CLI_NAME = 'astra';
export const CLI_VERSION: string = (packageJson as { version?: string }).version ?? '0.1.0';

export const EXIT_OK = 0;
export const EXIT_FAIL = 1;
export const EXIT_USAGE = 2;

/** Fallback terminal width when --wrap cannot detect one (piped stdout). */
export const INSPECT_DEFAULT_WRAP_COLUMNS = 80;

/**
 * Resolve the human-presentation wrap width for inspect: an explicit
 * --width wins; --wrap auto-detects the terminal width (fallback 80 when
 * piped or unknown); neither means no wrapping (machine 1-line-per-node
 * default stays). `columns` is injectable for tests; callers omit it to
 * read process.stdout.columns.
 */
export function resolveInspectWidth(opts: {
  width?: number;
  wrap?: boolean;
  columns?: number | undefined;
}): number | undefined {
  if (opts.width !== undefined) return opts.width;
  if (!opts.wrap) return undefined;
  const detected = opts.columns ?? (typeof process.stdout?.columns === 'number' ? process.stdout.columns : undefined);
  return typeof detected === 'number' && Number.isFinite(detected) && detected > 0
    ? Math.floor(detected)
    : INSPECT_DEFAULT_WRAP_COLUMNS;
}

// ── Normalizer display surface ──────────────────────────────────────────
// Display-only: every saver renders `normalized[]` (counts + transforms) and
// dry-run lists would-be rewrites, but the rewrites themselves live in the
// engine core (allowlist normalizer inside `DocxParser.rezip()`, surfaced via
// its public `lastNormalizations` log). This file holds NO repair logic: it
// only drives the core API (`DocxParser` is imported from the engine above)
// and renders the core's own log. Reads are defensive — if the core seam is
// absent or mid-flight they fall back to an honest empty list, never a throw.
// Repair entry is the dedicated `normalize` command below — `validate` stays
// read-only fail-closed, so there is no `validate --fix` alias.
export type NormalizedRewrite = string;

/** Display-only read of the core's own normalizer log (defensive: [] when absent). */
export function readParserNormalizations(parser: unknown): NormalizedRewrite[] {
  const list = (parser as { lastNormalizations?: unknown } | null | undefined)?.lastNormalizations;
  return Array.isArray(list) ? list.filter((v): v is string => typeof v === 'string') : [];
}

/** Display-only read of an engine result's `normalized` report (defensive: [] when absent). */
export function readEngineNormalized(res: unknown): NormalizedRewrite[] {
  const list = (res as { normalized?: unknown } | null | undefined)?.normalized;
  return Array.isArray(list) ? list.filter((v): v is string => typeof v === 'string') : [];
}

/**
 * Display-only would-be rewrites for dry-run: parse + rezip a copy through
 * the core and report its log, discarding the bytes. No repair logic here —
 * the copy is thrown away, nothing is written. Parse failures propagate to
 * the caller so a corrupt input fails closed instead of reporting empty.
 */
export function previewNormalizedForDisplay(
  buf: Buffer,
  macroOpts?: { allowMacros?: boolean; stripMacros?: boolean }
): NormalizedRewrite[] {
  const probe = new DocxParser(buf, macroOpts);
  probe.rezip();
  return readParserNormalizations(probe);
}

/** Display-only repair entry: core parse + rezip (the hook heals), or passthrough on failure. */
export function applyNormalizerForDisplay(buf: Buffer): { buffer: Buffer; normalized: NormalizedRewrite[] } {
  try {
    const parser = new DocxParser(buf);
    const out = parser.rezip();
    return { buffer: out, normalized: readParserNormalizations(parser) };
  } catch {
    return { buffer: buf, normalized: [] };
  }
}

/** Inline `normalized=N [...]` fragment for result lines (bounded to 5 shown). */
export function formatNormalizedInline(list: readonly string[]): string {
  if (list.length === 0) return 'normalized=0';
  const shown = list.slice(0, 5).join('; ');
  const more = list.length > 5 ? `; +${list.length - 5} more` : '';
  return `normalized=${list.length} [${shown}${more}]`;
}

/**
 * Merge core logs across the rezips one save performs (gate rezip + save
 * rezip), dropping exact duplicates. The gate's internal rezip already heals
 * and logs, so reading only the final rezip would report a dishonest zero —
 * capture both, report every rewrite once.
 */
export function mergeNormalized(...lists: string[][]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const list of lists) {
    for (const entry of list) {
      if (!seen.has(entry)) {
        seen.add(entry);
        out.push(entry);
      }
    }
  }
  return out;
}

/**
 * T-13: ANSI escape sequence regex — CSI (with `;` and `:` sub-params for
 * colon truecolor `38:2::r:g:b` / `4:3`), OSC hyperlinks/titles (`ESC ] … BEL/ST`),
 * charset designators (`ESC ( B`), and residual single-char ESC sequences.
 * Alternation order matters: OSC first so `ESC ] …` is consumed whole instead
 * of partially matching the CSI branch.
 */
export const ANSI_REGEX = new RegExp(
  [
    '[\\u001B\\u009B]\\][^\\u0007\\u001B]*(?:\\u0007|\\u001B\\\\)',
    '[\\u001B\\u009B][()#][0-9A-Z]',
    '[\\u001B\\u009B][[\\]()#;?]*(?:\\d+(?:[;: ]\\d*)*)?[\\dA-PRZcf-nqry=><]',
    '\\u001B[@-Z\\\\-_]',
  ].join('|'),
  'g'
);

/**
 * T-13: Strip ANSI escape sequences from text.
 */
export function stripAnsi(text: string): string {
  return String(text).replace(ANSI_REGEX, '');
}

/**
 * T-13: Check if color stripping is required (!process.stdout.isTTY or
 * NO_COLOR present with a non-empty value, per https://no-color.org:
 * empty-string NO_COLOR must NOT strip).
 */
export function shouldStripColor(stream: 'stdout' | 'stderr' | { isTTY?: boolean } = 'stdout'): boolean {
  if (process.env.NO_COLOR !== undefined && process.env.NO_COLOR !== '') return true;
  if (typeof stream === 'object' && stream !== null) {
    return !stream.isTTY;
  }
  if (stream === 'stderr') {
    return !process.stderr.isTTY;
  }
  return !process.stdout.isTTY;
}

/**
 * T-13: Format message respecting TTY and NO_COLOR.
 */
export function formatCliOutput(message: string, stream: 'stdout' | 'stderr' | { isTTY?: boolean } = 'stdout'): string {
  if (shouldStripColor(stream)) {
    return stripAnsi(message);
  }
  return message;
}

/**
 * T-13: Serialize a --json payload, stripping ANSI when color is disabled
 * (non-TTY stdout or non-empty NO_COLOR). ANSI inside JSON string values
 * must not leak raw escapes into piped JSON consumers.
 */
/**
 * Recursively strip ANSI from JSON string values (arrays + plain objects).
 * JSON.stringify escapes raw ESC as `\\u001b` text, so stripping must happen
 * BEFORE serialization — a post-serialization regex would miss the escaped
 * form. Non-plain objects pass through untouched.
 */
export function stripAnsiDeep(value: unknown): unknown {
  if (typeof value === 'string') return stripAnsi(value);
  if (Array.isArray(value)) return value.map(stripAnsiDeep);
  if (value !== null && typeof value === 'object' && Object.getPrototypeOf(value) === Object.prototype) {
    return Object.fromEntries(Object.entries(value as Record<string, unknown>).map(([k, v]) => [k, stripAnsiDeep(v)]));
  }
  return value;
}

export function stringifyJsonPayload(payload: unknown, space?: string | number): string {
  const clean = shouldStripColor('stdout') ? stripAnsiDeep(payload) : payload;
  return space === undefined ? JSON.stringify(clean) : JSON.stringify(clean, null, space);
}

export type Fidelity = 'compact' | 'standard' | 'full';

export interface GlobalOpts {
  json: boolean;
  verbose: boolean;
}

export function print(out: GlobalOpts, message: string): void {
  if (!out.json) console.log(formatCliOutput(message, 'stdout'));
}

export function printStderr(message: string): void {
  console.error(formatCliOutput(message, 'stderr'));
}

export function printJson(out: GlobalOpts, payload: unknown): void {
  if (out.json) console.log(stringifyJsonPayload(payload));
}
