import { ENGINE_ERROR_CODES, getEngineErrorCode, isEngineErrorCode, type EngineErrorCode } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_FAIL, type GlobalOpts, printStderr, stringifyJsonPayload, stripAnsi } from './output';

// ── Error hygiene (CLI DX): default human stderr is exactly one line ────────
//   `astra: error: [CODE] message. Remedy: ...`
// with no V8 stack frames, no absolute paths, no emoji. Full stacks appear
// only under `--verbose` (stderr, after the one-liner). `--json` errors carry
// the same clean one-line string plus a machine `code` field (never frames).
// MCP-surface remedy names leaking out of engine errors (`get_document_toon()`,
// `inspect_document`, `search_content`, `toon({...})`) are rewritten to their
// CLI equivalents (`astra inspect` / `astra search`). The field preflight hint
// `then retry save_document.` is rewritten to `then retry astra batch.`.

// Q0b: known codes are the central engine set (previously 7; previously-
// dropped E_SESSION/E_CONFLICT/E_NOT_FOUND/E_SECURITY/E_STATE now pass
// through verbatim instead of falling back to E_VALIDATION).
export const KNOWN_CLI_CODES: readonly EngineErrorCode[] = ENGINE_ERROR_CODES;
export type KnownCliCode = EngineErrorCode;

/** Collapse to one line and strip V8 frames / absolute paths. */
export function cleanCliMessage(raw: string): string {
  const first = String(raw ?? '')
    .split('\n')
    .map((l) => l.trim())
    .find((l) => l.length > 0 && !/^\s*at\s/.test(l) && !/^at\s/.test(l)) ?? 'unknown error';
  // Drop any trailing stack fragment glued onto the first line.
  const noStack = first.split(/\s+at\s+\S+\s*\(/)[0];
  // Drop a redundant `Error:` class prefix the engine prepends to messages.
  const noClass = noStack
    .replace(/^(Error|TypeError|RangeError):\s*/, '')
    .replace(/:\s*(Error|TypeError|RangeError):\s*/g, ': ');
  // Neutralize parenthesized absolute paths: `at X (/abs/y.js:1:2)` remnants,
  // `( /abs/path )` hints — keep only the basename.
  const noParenPaths = noClass.replace(/\((\/[^()\s]*\/[^()\s]*)\)/g, (_m, p: string) => `(${String(p).split('/').pop()})`);
  // Neutralize a leading cwd (relative display) then any other absolute path.
  const cwd = process.cwd();
  const noCwd = cwd && cwd.length > 1 ? noParenPaths.split(cwd).join('.') : noParenPaths;
  const noAbs = noCwd.replace(/(^|[\s"'`(=])\/(?:[\w.@~+-]+\/)+[\w.@~+-]*/g, '$1<path>');
  return noAbs.replace(/\s+/g, ' ').trim().slice(0, 2000) || 'unknown error';
}

/** Rewrite engine/MCP-surface remedy names to CLI-correct ones. */
export function rewriteMcpRemedies(message: string): string {
  return message
    .replace(/get_document_toon\(\)/g, 'astra inspect')
    .replace(/\bget_document_toon\b/g, 'astra inspect')
    .replace(/\binspect_document\b/g, 'astra inspect')
    .replace(/\bsearch_content\b/g, 'astra search')
    .replace(/then retry save_document\./g, 'then retry astra batch.')
    .replace(/toon\(\{[^}]*\}\)/g, 'astra inspect')
    .replace(/\btoon\(\{\}\)/g, 'astra inspect');
}

/**
 * Sanitize a verbose stack for stderr: same hygiene as the one-liner
 * (relative paths, no emoji, no absolute paths) but multi-line so `at`
 * frames survive. Used only when `--verbose` is set; default output is
 * untouched (exactly one line, never frames).
 */
export function sanitizeVerboseDetail(raw: string): string {
  let text = stripAnsi(String(raw ?? ''));
  // eslint-disable-next-line no-misleading-character-class -- intentional: strip emoji + ZWJ (U+200D) + variation selectors (U+FE00-FE0F); combined chars must match
  text = text.replace(/[\u{1F000}-\u{1FAFF}\u{2600}-\u{27BF}\u{2B00}-\u{2BFF}\u{25A0}-\u{25FF}\u{2300}-\u{23FF}\u{FE00}-\u{FE0F}\u{2190}-\u{21FF}\u{200D}]/gu, '');
  const cwd = process.cwd();
  if (cwd && cwd.length > 1) text = text.split(cwd).join('.');
  text = text.replace(/\/[^\s():]*\/packages\//g, 'packages/');
  text = text.replace(/[A-Za-z]:[\\/][^"'`()\n]*/g, '<path>');
  text = text.replace(/(^|[\s"'`(=])\/(?![A-Za-z_][\w.-]*:)(?:[\w.@~+\-: ]+\/)*[\w.@~+\-: ]+/g, '$1<path>');
  const lines = text.split('\n').map((l) => l.replace(/[ \t]+$/, '').slice(0, 2000));
  return lines.join('\n').trim().slice(0, 8000) || 'unknown error';
}

/** Classify an operational failure into { code, message, remedy }. */
export function toCliError(
  err: unknown,
  opts?: { file?: string; fallbackCode?: KnownCliCode; fallbackRemedy?: string }
): { code: KnownCliCode; message: string; remedy: string; verboseDetail?: string } {
  const rawMsg = err instanceof Error ? err.message : String(err ?? 'unknown error');
  const verboseDetail = err instanceof Error && typeof err.stack === 'string' ? err.stack : undefined;
  const msg = rewriteMcpRemedies(cleanCliMessage(rawMsg));
  const file = opts?.file;
  if (/EISDIR|illegal operation on a directory/i.test(rawMsg)) {
    return {
      code: 'E_IO',
      message: file ? cleanCliMessage(`cannot read "${file}": is a directory`) : `path is a directory: ${msg}`,
      remedy: 'pass a .docx file path (not a directory) and retry',
      verboseDetail,
    };
  }
  if (/ENOENT/i.test(rawMsg) || /no such file/i.test(rawMsg)) {
    return {
      code: 'E_IO',
      message: file ? cleanCliMessage(`cannot read "${file}": no such file or directory`) : `file not found: ${msg}`,
      remedy: 'check the path exists and is readable, then retry',
      verboseDetail,
    };
  }
  if (/Invalid or unsupported zip format|No END header|corrupt.*zip|not a valid.*(zip|docx)|ADM-ZIP/i.test(rawMsg)) {
    return {
      code: 'E_VALIDATION',
      message: file ? cleanCliMessage(`"${file}" is not a valid .docx package`) : `not a valid .docx package: ${msg}`,
      remedy: 're-export the file from Word (or pick a valid .docx) and retry',
      verboseDetail,
    };
  }
  // Byte-corrupt containers can fail inside buffer allocation with a raw
  // engine message. Map that to a clean container message naming the file.
  if (/Cannot create a Buffer|Buffer larger than|allocation failed|out of memory|heap out of memory|ERR_BUFFER_TOO_LARGE|maxOutputLength/i.test(rawMsg)) {
    return {
      code: 'E_VALIDATION',
      message: file ? cleanCliMessage(`"${file}" is not a valid .docx package (unreadable container)`) : `not a valid .docx package (unreadable container): ${msg}`,
      remedy: 'restore the file from source and retry',
      verboseDetail,
    };
  }
  // Q0b: .code prop wins over message substring (engine throws such as the
  // session-store E_SANDBOX/E_SESSION errors carry the machine code ONLY on
  // `.code`, with no E_* text in the message). Message scan preserves the
  // previously-passing behavior for code-in-message errors.
  const rawCodeProp = (err as { code?: unknown } | null | undefined)?.code;
  const propCode = isEngineErrorCode(rawCodeProp) ? rawCodeProp : undefined;
  const msgCode = KNOWN_CLI_CODES.find((c) => msg.includes(c));
  const hitCode = propCode ?? msgCode;
  if (hitCode !== undefined) {
      const code = hitCode;
      // R-8: engine E_LIMIT errors already carry `Remedy: raise
      // ASTRA_MAX_RATIO / --max-ratio ... (observed vs ceiling)`. Preserve
      // that engine remedy (single Remedy, no duplication) instead of
      // overwriting it with the generic inspect hint. The `.remedy` prop is
      // the fallback source: tmp/sandbox throws (e.g. $ASTRA_TMPDIR symlink
      // refusals) carry it ONLY on the prop, with no `Remedy:` in the message.
      const remIdx = msg.indexOf('Remedy:');
      const msgRemedy =
        remIdx >= 0 ? normalizeRemedyText(msg.slice(remIdx + 'Remedy:'.length)) : undefined;
      const engineRemedy =
        (msgRemedy !== undefined && msgRemedy !== '' ? msgRemedy : undefined) ?? readErrRemedyProp(err);
      const body = remIdx >= 0 ? msg.slice(0, remIdx).trim() : msg;
      // Strip the embedded code from the message body (failHuman re-adds it once).
      const bare = body.replace(`[${code}]`, '').replace(code, '').replace(/\s+/g, ' ').trim().replace(/^[:\-.]+/, '').trim();
      let remedy: string;
      if (engineRemedy) {
        remedy = engineRemedy;
        if (code === 'E_LIMIT' && !/ASTRA_MAX/.test(remedy)) {
          remedy += '. To allow larger inputs, raise ASTRA_MAX_RATIO / --max-ratio or ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes';
        }
      } else {
        remedy =
          opts?.fallbackRemedy ??
          (code === 'E_QUOTA'
            ? 'split into batches of <= 500 operations (sequential:true) and retry'
            : code === 'E_LIMIT'
              ? 'raise ASTRA_MAX_RATIO / --max-ratio or ASTRA_MAX_UNCOMPRESSED_BYTES / --max-uncompressed-bytes above the observed value, or use a smaller document'
              : code === 'E_DEPENDENCY'
                ? 'install the missing backend (LibreOffice for PDF/visual) and retry'
                : code === 'E_AUTH'
                  ? 'provide a valid key via --key and retry'
                  : 're-inspect ids via `astra inspect` and retry');
      }
      return { code, message: bare || body || msg, remedy, verboseDetail };
  }
  return {
    code: opts?.fallbackCode ?? 'E_VALIDATION',
    message: msg,
    remedy: opts?.fallbackRemedy ?? 're-inspect ids via `astra inspect` and retry',
    verboseDetail,
  };
}

/** Emit the canonical one-line human error (+ optional --json) and return EXIT_FAIL. */
export function failHuman(
  out: GlobalOpts,
  err: unknown,
  opts?: { file?: string; fallbackCode?: KnownCliCode; fallbackRemedy?: string; extra?: Record<string, unknown> }
): number {
  const classified = toCliError(err, opts);
  const msgSingle = classified.message.replace(/[.]+$/, '').trim() || classified.message;
  const remedySingle = classified.remedy.replace(/[.]+$/, '').trim() || classified.remedy;
  printStderr(`${CLI_NAME}: error: [${classified.code}] ${msgSingle}. Remedy: ${remedySingle}`);
  if (out.verbose && classified.verboseDetail) {
    printStderr(sanitizeVerboseDetail(classified.verboseDetail));
  }
  if (out.json) {
    console.log(
      stringifyJsonPayload({
        ok: false,
        error: `[${classified.code}] ${msgSingle}. Remedy: ${remedySingle}`,
        code: classified.code,
        ...(opts?.extra ?? {}),
      })
    );
  }
  return EXIT_FAIL;
}

export function fail(out: GlobalOpts, message: string, extra?: Record<string, unknown>): number {
  return failHuman(out, message, { fallbackCode: 'E_VALIDATION', extra });
}

/**
 * Gate engine-side console chatter (console.error/console.warn) during an
 * engine call. The engine logs raw diagnostics (stacks, absolute paths) that
 * would break CLI error hygiene; the CLI's own one-line [CODE] error carries
 * the failure instead. Both default and `--verbose` gate the raw output so
 * the canonical one-liner stays first; `--verbose` re-adds sanitized frames
 * after it via failHuman/failWithCode (relative paths, no emoji).
 */
export async function withGatedEngineOutput<T>(_out: GlobalOpts, fn: () => Promise<T>): Promise<T> {
  void _out;
  const origError = console.error;
  const origWarn = console.warn;
  console.error = () => undefined;
  console.warn = () => undefined;
  try {
    return await fn();
  } finally {
    console.error = origError;
    console.warn = origWarn;
  }
}

// ── CLI error-code helpers  ─
// Machine codes the CLI must never swallow (backend throws them).
// Q0b: full central engine set (previously 4; previously-dropped codes now
// pass through verbatim instead of returning undefined → fallback).
export type CliErrorCode = EngineErrorCode;

/** Extract the engine code from an error (`.code` prop or message prefix). */
export function extractCliErrorCode(err: unknown): CliErrorCode | undefined {
  return getEngineErrorCode(err);
}

/** Extract with fallback (strict: callers always name the expected code). */
export function resolveCliErrorCode(err: unknown, fallback: CliErrorCode): CliErrorCode {
  return extractCliErrorCode(err) ?? fallback;
}

/** Ensure a human message carries its `[CODE]` prefix plus remedy. */
export function ensureCliErrorCodePrefix(message: string, code: CliErrorCode, remedy?: string): string {
  const base = message.includes(code) ? message : `[${code}] ${message}`;
  if (remedy !== undefined && remedy !== '' && !base.includes('Remedy:')) {
    return `${base} Remedy: ${remedy}`;
  }
  return base;
}

// F5-CLI: batch failures must carry E_VALIDATION/E_QUOTA/E_LIMIT in BOTH human
// (stderr `[CODE]` prefix + remedy) and --json (`code` field) outputs.
// Backend already throws the code — the CLI must not swallow it.
export function failWithCode(
  out: GlobalOpts,
  message: string,
  code: KnownCliCode,
  extra?: Record<string, unknown>,
  remedy?: string,
  verboseDetail?: unknown
): number {
  const cleaned = rewriteMcpRemedies(cleanCliMessage(message));
  // R-8: engine E_LIMIT errors carry their own `Remedy: ...` (knob + values).
  // Split it off the body so the emitted line has exactly one [CODE] prefix
  // and one Remedy (explicit caller remedy wins, else the engine's — message
  // text first, then the `.remedy` prop for tmp/sandbox throws that carry it
  // prop-only).
  const remIdx = cleaned.indexOf('Remedy:');
  const msgRemedy =
    remIdx >= 0 ? normalizeRemedyText(cleaned.slice(remIdx + 'Remedy:'.length)) : undefined;
  const propRemedy = verboseDetail instanceof Error ? readErrRemedyProp(verboseDetail) : undefined;
  const engineRemedy =
    (msgRemedy !== undefined && msgRemedy !== '' ? msgRemedy : undefined) ?? propRemedy;
  const body = remIdx >= 0 ? cleaned.slice(0, remIdx).trim() : cleaned;
  const bare = body
    .replace(`[${code}]`, '')
    .replace(code, '')
    .replace(/\s+/g, ' ')
    .trim()
    .replace(/^[:\-.]+/, '')
    .trim();
  const text = bare || body || cleaned;
  const finalRemedy = normalizeRemedyText(remedy) ?? engineRemedy;
  // Single join: engine bodies end with '.', so strip trailing dots before
  // adding the join dot (else "(...). . Remedy:").
  const textSingle = text.replace(/[.]+$/, '').trim() || text;
  const remedySingle = finalRemedy ? (finalRemedy.replace(/[.]+$/, '').trim() || finalRemedy) : undefined;
  const line = remedySingle
    ? `${CLI_NAME}: error: [${code}] ${textSingle}. Remedy: ${remedySingle}`
    : `${CLI_NAME}: error: [${code}] ${textSingle}`;
  printStderr(line);
  if (out.verbose && verboseDetail !== undefined) {
    const raw =
      verboseDetail instanceof Error
        ? (verboseDetail.stack ?? verboseDetail.message)
        : typeof verboseDetail === 'string'
          ? verboseDetail
          : String(verboseDetail ?? '');
    if (raw.trim().length > 0) printStderr(sanitizeVerboseDetail(raw));
  }
  if (out.json) {
    const errStr = remedySingle ? `[${code}] ${textSingle}. Remedy: ${remedySingle}` : `[${code}] ${textSingle}`;
    console.log(stringifyJsonPayload({ ok: false, error: errStr, code, ...(extra ?? {}) }));
  }
  return EXIT_FAIL;
}

export function inferBatchCode(err: unknown, fallback: CliErrorCode = 'E_VALIDATION'): CliErrorCode {
  return resolveCliErrorCode(err, fallback);
}


export interface FileReadClassification {
  code: KnownCliCode;
  message: string;
  remedy: string;
}

/**
 * Normalize a file-read failure to the single missing/corrupt shape the
 * human one-liner uses (same regexes, messages, and remedies as
 * toCliError above). Batch read sites bypass toCliError via failWithCode
 * with the raw message, so they classify here first: without this the raw
 * ENOENT/ADM-ZIP/EISDIR text leaks into the one-liner with the wrong
 * remedy. Returns undefined when the error is not a read failure and the
 * caller should keep its existing code/remedy path.
 */
export function classifyFileReadError(err: unknown, file: string): FileReadClassification | undefined {
  const rawMsg = err instanceof Error ? err.message : String(err ?? 'unknown error');
  if (/EISDIR|illegal operation on a directory/i.test(rawMsg)) {
    return {
      code: 'E_IO',
      message: `cannot read "${file}": is a directory`,
      remedy: 'pass a .docx file path (not a directory) and retry',
    };
  }
  if (/ENOENT|no such file/i.test(rawMsg)) {
    return {
      code: 'E_IO',
      message: `cannot read "${file}": no such file or directory`,
      remedy: 'check the path exists and is readable, then retry',
    };
  }
  if (/Invalid or unsupported zip format|No END header|corrupt.*zip|not a valid.*(zip|docx)|ADM-ZIP/i.test(rawMsg)) {
    return {
      code: 'E_VALIDATION',
      message: `"${file}" is not a valid .docx package`,
      remedy: 're-export the file from Word (or pick a valid .docx) and retry',
    };
  }
  if (/Cannot create a Buffer|Buffer larger than|allocation failed|out of memory|heap out of memory|ERR_BUFFER_TOO_LARGE|maxOutputLength/i.test(rawMsg)) {
    return {
      code: 'E_VALIDATION',
      message: `"${file}" is not a valid .docx package (unreadable container)`,
      remedy: 'restore the file from source and retry',
    };
  }
  return undefined;
}

/**
 * Location-aware remedy for a post-batch OOXML gate failure. Mirrors the
 * wording and branch logic of the MCP save path
 * (packages/mcp-server/src/lib/mcp/diff.ts remedyForValidationError) with
 * CLI verbs (`astra inspect` / `astra batch` / retry instead of TOON /
 * execute_batch / save_document). The fallback keeps the long-standing
 * generic values hint byte-identical for shapes no branch names.
 */
export function remedyForPostBatchValidation(xPath: unknown, description: unknown): string {
  const xp = typeof xPath === 'string' ? xPath : '';
  const desc = typeof description === 'string' ? description : '';
  const xpLower = xp.toLowerCase();
  const hay = `${xpLower} ${desc.toLowerCase()}`;
  if (xpLower.includes('docpr') || (xpLower.includes('@id') && hay.includes('drawing'))) {
    return (
      'Duplicate or invalid drawing id (wp:docPr): re-inspect the flagged image/drawing node, ' +
      'ensure every drawing carries a unique non-zero docPr id (no copies sharing one id), ' +
      'fix via `astra batch`, then retry'
    );
  }
  if (
    xpLower.includes('w:tbl') ||
    xpLower.includes('/w:tbl') ||
    xpLower.includes('tblpr') ||
    xpLower.includes('tblgrid') ||
    xpLower.includes('/w:tr') ||
    xpLower.includes('/w:tc') ||
    xpLower.includes('gridcol')
  ) {
    return (
      'Table structure invalid (tblPr/tblGrid/tr/tc): re-inspect the flagged table, ' +
      'check tblGrid gridCol count matches every row cell count and children follow schema order, ' +
      'fix via `astra batch`, then retry'
    );
  }
  if (
    hay.includes('relationship') ||
    hay.includes('.rels') ||
    hay.includes('r:id') ||
    /rid\d/.test(hay)
  ) {
    return (
      'Document relationships invalid: re-inspect hyperlink/image/chart targets, ' +
      'ensure every rId referenced from the document exists in document.xml.rels with a valid target, ' +
      'fix via `astra batch`, then retry'
    );
  }
  if (hay.includes('duplicate') && (hay.includes('paraid') || hay.includes('w14:paraid'))) {
    return (
      `Duplicate paragraph id${xp ? ` (${xp})` : ''}: two paragraphs share one native w14:paraId, ` +
      'so a merge/migration cannot tell them apart. ' +
      'Delete the repeat paragraph (its session id carries a _dup suffix — see `astra inspect`) and reinsert its text for a fresh id, ' +
      'then retry'
    );
  }
  if (hay.includes('unknown element')) {
    return (
      `Non-OOXML markup${xp ? ` (${xp})` : ''}: the flagged element is not valid WordprocessingML. ` +
      'Remove or repair the element (re-export the file from Word when the package itself is suspect), ' +
      'fix via `astra batch`, then retry'
    );
  }
  if (hay.includes('data-astra-id')) {
    const part =
      xpLower.includes('w:ftr') || xpLower.includes('footer')
        ? 'footer part'
        : xpLower.includes('w:hdr') || xpLower.includes('header')
          ? 'header part'
          : 'saved part';
    return (
      `Session tracking attribute leaked into the ${part}${xp ? ` (${xp})` : ''}: ` +
      'a field or insert op stamped a session-only id the validator rejects. ' +
      'Re-run the field insertPageNumber/insertWordCount op (or headerFooter set) so the part is rebuilt without session ids, ' +
      'then retry'
    );
  }
  if (
    xpLower.includes('w:ftr') ||
    xpLower.includes('w:hdr') ||
    xpLower.includes('headerreference') ||
    xpLower.includes('footerreference') ||
    xpLower.includes('footer') ||
    xpLower.includes('header')
  ) {
    const part = xpLower.includes('ftr') || xpLower.includes('footer') ? 'footer' : 'header';
    return (
      `Header/footer part invalid${xp ? ` (${xp})` : ''}: the ${part} part failed schema validation. ` +
      'Re-inspect the header/footer content (`astra inspect`; headerFooter set rebuilds the part), ' +
      'fix via `astra batch`, then retry'
    );
  }
  if (hay.includes('invalid child element') || hay.includes('unexpected child element')) {
    return (
      `Element not valid at this position${xp ? ` (${xp})` : ''}: the validator found an element that WordprocessingML does not allow here ` +
      '(an unknown element, or a valid one in the wrong place or order). ' +
      'Remove or move the flagged element, fix via `astra batch`, then retry'
    );
  }
  return 'fix hex/half-point/twips values and retry';
}

/** Remedy text without a leading `Remedy:` label or trailing punctuation. */
export function normalizeRemedyText(raw: unknown): string | undefined {
  if (typeof raw !== 'string') return undefined;
  const text = raw.trim().replace(/^Remedy:\s*/i, '').replace(/[.]+$/, '').trim();
  return text === '' ? undefined : text;
}

/** Pull an engine `Remedy: ...` out of a raw error message, if present. */
export function extractEngineRemedy(rawMsg: string): string | undefined {
  const idx = rawMsg.indexOf('Remedy:');
  if (idx < 0) return undefined;
  return normalizeRemedyText(rawMsg.slice(idx + 'Remedy:'.length));
}

/**
 * Read an engine `.remedy` prop (tmp/sandbox throws carry the remedy ONLY on
 * the prop — e.g. `$ASTRA_TMPDIR` symlink refusals — with no `Remedy:` text
 * in the message, so message-only extraction loses it).
 */
export function readErrRemedyProp(err: unknown): string | undefined {
  const r = (err as { remedy?: unknown } | null | undefined)?.remedy;
  return normalizeRemedyText(r);
}
