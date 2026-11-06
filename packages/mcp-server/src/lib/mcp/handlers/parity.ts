// Document-audit handlers: validate_document / get_stats / sanitize_document /
// normalize_document / import_document.
//
// Handler discipline follows handlers/search-batch.ts + handlers/export-audit.ts:
// strict Zod gate first (invalid wire input -> E_VALIDATION with remedy, never
// a bare SDK rejection from inside the handler), every failure returns the
// {ok:false,code,message,remedy} envelope via mcpErrorFrom (engine codes pass
// through verbatim), success envelopes carry revision + dirty, and a dropped
// redo chain is signaled explicitly (same redoCleared convention as
// execute_batch / compose_document).
//
// Engine reuse (no second implementations):
// - stats counters: computeDocStats from @attentionform/astra-engine (the same definition
//   CLI stats prints).
// - scrub: validateSanitizeOptionsStrict + sanitizeDocxBuffer from
//   @attentionform/astra-engine (the exact validator + buffer scrub CLI sanitize calls).
// - repair: the core rezip hook (the same hook CLI normalize drives through
//   applyNormalizerForDisplay).
// - ingest: handleComposeDocument (the compose tool itself), so an imported
//   file composes byte-equal to the same Markdown passed inline.
import { adoptForkNumbering } from '../../fork-numbering';
import * as fs from 'fs';
import * as path from 'path';
import {
  AdmZip,
  DocxParser,
  type DocumentAST,
  PII_COVERAGE,
  computeDocStats,
  piiCoverageSentence,
  refreshAstInPlace,
  revisionsRemedy,
  sanitizeDocxBuffer,
  toError,
  validateSanitizeOptionsStrict,
  BASELINE_MAX_ERRORS,
  judgeAgainstBaseline,
  repairDocxBuffer,
  commitStateIfDiffersFromHead,
  type BaselineVerdict,
  type ValidationResult,
} from '@attentionform/astra-engine';
import { introducedError, resolveBaseline, strictValidationDefault } from '../baseline-gate';
import {
  canonicalizeForMcp,
  collectTrackedExtrasForGit,
  getMcpAllowedRoots,
  type DocHandle,
  type McpSessionStore,
} from '../../session-store';
import { handleComposeDocument, COMPOSE_MARKDOWN_MAX } from '../../../tools/compose-document';
import { readParserNormalizations } from '../normalize-read';
import { readPendingNormalizations } from './open-inspect';
import { countPendingRevisionsInParser } from '../diff';
import { mcpErrorFrom, invalidInput } from '../errors';
import type { McpErrorShape } from '../errors';
import {
  GetStatsSchema,
  ImportDocumentSchema,
  NormalizeDocumentSchema,
  SanitizeDocumentSchema,
  ValidateDocumentSchema,
} from '../schemas';

/** Validation errors carried on success envelopes (bounded, first errors only). */
const AUDIT_VALIDATION_ERRORS_MAX = 10;

/** Import file cap (bytes): single source is the compose Markdown cap. */
const IMPORT_FILE_MAX_BYTES = COMPOSE_MARKDOWN_MAX;

export interface AuditValidation {
  valid: boolean;
  errorCount: number;
  errors: Array<{ description: string; xPath?: string }>;
}

function toAuditValidation(
  validation: { ok: boolean; errors?: unknown },
  max: number = AUDIT_VALIDATION_ERRORS_MAX
): AuditValidation {
  const rawErrors = Array.isArray(validation.errors) ? validation.errors.slice(0, max) : [];
  const errors = rawErrors.map((e) => {
    const err = e as { description?: unknown; message?: unknown; xPath?: unknown; path?: unknown };
    const description = String(
      (typeof err.description === 'string' && err.description) ||
        (typeof err.message === 'string' && err.message) ||
        'OOXML validation error'
    ).slice(0, 500);
    const xp =
      typeof err.xPath === 'string' ? err.xPath : typeof err.path === 'string' ? err.path : undefined;
    return xp ? { description, xPath: xp } : { description };
  });
  return {
    valid: validation.ok,
    errorCount: Array.isArray(validation.errors) ? validation.errors.length : 0,
    errors,
  };
}

/**
 * Refresh the session AST in place after the package bytes changed under it
 * (sanitize / normalize replace entries the live DOM was built from). Same
 * object the session engine holds (sections spliced, maps re-filled), so
 * later batches keep working on the identical handle. Refresh recipe mirrors
 * the compose tool (tools/compose-document.ts): fresh parse, splice in place.
 */
async function refreshSessionAst(handle: DocHandle): Promise<void> {
  await refreshAstInPlace(handle.ast, handle.parser);
}

/**
 * Serialize the live session through an isolated parser copy (fork), so
 * read-only audits never disturb live session state: the core rezip heals
 * the DOM it runs on and records what it did on the parser, which would
 * silently heal the session and misattribute later save reports. Fork recipe
 * mirrors the export tool (tools/export.ts serializeSessionCopy).
 */
function serializeSessionFork(parser: DocxParser, ast: DocumentAST): Buffer {
  const snap = parser.capturePackageSnapshot();
  try {
    const stylesXml = parser.getStructuralSummary().stylesXml;
    if (typeof stylesXml === 'string' && stylesXml.length > 0) {
      snap.set('word/styles.xml', Buffer.from(stylesXml, 'utf8'));
    }
  } catch {
    // Keep the snapshot styles bytes when the live read is unavailable.
  }
  const zip = new AdmZip();
  for (const [name, buf] of snap) {
    zip.addFile(name, Buffer.from(buf));
  }
  const staleBytes = zip.toBuffer();
  // The snapshot already carries the live document.xml serialization, so the
  // stale bytes validate correctly on their own. The fork rezip below only
  // adds numbering fidelity — when the package trips a content guard
  // (duplicate ids), the rezip throws before the validator can report it, so
  // fall back to the stale bytes and let validateOOXML map the guard itself.
  try {
    const fork = new DocxParser(staleBytes);
    try {
      adoptForkNumbering(fork, ast);
    } catch {
      // Fall through: a bare fork rezip preserves the snapshot as-is.
    }
    return fork.rezip();
  } catch {
    return staleBytes;
  }
}

/** Read-only OOXML verdict over the live session (forked, never writes). */
async function forkValidation(handle: DocHandle): Promise<{ validation: AuditValidation; buf: Buffer }> {
  const buf = serializeSessionFork(handle.parser, handle.ast);
  const fork = new DocxParser(buf);
  const validation = await fork.validateOOXML();
  return { validation: toAuditValidation(validation), buf };
}

/** Same fork, with the full error list so the result can be compared with the baseline. */
async function forkValidationFull(handle: DocHandle): Promise<ValidationResult> {
  const buf = serializeSessionFork(handle.parser, handle.ast);
  const fork = new DocxParser(buf);
  return fork.validateOOXML({ maxErrors: BASELINE_MAX_ERRORS });
}

/**
 * validate_document: strictly read-only preflight ISO audit. The verdict runs
 * on an isolated fork (see serializeSessionFork): revision, dirty flag,
 * normalizer log, disk bytes, and history are all untouched. Never throws —
 * every failure maps to {ok:false,code,remedy}.
 */
export async function handleValidateDocument(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      revision: number;
      dirty: boolean;
      validation: AuditValidation;
      /** Repairs a save would make (e.g. children put in schema order); absent when none. */
      pendingNormalizations?: string[];
      /** Advice when normalize_document can repair errors in the document. */
      normalizeHint?: string;
    }
  | McpErrorShape
> {
  const parsed = ValidateDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`validate_document: invalid input. ${detail}`);
  }
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      const { validation } = await forkValidation(handle);
      // The verdict judges the document as a save would write it, so say what
      // that save would repair instead of leaving a plain ok.
      const pendingNormalizations = readPendingNormalizations(handle);
      const normalizeHint = validation.valid
        ? undefined
        : 'Call `normalize_document` to attempt the supported repairs, then call `validate_document` again.';
      return {
        ok: true as const,
        sessionId: handle.sessionId,
        revision: handle.revision,
        dirty: store.isDirty(handle),
        validation,
        ...(pendingNormalizations.length > 0 ? { pendingNormalizations } : {}),
        ...(normalizeHint ? { normalizeHint } : {}),
      };
    });
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_SESSION', 'Call open_document first, then retry validate_document with the returned sessionId.');
  }
}

/**
 * get_stats: token-cheap census over the live AST (pure read: no revision
 * bump, no dirty flip). Counters come from computeDocStats in @attentionform/astra-engine —
 * the same definition CLI stats prints — so the two surfaces cannot drift.
 * Never throws.
 */
export async function handleGetStats(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      revision: number;
      dirty: boolean;
      words: number;
      paragraphs: number;
      tables: number;
      images: number;
      pages: number;
    }
  | McpErrorShape
> {
  const parsed = GetStatsSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`get_stats: invalid input. ${detail}`);
  }
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      const stats = computeDocStats(handle.ast);
      return {
        ok: true as const,
        sessionId: handle.sessionId,
        revision: handle.revision,
        dirty: store.isDirty(handle),
        words: stats.words,
        paragraphs: stats.paragraphs,
        tables: stats.tables,
        images: stats.images,
        pages: stats.pages,
      };
    });
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_SESSION', 'Call open_document first, then retry get_stats with the returned sessionId.');
  }
}

export interface SanitizeDocumentOk {
  ok: true;
  sessionId: string;
  revision: number;
  dirty: boolean;
  removedRSIDs: number;
  purgedAuthors: number;
  strippedExifBytes: number;
  purgedPII: number;
  bodyPII: number;
  piiCoverage: typeof PII_COVERAGE;
  piiNote: string;
  acceptedRevisions: number;
  strippedComments: number;
  strippedMacros: string[];
  macroParts: string[];
  filesTouched: string[];
  droppedRogue: string[];
  revisionsPending: number;
  revisionsRemedy?: string;
  redoCleared?: boolean;
  redoClearedWarning?: string;
}

/**
 * sanitize_document: scrub the open session through the engine scrub path
 * (validateSanitizeOptionsStrict + sanitizeDocxBuffer — the exact validator
 * and buffer scrub CLI sanitize calls). The sanitized bytes replace the live
 * package (snapshot restore on the same parser, AST refreshed in place), the
 * revision bumps, and the redo chain clears (a fresh edit invalidates redo).
 * Never throws.
 */
export async function handleSanitizeDocument(
  store: McpSessionStore,
  input: unknown
): Promise<SanitizeDocumentOk | McpErrorShape> {
  const parsed = SanitizeDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`sanitize_document: invalid input. ${detail}`);
  }
  const { sessionId, ...scrubInput } = parsed.data;
  let options: ReturnType<typeof validateSanitizeOptionsStrict>['options'];
  try {
    options = validateSanitizeOptionsStrict(scrubInput).options;
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_VALIDATION', 'Retry sanitize_document with a strict subset of the documented scrub options.');
  }
  try {
    return await store.withSession(sessionId, async (handle) => {
      // Flush live edits through the core hook first (same bytes a save
      // would persist), then scrub the buffer through the engine path.
      const live = handle.parser.rezip();
      const { buffer, report } = sanitizeDocxBuffer(live, options);
      const restored = new Map<string, Buffer>();
      const snap = new AdmZip(buffer);
      for (const entry of snap.getEntries()) {
        if (entry.isDirectory) continue;
        restored.set(entry.entryName, Buffer.from(entry.getData()));
      }
      handle.parser.restorePackageSnapshot(restored);
      await refreshSessionAst(handle);
      const revision = store.bumpRevision(handle);
      store.refreshByteSize(handle);
      const redoClearedDepth = handle.redoStack.length;
      handle.redoStack.length = 0;
      const revisionsPending = countPendingRevisionsInParser(handle.parser);
      const pendingRemedy = revisionsRemedy(revisionsPending);
      return {
        ok: true as const,
        sessionId: handle.sessionId,
        revision,
        dirty: store.isDirty(handle),
        removedRSIDs: report.removedRSIDs,
        purgedAuthors: report.purgedAuthors,
        strippedExifBytes: report.strippedExifBytes,
        purgedPII: report.purgedPII,
        bodyPII: report.bodyPII,
        piiCoverage: PII_COVERAGE,
        piiNote: piiCoverageSentence(),
        acceptedRevisions: report.acceptedRevisions,
        strippedComments: report.strippedComments,
        strippedMacros: [...report.strippedMacros],
        macroParts: [...report.macroParts],
        filesTouched: [...report.filesTouched],
        droppedRogue: [...report.droppedRogue],
        revisionsPending,
        ...(pendingRemedy ? { revisionsRemedy: pendingRemedy } : {}),
        ...(redoClearedDepth > 0
          ? {
              redoCleared: true as const,
              redoClearedWarning:
                `This sanitize cleared ${redoClearedDepth} redo entr${redoClearedDepth === 1 ? 'y' : 'ies'}: ` +
                'a fresh edit invalidates redo, so redo_document now reports an empty stack. ' +
                'Call revert_document to record a fresh redoable state.',
            }
          : {}),
      };
    });
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_VALIDATION', 'Retry sanitize_document; re-open the document if the session expired.');
  }
}

/**
 * normalize_document: canonical repair of the open session through the core
 * rezip hook (the hook CLI normalize drives). Heals the live DOM, refreshes
 * the session view, then proves the result with a forked validation verdict.
 * When the package is still invalid afterwards, the pre-run snapshot is
 * restored (revision untouched) and the call fails closed with E_VALIDATION —
 * the in-memory analogue of CLI normalize writing nothing on failure. Output
 * bytes match CLI normalize on the same input. Never throws.
 */
export async function handleNormalizeDocument(
  store: McpSessionStore,
  input: unknown
): Promise<
  | {
      ok: true;
      sessionId: string;
      revision: number;
      dirty: boolean;
      normalized: string[];
      preStateCommit?: string;
      remaining: number;
      rolledBack: { rule: string; reason: string }[];
      validation: AuditValidation & { tolerated?: true; preExistingErrors?: number; introducedErrors?: 0 };
      revisionsPending: number;
      revisionsRemedy?: string;
      redoCleared?: boolean;
      redoClearedWarning?: string;
    }
  | McpErrorShape
> {
  const parsed = NormalizeDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return invalidInput(`normalize_document: invalid input. ${detail}`);
  }
  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      const before = handle.parser.capturePackageSnapshot();
      const beforeRevision = handle.revision;
      let normalized: string[];
      // The state before the change, for the history: one revert returns to it.
      const preXml = handle.parser.readDocumentXml();
      const preExtras = collectTrackedExtrasForGit(handle.parser);
      let remaining = -1;
      let rolledBack: { rule: string; reason: string }[] = [];
      try {
        const rezipped = handle.parser.rezip();
        normalized = readParserNormalizations(handle.parser);
        // Repair rules run on the rezipped bytes; each kept rule adds no validator error and keeps the visible text.
        const repair = await repairDocxBuffer(rezipped);
        remaining = repair.remaining;
        rolledBack = repair.rolledBack;
        if (repair.normalized.length > 0) {
          const repairedParts = new Map<string, Buffer>();
          for (const entry of new AdmZip(repair.buffer).getEntries()) {
            if (!entry.isDirectory) repairedParts.set(entry.entryName, Buffer.from(entry.getData()));
          }
          handle.parser.restorePackageSnapshot(repairedParts);
          normalized = [...normalized, ...repair.normalized];
        }
      } catch (err: unknown) {
        try {
          handle.parser.restorePackageSnapshot(before);
          await refreshSessionAst(handle);
        } catch {
          /* restore is best-effort; the original throw carries the verdict */
        }
        throw err;
      }
      await refreshSessionAst(handle);
      const fullResult = await forkValidationFull(handle);
      const verdict: BaselineVerdict = fullResult.ok
        ? { kind: 'clean' }
        : strictValidationDefault()
          ? { kind: 'incomparable', reason: 'no-baseline' }
          : judgeAgainstBaseline(fullResult, await resolveBaseline(handle));
      const validation = toAuditValidation(fullResult);
      if (!validation.valid && verdict.kind !== 'tolerated') {
        try {
          handle.parser.restorePackageSnapshot(before);
          await refreshSessionAst(handle);
        } catch {
          /* restore is best-effort; the refusal below carries the verdict */
        }
        if (verdict.kind === 'introduced') throw introducedError('normalize_document', verdict);
        const first = validation.errors[0];
        const desc = first ? first.description : 'OOXML invalid';
        throw toError(
          'E_VALIDATION',
          `normalize applied but validation still fails, session restored (revision ${beforeRevision}): ${String(desc).slice(0, 500)}.`,
          'Fix the flagged content in Word and retry (nothing was kept).'
        );
      }
      // With a history repository and a change made, record the state before it (when it differs from the head).
      // The next save_document records the result, so one revert_document returns to this state.
      const preStateCommit =
        normalized.length > 0
          ? await commitStateIfDiffersFromHead(handle.filePath, preXml, 'normalize_document: state before repairs', preExtras)
          : null;
      const revision = store.bumpRevision(handle);
      store.refreshByteSize(handle);
      const redoClearedDepth = handle.redoStack.length;
      handle.redoStack.length = 0;
      const revisionsPending = countPendingRevisionsInParser(handle.parser);
      const pendingRemedy = revisionsRemedy(revisionsPending);
      return {
        ok: true as const,
        sessionId: handle.sessionId,
        revision,
        dirty: store.isDirty(handle),
        normalized,
        ...(preStateCommit ? { preStateCommit: preStateCommit.slice(0, 7) } : {}),
        remaining,
        rolledBack,
        validation:
          verdict.kind === 'tolerated'
            ? { ...validation, tolerated: true as const, preExistingErrors: verdict.preExisting, introducedErrors: 0 as const }
            : validation,
        revisionsPending,
        ...(pendingRemedy ? { revisionsRemedy: pendingRemedy } : {}),
        ...(redoClearedDepth > 0
          ? {
              redoCleared: true as const,
              redoClearedWarning:
                `This normalize cleared ${redoClearedDepth} redo entr${redoClearedDepth === 1 ? 'y' : 'ies'}: ` +
                'a fresh edit invalidates redo, so redo_document now reports an empty stack. ' +
                'Call revert_document to record a fresh redoable state.',
            }
          : {}),
      };
    });
  } catch (err: unknown) {
    return mcpErrorFrom(err, 'E_VALIDATION', 'Pick a valid document and retry normalize_document (nothing was kept).');
  }
}

/**
 * Jailed Markdown read for import_document. Verdict table mirrors the
 * batch image-file read precedent: outside the workspace roots -> E_SANDBOX,
 * missing/unreadable -> E_IO, oversize -> E_LIMIT, unusable -> E_VALIDATION.
 * Symlinks are refused at open (O_NOFOLLOW) with containment re-checked after
 * the read, so a swapped link cannot escape the jail mid-call.
 */
export function readImportMarkdownFile(rawPath: unknown, roots: string[] = getMcpAllowedRoots()): string {
  if (typeof rawPath !== 'string' || !rawPath.trim()) {
    throw toError('E_SANDBOX', 'import path must be a non-empty string', 'Pass a .md path inside documents/ or the system temp dir.');
  }
  if (rawPath.includes('\0') || /%00/i.test(rawPath)) {
    throw toError('E_SANDBOX', 'import path contains a null byte', 'Remove null bytes from the path and retry.');
  }
  if (rawPath.includes('\n') || rawPath.includes('\r')) {
    throw toError('E_SANDBOX', 'import path contains forbidden line break', 'Remove line breaks from the path and retry.');
  }
  const trimmed = rawPath.trim();
  const segments = trimmed.split(/[\\/]/);
  if (segments.includes('..')) {
    throw toError('E_SANDBOX', 'import path contains forbidden ".." traversal', 'Use a path inside the workspace documents directory.');
  }
  const resolved = path.resolve(trimmed);
  const canonical = canonicalizeForMcp(resolved);
  const inside = roots.some((root) => {
    let canonicalRoot = root;
    try {
      canonicalRoot = fs.realpathSync(root);
    } catch {
      /* root may not exist — use lexical form */
    }
    const normC = path.normalize(canonical);
    const normR = path.normalize(canonicalRoot);
    return normC === normR || normC.startsWith(normR + path.sep);
  });
  if (!inside) {
    throw toError(
      'E_SANDBOX',
      `import path is outside the allowed directories (${roots.join(', ')})`,
      'Place the Markdown file inside one of the allowed directories, or name its folder with --allowed-dir or ASTRA_EXTRA_DOC_ROOTS.'
    );
  }
  const base = path.basename(canonical);
  const nofollow = fs.constants.O_NOFOLLOW as number | undefined;
  const flags =
    typeof nofollow === 'number' && nofollow !== 0
      ? (fs.constants.O_RDONLY as number) | nofollow
      : 'r';
  let fd: number | undefined;
  try {
    fd = fs.openSync(canonical, flags);
  } catch (err: unknown) {
    const code = (err as { code?: string } | undefined)?.code;
    if (code === 'ENOENT') {
      throw toError('E_IO', `import file is missing (${base})`, 'Check the Markdown path exists and is readable, then retry.');
    }
    if (code === 'ELOOP') {
      throw toError('E_SANDBOX', `import path traverses a symlink (${base})`, 'Point the import at a regular Markdown file inside the allowed directory.');
    }
    throw toError('E_IO', `import file cannot be read (${base})`, 'Check the file exists and is readable, then retry.');
  }
  try {
    const st = fs.fstatSync(fd as number);
    if (!st.isFile()) {
      throw toError('E_IO', `import reference is not a file (${base})`, 'Point the import at a regular Markdown file.');
    }
    if (st.size <= 0) {
      throw toError('E_VALIDATION', `import file is empty (${base})`, 'Use a non-empty Markdown file.');
    }
    if (st.size > IMPORT_FILE_MAX_BYTES) {
      throw toError(
        'E_LIMIT',
        `import file is too large (${base}, ${st.size} bytes > ${IMPORT_FILE_MAX_BYTES} cap)`,
        'Split the Markdown source into smaller files and import each in its own call.'
      );
    }
    const buf = fs.readFileSync(fd as number);
    if (buf.length > IMPORT_FILE_MAX_BYTES) {
      throw toError(
        'E_LIMIT',
        `import file is too large (${base}, ${buf.length} bytes > ${IMPORT_FILE_MAX_BYTES} cap)`,
        'Split the Markdown source into smaller files and import each in its own call.'
      );
    }
    let post = canonical;
    try {
      post = fs.realpathSync(canonical);
    } catch {
      /* target vanished mid-read — pre-open verdict stands */
    }
    const stillInside = roots.some((root) => {
      let canonicalRoot = root;
      try {
        canonicalRoot = fs.realpathSync(root);
      } catch {
        /* lexical form */
      }
      const normC = path.normalize(post);
      const normR = path.normalize(canonicalRoot);
      return normC === normR || normC.startsWith(normR + path.sep);
    });
    if (!stillInside) {
      throw toError('E_SANDBOX', 'import path points outside the allowed directory', 'Keep Markdown files inside documents/ or the system temp dir.');
    }
    return buf.toString('utf8');
  } finally {
    try {
      if (fd !== undefined) fs.closeSync(fd);
    } catch {
      /* close must never mask the read verdict */
    }
  }
}

/**
 * import_document: jailed file read + the existing compose pipeline, ingested
 * into the OPEN session only (sessionId strictly required, never
 * auto-created). Format table mirrors CLI import: .md/.markdown/.txt read as
 * Markdown; .html/.htm refused with the CLI-identical verdict; .docx refused
 * naming open_document; anything else refused as an unknown format. After the
 * gates, the call delegates to handleComposeDocument itself, so the result is
 * byte-equal to composing the same Markdown inline. Never throws.
 */
export async function handleImportDocument(
  store: McpSessionStore,
  input: unknown
): Promise<Awaited<ReturnType<typeof handleComposeDocument>>> {
  const parsed = ImportDocumentSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return {
      ok: false as const,
      code: 'E_VALIDATION',
      message: `import_document: invalid input. ${detail}`,
      remedy: 'Retry with {"sessionId":"<id from open_document>","path":"<allowed-dir>/file.md"} (optional allowedDir for image refs, optional template report|memo|blank, optional styleGuide).',
    };
  }
  const srcExt = path.extname(parsed.data.path).toLowerCase();
  if (srcExt === '.html' || srcExt === '.htm') {
    return {
      ok: false as const,
      code: 'E_VALIDATION',
      message: `HTML import is not supported directly ("${path.basename(parsed.data.path)}")`,
      remedy: 'ASTra imports Markdown (.md). Convert the HTML file to Markdown or supply a .md document',
    };
  }
  if (srcExt === '.docx') {
    return {
      ok: false as const,
      code: 'E_VALIDATION',
      message: `import_document cannot ingest .docx files ("${path.basename(parsed.data.path)}"): open .docx files directly with open_document instead`,
      remedy: 'Call open_document with the .docx path to work with it, or convert it to Markdown first and import the .md file.',
    };
  }
  if (srcExt !== '.md' && srcExt !== '.markdown' && srcExt !== '.txt') {
    return {
      ok: false as const,
      code: 'E_VALIDATION',
      message: `import_document: unsupported format ("${srcExt || '(none)'}"): supply a Markdown file`,
      remedy: 'Pass a .md (or .markdown/.txt) file; .html converts to Markdown first, .docx opens via open_document.',
    };
  }
  let markdown: string;
  try {
    markdown = readImportMarkdownFile(parsed.data.path);
  } catch (err: unknown) {
    const code = (err as { code?: string })?.code;
    const remedy = (err as { remedy?: string })?.remedy;
    return {
      ok: false as const,
      code: code === 'E_SANDBOX' || code === 'E_IO' || code === 'E_LIMIT' || code === 'E_VALIDATION' ? code : 'E_IO',
      message: err instanceof Error ? err.message : String(err),
      remedy: typeof remedy === 'string' ? remedy : 'Check the Markdown path and retry import_document.',
    };
  }
  return handleComposeDocument(store, {
    sessionId: parsed.data.sessionId,
    markdown,
    ...(parsed.data.allowedDir !== undefined ? { allowedDir: parsed.data.allowedDir } : {}),
    ...(parsed.data.template !== undefined ? { template: parsed.data.template } : {}),
    ...(parsed.data.styleGuide !== undefined ? { styleGuide: parsed.data.styleGuide } : {}),
    ...(parsed.data.dollarMath ? { dollarMath: true } : {}),
  });
}
