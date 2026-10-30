import * as fs from 'fs';
import * as path from 'path';
import { z } from 'zod';
import {
  compileMarkdown,
  composeDocument,
  getEngineErrorCode,
  isFileGuideRef,
  refreshAstInPlace,
  resolveStyleGuide,
  revisionsRemedy,
  summarizeStyleGuide,
  validateOperationsStrict,
  type EngineErrorCode,
  type ResolvedStyleGuide,
  type StyleGuideConfig,
  type StyleGuideSummary,
} from '@attentionform/astra-engine';
import {
  canonicalizeForMcp,
  getMcpAllowedRoots,
  type DocHandle,
  type McpSessionStore,
} from '../lib/session-store';
import {
  attributeCreatedIds,
  countPendingRevisionsInParser,
  mergeNormalized,
  readEngineNormalized,
  readParserNormalizations,
  rewritePhantomMcpToolNames,
  sanitizeMcpClientMessage,
  type BatchOpCreatedIds,
} from '../lib/mcp-tool-adapter';

/** Upper bound on the Markdown source per call (chars). Larger sources should be split by the caller. */
export const COMPOSE_MARKDOWN_MAX = 500000;
/** Validation errors carried on the success envelope (bounded, first errors only). */
const COMPOSE_VALIDATION_ERRORS_MAX = 3;

/** True for refs carrying a URI scheme (`https://…`, `data:…`) — never a local guide. */
function hasGuideScheme(ref: string): boolean {
  return /^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(ref.trim());
}

function guideSandboxError(message: string, remedy: string): Error & { code: string; remedy: string } {
  return Object.assign(new Error(message), { code: 'E_SANDBOX', remedy });
}

export interface McpGuideJail {
  /**
   * Explicit jail directory (already validated against the MCP roots, e.g.
   * the batch/compose `allowedDir`). Wins over the session anchor.
   */
  allowedDir?: string;
  /**
   * Session filesystem anchor (absolute open-document path): file guide
   * refs resolve under its directory (same trust as batch image reads).
   */
  sessionFilePath?: string;
}

/**
 * Resolves a `styleGuide` ref under the MCP jail. Named guides resolve to
 * bundled files (unknown names fail closed E_VALIDATION naming the
 * available set); file refs resolve under the explicit `allowedDir` else
 * the session anchor directory; remote URLs fail closed. A file-backed ref
 * with no anchor (in-memory session, no allowedDir) fails E_SANDBOX —
 * same rule as batch image `filePath` (never silently cwd-anchored).
 */
export function resolveMcpStyleGuide(ref: string, jail?: McpGuideJail): ResolvedStyleGuide {
  const trimmed = typeof ref === 'string' ? ref.trim() : '';
  let anchor: string | undefined;
  if (jail?.allowedDir !== undefined) {
    anchor = jail.allowedDir;
  } else if (jail?.sessionFilePath !== undefined && path.isAbsolute(jail.sessionFilePath)) {
    anchor = canonicalizeForMcp(path.resolve(path.dirname(jail.sessionFilePath)));
  }
  if (anchor === undefined && !hasGuideScheme(trimmed) && isFileGuideRef(trimmed)) {
    throw guideSandboxError(
      `style guide '${trimmed}' is a file path but the session has no filesystem anchor (in-memory document).`,
      'pass allowedDir pointing at the folder holding the guide file and retry.'
    );
  }
  // Remote URLs route through the engine resolver (fail-closed E_VALIDATION
  // with a download-into-workspace remedy); named guides resolve bundled;
  // file refs stay jailed under the anchor above.
  return resolveStyleGuide(trimmed, anchor !== undefined ? { allowedDir: anchor } : undefined);
}

/**
 * Attaches a resolved guide to a session (persists for the session,
 * mirrored onto the session engine so later batches inherit it).
 */
export function attachSessionStyleGuide(
  handle: DocHandle,
  ref: string,
  resolved: ResolvedStyleGuide
): StyleGuideConfig {
  handle.styleGuide = resolved.config;
  handle.styleGuideRef = ref.trim();
  handle.engine.setStyleGuide(resolved.config);
  return resolved.config;
}

export const ComposeDocumentSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    markdown: z.string().min(1).max(COMPOSE_MARKDOWN_MAX),
    allowedDir: z.string().min(1).max(1024).optional(),
    template: z.enum(['report', 'memo', 'blank']).optional(),
    // Style guide attach (bundled name or guide file path). Persists for
    // the session and skins this compose; batches inherit it afterwards.
    // Empty is rejected in the handler with a detach hint — no min(1) here
    // on purpose so the hint (not a bare schema error) reaches the caller.
    styleGuide: z.string().max(1024).optional(),
    // Read $...$ as inline math (off by default: prices and shell text can look like math).
    dollarMath: z.boolean().optional(),
  })
  .strict();

export type ComposeDocumentInput = z.infer<typeof ComposeDocumentSchema>;

export const COMPOSE_DOCUMENT_DESCRIPTION =
  'Compose a Markdown string into the open session document (appends after existing content, except that a pristine blank template has its two seed placeholders replaced so output holds only composed content). ' +
  'Runs the engine compose path, then applies the resulting operations through the same execute_batch path ' +
  '(strict schema gate, per-chunk quotas, structural checks, per-chunk atomic rollback, then inline-span run splitting), so every gate a batch meets applies here unchanged. ' +
  'Supported input: #/##/### headings (real Heading1..3 styles), blank-line paragraphs (single newlines join with a space), ' +
  'flat bullet (-, *, +) and ordered (1., 1)) lists with real numbering, GitHub-style tables with a header separator row ' +
  '(header row shaded and repeated), --- dividers, fenced or indented code blocks (monospace, line breaks kept), ' +
  '[^label] footnotes with definitions, inline spans (**bold**, *italic*, `code`, [text](url)) applied as runs with real hyperlink targets, ' +
  'and standalone block images/charts (![alt](path)) resolved from allowedDir. Out-of-scope input fails closed with E_VALIDATION plus a 1-based line number and writes nothing: ' +
  'raw HTML (except [!NOTE]-style alert callouts, which compile to callout containers), custom styles, general blockquotes, deeper headings, nested lists or tables, task items. ' +
  'Block image/chart refs need allowedDir pointing at the folder holding the files (refs stay jailed inside it; without allowedDir they fail closed); ' +
  'remote URL refs fail closed; inline data-URL block images bypass allowedDir by design (bytes decode in memory through the same image gates). ' +
  'Sources compiling past 500 operations apply as sequential 500-operation chunks; each chunk is atomic, but a later-chunk ' +
  'failure leaves earlier chunks applied (same exposure as any multi-batch caller). ' +
  'Compiled node ids are deterministic per source, so a second compose into the same session is rejected by the duplicate-id gate ' +
  '(compose once per session with the full source). Optional template (report|memo|blank) requires the open document to still hold ' +
  'that template seed, otherwise composition fails before any write ' +
  '(for a fresh templated document call new_document first, then compose_document with template), and placement is append after existing content except for the pristine blank replace above. ' +
  'Returns {ok:true,opCount,appliedCount,revision,createdNodeIds,results,validation:{valid,errorCount,errors}} where createdNodeIds ' +
  'lists new node ids in document order (use them as anchors without re-inspecting) and validation proves the post-compose ' +
  'package state via a read-only OOXML check (save_document re-gates before writing). ' +
  'A compose that drops a non-empty redo stack signals it (redoCleared:true + redoClearedWarning naming the depth): ' +
  'a fresh edit invalidates redo. ' +
  'Optional styleGuide (bundled guide name or guide .md path, jailed like image refs: explicit allowedDir else the session anchor; remote URLs fail closed) ' +
  'replaces the session guide (not a one-batch override: later batches inherit it until another compose replaces it) and skins this compose (table/header-fill, chart-palette, callout, and code defaults apply where ops omit params; explicit params still win). ' +
  'Guides cannot be detached: open a new session to work without a guide. ' +
  'With no guide, compose uses the built-in Default Corporate look (navy table header, bold white header text, striped rows). styleGuide "legacy" gives the older grey header for this call only and is not stored on the session. ' +
  'An italic-only line directly under the opening # title becomes the Subtitle; a [!NOTE]-style line outside a blockquote stays plain text and comes back in warnings. ' +
  'Later batches inherit it, inspect_document shows its compact summary, save_document reports advisory styleWarnings. ' +
  'Unknown names fail with E_VALIDATION naming the available set. ' +
  'Math: $$ ... $$ or \\[ ... \\] blocks become display equations and \\(...\\) becomes inline math (one line, up to 500 characters; refused inside links and table cells). Display blocks are top level only (inside a list item or alert body they fail closed). Optional dollarMath (default false) treats $...$ as inline math; off by default because prices and shell text can look like math. ' +
  'Workflow: open_document -> compose_document -> save_document. ' +
  'Example {"sessionId":"sess_abc","markdown":"# Plan\\n\\n- a\\n- b\\n"} (add "allowedDir":"<dir>" for image/chart refs, "template":"report|memo|blank" to verify the seed). ' +
  'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions. ' +
  'persistentIds maps each created paragraph id (md_0003) to the paragraph id the saved file carries, so a reopened session can be addressed with it (tables are renumbered t_001.. by position on reopen).';

/** Created paragraph ids mapped to the w14:paraId a saved-and-reopened file reports for them. */
function persistentParagraphIds(ast: DocHandle['ast'], createdIds: string[]): Record<string, string> {
  const wanted = new Set(createdIds);
  const out: Record<string, string> = {};
  for (const section of ast.sections) {
    for (const node of section.nodes) {
      if (!wanted.has(node.id)) continue;
      const paraId = (node.xml as Element | undefined)?.getAttribute?.('w14:paraId');
      if (paraId && paraId.toUpperCase() !== node.id) out[node.id] = paraId.toUpperCase();
    }
  }
  return out;
}

type ComposeDocumentOk = {
  ok: true;
  sessionId: string;
  /** Operations the source compiled to (before chunking). */
  opCount: number;
  appliedCount: number;
  revision: number;
  /** True when the in-memory session differs from the last saved file (always true after a compose). */
  dirty: boolean;
  /** New node ids created by this compose, in post-compose document order. */
  createdNodeIds: string[];
  /**
   * Created paragraph ids (`md_0003`) mapped to the w14:paraId the saved file carries (what
   * a reopened session reports). Absent when nothing mapped; tables have no paraId and are
   * renumbered by position on reopen.
   */
  persistentIds?: Record<string, string>;
  /** Per-op attribution (same rule as execute_batch: exact id claims, remainder best-effort). */
  results: BatchOpCreatedIds[];
  /** Read-only post-compose OOXML verdict (proof the composed package validates; save re-gates before writing). */
  validation: {
    valid: boolean;
    errorCount: number;
    errors: Array<{ description: string; xPath?: string }>;
  };
  skippedIds: string[];
  warnings: string[];
  revisionsPending: number;
  revisionsRemedy?: string;
  possiblyStaleBodyRefs: { nodeId: string; preview: string }[];
  normalized: string[];
  /** True only when this compose dropped a non-empty redo stack (absent otherwise). */
  redoCleared?: boolean;
  /** Names the dropped redo depth (present only when redoCleared). */
  redoClearedWarning?: string;
  /** Compact summary of the guide this compose attached (absent when none attached). */
  styleGuide?: StyleGuideSummary;
  /** Parser diagnostics of the attached guide (present only when this compose attached one). */
  styleGuideWarnings?: string[];
};

type ComposeDocumentErr = {
  ok: false;
  code: EngineErrorCode;
  message: string;
  remedy: string;
};

function composeError(
  code: ComposeDocumentErr['code'],
  message: string,
  remedy: string
): ComposeDocumentErr {
  return {
    ok: false,
    code,
    message: rewritePhantomMcpToolNames(message),
    remedy: rewritePhantomMcpToolNames(remedy),
  };
}

/**
 * Q0b-gap: pure mapper for compose-stage failures (exported for the E-code
 * contract test). Every engine code survives verbatim via the shared
 * getEngineErrorCode table; E_VALIDATION stays the fallback for uncoded
 * failures. Previously E_LIMIT/E_SECURITY/E_DEPENDENCY/E_AUTH collapsed to
 * E_VALIDATION via the local KNOWN_COMPOSE_CODES allowlist.
 */
export function composeErrorFrom(
  err: unknown,
  fallbackRemedy: string
): ComposeDocumentErr {
  const resolved: EngineErrorCode = getEngineErrorCode(err) ?? 'E_VALIDATION';
  const rawMessage = err instanceof Error ? err.message : String(err);
  const remedy =
    (err as { remedy?: unknown } | null)?.remedy &&
    typeof (err as { remedy?: unknown }).remedy === 'string'
      ? String((err as { remedy?: unknown }).remedy)
      : fallbackRemedy;
  return composeError(resolved, sanitizeMcpClientMessage(rawMessage), remedy);
}

/** Live node ids in document order (single AST walk; same shape the batch path diffs). */
function liveNodeIdsInOrder(ast: {
  sections: Array<{ nodes: Array<{ id: string }> }>;
}): string[] {
  const out: string[] = [];
  try {
    for (const sec of ast.sections ?? []) {
      for (const n of sec.nodes ?? []) {
        if (typeof n.id === 'string' && n.id) out.push(n.id);
      }
    }
  } catch {
    /* best-effort: return what was collected */
  }
  return out;
}

/**
 * Jail check for the asset directory: the resolved canonical location must
 * sit inside one of the workspace roots. Returns the canonical absolute path.
 * Throws E_SANDBOX (with remedy) when the directory escapes the roots.
 */
export function assertPathAllowed(allowedDir: string, roots: string[]): string {
  if (typeof allowedDir !== 'string' || !allowedDir.trim()) {
    throw Object.assign(new Error('allowedDir must be a non-empty string'), {
      code: 'E_SANDBOX',
      remedy: 'Point allowedDir at the folder holding the referenced files.',
    });
  }
  if (allowedDir.includes('\0') || /%00/i.test(allowedDir)) {
    throw Object.assign(new Error('allowedDir path contains a null byte'), {
      code: 'E_SANDBOX',
      remedy: 'Remove null bytes from the path and retry.',
    });
  }
  if (allowedDir.includes('\n') || allowedDir.includes('\r')) {
    throw Object.assign(new Error('allowedDir path contains forbidden line break'), {
      code: 'E_SANDBOX',
      remedy: 'Remove line breaks from the path and retry.',
    });
  }
  const trimmed = allowedDir.trim();
  const segments = trimmed.split(/[\\/]/);
  if (segments.includes('..')) {
    throw Object.assign(new Error('allowedDir contains forbidden ".." traversal'), {
      code: 'E_SANDBOX',
      remedy: 'Use a folder inside the workspace documents directory or the system temp dir.',
    });
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
    if (process.platform === 'win32') {
      const lc = normC.toLowerCase();
      const lr = normR.toLowerCase();
      return lc === lr || lc.startsWith(lr + path.sep.toLowerCase());
    }
    return normC === normR || normC.startsWith(normR + path.sep);
  });
  if (!inside) {
    throw Object.assign(
      new Error(`allowedDir is outside the allowed directories (${roots.join(', ')})`),
      {
        code: 'E_SANDBOX',
        remedy: 'Point allowedDir at a folder under documents/ or the system temp dir, or set ASTRA_EXTRA_DOC_ROOTS.',
      }
    );
  }
  return canonical;
}

/**
 * Compile Markdown and apply it to an open session through the engine compose
 * path. The compile is pure (runs before the session lock, so input errors
 * write nothing and touch no session); execution delegates to
 * `composeDocument` on the session parser (chunked at the engine op cap with
 * template-seed check, then inline-span run splitting), so the strict schema
 * gate, quotas, structural checks and per-chunk atomic rollback all apply
 * exactly as they do for execute_batch, plus spans. Never throws — every
 * failure maps to {ok:false,code,remedy}.
 */
export async function handleComposeDocument(
  store: McpSessionStore,
  rawArgs: unknown
): Promise<ComposeDocumentOk | ComposeDocumentErr> {
  const parsed = ComposeDocumentSchema.safeParse(rawArgs);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return composeError(
      'E_VALIDATION',
      `compose_document: invalid input. ${detail}`,
      'Retry with {"sessionId":"<id from open_document>","markdown":"# Title\\n\\nBody text\\n"} (markdown 1..500000 chars; optional allowedDir for image/chart refs, optional template report|memo|blank to verify the seed).'
    );
  }

  const allowedDir = parsed.data.allowedDir;
  const template = parsed.data.template;

  // Jail gate before any compile or asset read: an allowedDir outside the
  // workspace roots fails closed here, even for sources without asset refs.
  // The canonical path is propagated to compile and compose below so the
  // engine never re-resolves a non-canonical alias.
  let canonicalAllowedDir: string | undefined;
  if (allowedDir !== undefined) {
    try {
      canonicalAllowedDir = assertPathAllowed(allowedDir, getMcpAllowedRoots());
    } catch (err: unknown) {
      return composeErrorFrom(
        err,
        'Point allowedDir at a folder under documents/ or the system temp dir, or set ASTRA_EXTRA_DOC_ROOTS.'
      );
    }
  }

  // Pure compile before the session lock: input errors (E_VALIDATION with a
  // 1-based line number) throw here, before any batch runs, so zero writes.
  // Asset refs resolve against allowedDir here exactly as they will inside
  // the compose call below (same options, deterministic output).
  // Schema-level input errors (including empty markdown rejected by the
  // min(1) gate) return E_VALIDATION here before the jail gate above runs
  // for a missing allowedDir; a present allowedDir is always jailed first.
  let operations: ReturnType<typeof compileMarkdown>;
  try {
    operations = compileMarkdown(parsed.data.markdown, {
      ...(canonicalAllowedDir !== undefined ? { allowedDir: canonicalAllowedDir } : {}),
    });
  } catch (err: unknown) {
    return composeErrorFrom(
      err,
      'Fix the flagged Markdown block (line-numbered) and retry compose_document.'
    );
  }
  try {
    validateOperationsStrict(operations);
  } catch (err: unknown) {
    return composeErrorFrom(
      err,
      'The compiled operations failed the schema gate; simplify the Markdown and retry compose_document.'
    );
  }

  const shadowsOf = (ops: ReturnType<typeof compileMarkdown>) => {
    const shadows: Array<Record<string, unknown>> = (ops as unknown as Array<Record<string, unknown>>).map((o) => ({
      ...(o as Record<string, unknown>),
    }));
    const targets: string[] = [];
    for (const o of shadows) {
      const t = o['target'];
      if (typeof t === 'string' && t) targets.push(t);
      else if (Array.isArray(t)) {
        for (const v of t) if (typeof v === 'string' && v) targets.push(v);
      }
    }
    return { shadows, actions: shadows.map((o) => String(o?.['action'] ?? '')), targets };
  };
  let { shadows: opShadows, actions: opActions, targets: opTargets } = shadowsOf(operations);

  try {
    return await store.withSession(parsed.data.sessionId, async (handle) => {
      const beforeIds = new Set<string>(liveNodeIdsInOrder(handle.ast));

      // Guide attach (persists for the session, skins this compose): an
      // explicit styleGuide resolves under the compose jail (explicit
      // allowedDir else the session anchor) and replaces the session guide;
      // omitted inherits whatever the session already carries. Unknown
      // names, jail escapes, and remote URLs fail closed before any write.
      // Empty (including whitespace-only) means "detach", which does not
      // exist — hint the new-session remedy instead of a bare schema error.
      let attachedWarnings: string[] | undefined;
      let attachedSummary: StyleGuideSummary | undefined;
      let legacyLook = false;
      if (parsed.data.styleGuide !== undefined) {
        if (parsed.data.styleGuide.trim() === '') {
          throw Object.assign(
            new Error(
              'compose_document: styleGuide is empty: guides cannot be detached — open a new session to work without a guide.'
            ),
            {
              code: 'E_VALIDATION',
              remedy: 'Omit styleGuide to inherit the session guide, or open a new session for guide-free composition.',
            }
          );
        }
        if (parsed.data.styleGuide.trim().toLowerCase() === 'legacy') {
          // Compose-only: the old unskinned look for this call; the session guide is left as it was.
          legacyLook = true;
        } else {
          const resolved = resolveMcpStyleGuide(parsed.data.styleGuide, {
            ...(canonicalAllowedDir !== undefined ? { allowedDir: canonicalAllowedDir } : {}),
            sessionFilePath: handle.filePath,
          });
          attachSessionStyleGuide(handle, parsed.data.styleGuide, resolved);
          attachedWarnings = [...resolved.warnings];
          attachedSummary = summarizeStyleGuide(resolved.config);
        }
      }
      const sessionGuide = legacyLook ? 'legacy' : handle.styleGuide;
      if (sessionGuide !== undefined) {
        // The pre-compile above used the default guide. Another guide can add or drop
        // ops (the zebra stripes), so rebuild what result attribution and the audit count.
        operations = compileMarkdown(parsed.data.markdown, {
          ...(canonicalAllowedDir !== undefined ? { allowedDir: canonicalAllowedDir } : {}),
          styleGuide: sessionGuide,
        });
        ({ shadows: opShadows, actions: opActions, targets: opTargets } = shadowsOf(operations));
      }

      // Single compose call: template-seed check, chunked batches at the
      // engine op cap, then inline-span run splitting. Throws on input
      // errors (line-numbered, before any write); returns a failing
      // aggregate on runtime batch errors.
      const aggregate = await composeDocument(handle.parser, parsed.data.markdown, {
        ...(canonicalAllowedDir !== undefined ? { allowedDir: canonicalAllowedDir } : {}),
        ...(template !== undefined ? { template } : {}),
        ...(sessionGuide !== undefined ? { styleGuide: sessionGuide } : {}),
        ...(parsed.data.dollarMath ? { dollarMath: true } : {}),
      });

      if (!aggregate.success) {
        const first = aggregate.errors[0];
        const rawMessage = String(first?.error ?? 'Compose failed.').slice(0, 2000);
        try {
          console.error(`[mcp compose_document] E_VALIDATION (client-sanitized): ${rawMessage.slice(0, 500)}`);
        } catch {
          /* stderr-only, never throws */
        }
        const hint = typeof first?.hint === 'string' && first.hint ? first.hint : undefined;
        throw Object.assign(new Error(`compose_document failed: ${sanitizeMcpClientMessage(rawMessage)}`), {
          code: 'E_VALIDATION',
          remedy:
            hint ?? 'Re-read ids via inspect_document and retry compose_document with the full Markdown source in one call.',
        });
      }

      // composeDocument works on a fresh AST from the parser, so refresh the
      // session handle in place (same object the session engine holds) before
      // diffing ids or serving later reads.
      await refreshAstInPlace(handle.ast, handle.parser);

      const appliedCount = aggregate.appliedCount;
      const warnings: string[] = [];
      for (const warning of aggregate.warnings ?? []) {
        if (typeof warning === 'string' && !warnings.includes(warning)) warnings.push(warning);
      }
      const skippedIds: string[] = [];
      for (const id of aggregate.skippedIds ?? []) {
        if (!skippedIds.includes(id)) skippedIds.push(id);
      }
      const staleRefs: { nodeId: string; preview: string }[] = [];
      for (const ref of aggregate.possiblyStaleBodyRefs ?? []) {
        const refRec = ref as unknown as Record<string, unknown>;
        if (
          typeof ref === 'object' &&
          ref !== null &&
          typeof refRec['nodeId'] === 'string' &&
          typeof refRec['preview'] === 'string' &&
          !staleRefs.some((entry) => entry.nodeId === (refRec['nodeId'] as string))
        ) {
          staleRefs.push({ nodeId: refRec['nodeId'] as string, preview: refRec['preview'] as string });
        }
      }
      const engineNormalized: string[] = [];
      for (const entry of readEngineNormalized(aggregate)) {
        if (!engineNormalized.includes(entry)) engineNormalized.push(entry);
      }
      const revisionsPending =
        typeof aggregate.revisionsPending === 'number'
          ? aggregate.revisionsPending
          : countPendingRevisionsInParser(handle.parser);
      const aggregateRemedy =
        typeof aggregate.revisionsRemedy === 'string' && aggregate.revisionsRemedy
          ? aggregate.revisionsRemedy
          : revisionsRemedy(revisionsPending);
      const revisionsRemedyText = aggregateRemedy ? aggregateRemedy : undefined;

      const revision = store.bumpRevision(handle);
      store.refreshByteSize(handle);
      // A fresh edit invalidates the redo chain (same rule as execute_batch).
      // The drop is signaled on the envelope so no caller discovers the
      // cleared stack via a surprise empty redo.
      const redoClearedDepth = handle.redoStack.length;
      handle.redoStack.length = 0;
      const afterOrder = liveNodeIdsInOrder(handle.ast);
      const createdNodeIds = afterOrder.filter((id) => !beforeIds.has(id));
      const results = attributeCreatedIds(opActions, opShadows, createdNodeIds);
      const persistentIds = persistentParagraphIds(handle.ast, createdNodeIds);
      store.accumulatePendingAudit(handle, appliedCount, [...opTargets, ...createdNodeIds]);
      const pendingRemedy = revisionsRemedyText ?? revisionsRemedy(revisionsPending);

      // Validation proof: read-only OOXML verdict over the composed package
      // (save_document re-gates before writing; this proof never writes).
      const validation = await handle.parser.validateOOXML();
      const rawErrors = Array.isArray(validation.errors) ? validation.errors.slice(0, COMPOSE_VALIDATION_ERRORS_MAX) : [];
      const validationErrors = rawErrors.map((e) => {
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
        ok: true as const,
        sessionId: handle.sessionId,
        opCount: operations.length,
        appliedCount,
        revision,
        dirty: store.isDirty(handle),
        createdNodeIds,
        ...(Object.keys(persistentIds).length > 0 ? { persistentIds } : {}),
        results,
        validation: {
          valid: validation.ok,
          errorCount: Array.isArray(validation.errors) ? validation.errors.length : 0,
          errors: validationErrors,
        },
        skippedIds,
        warnings,
        revisionsPending,
        ...(pendingRemedy ? { revisionsRemedy: pendingRemedy } : {}),
        possiblyStaleBodyRefs: staleRefs,
        normalized: mergeNormalized(engineNormalized, readParserNormalizations(handle.parser)),
        ...(attachedWarnings !== undefined && attachedSummary !== undefined
          ? {
              styleGuide: attachedSummary,
              styleGuideWarnings: attachedWarnings,
            }
          : {}),
        ...(redoClearedDepth > 0
          ? {
              redoCleared: true as const,
              redoClearedWarning:
                `This compose cleared ${redoClearedDepth} redo entr${redoClearedDepth === 1 ? 'y' : 'ies'}: ` +
                'a fresh edit invalidates redo, so redo_document now reports an empty stack. ' +
                'Call revert_document to record a fresh redoable state.',
            }
          : {}),
      };
    });
  } catch (err: unknown) {
    return composeErrorFrom(
      err,
      'Retry compose_document with the full Markdown source in one call (open_document first when the session expired).'
    );
  }
}
