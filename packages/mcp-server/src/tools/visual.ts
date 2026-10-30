#!/usr/bin/env node
/**
 * @attentionform/astra-mcp-server visual tool.
 *
 * Dual-mode visual review for MCP:
 * - Level 1 (default, always-on, <500 tokens): deterministic spatial
 *   telemetry (ASCII wireframe + large-defect warnings) in every inspect.
 * - Level 2 (opt-in `withVisual:true` tool param only, no launch flag): headless
 *   LO -> PDF -> pdftoppm render (PNG default, JPEG opt-in) returning
 *   server-side FILE PATHS + metadata (never base64 — token bloat).
 *
 * Missing binaries degrade gracefully with `{ok:false, code:'E_DEPENDENCY',
 * remedy}` (install LibreOffice + poppler-utils). Prompt law applies:
 * image is approximate, XML is truth, large defects only.
 *
 * Engine wiring: session store via ../lib/session-store, visual review via
 * @attentionform/astra-engine (single source of truth, no drift).
 * Standalone module: registered in `src/index.ts` tool table.
 */

import * as fs from 'fs';
import * as path from 'path';
import { z } from 'zod';
import { getMcpAllowedRoots, type McpSessionStore } from '../lib/session-store';
import { assertPathAllowed } from './compose-document';
import {
  AdmZip,
  buildVisualTelemetry,
  DocxParser,
  estimateTokenCount,
  getEngineErrorCode,
  getVisualArtifactsDir,
  hasVisualPdfBackend as hasPdfBackend,
  renderDocxToImages,
  VISUAL_PROMPT_LAW,
  VISUAL_TELEMETRY_TOKEN_BUDGET,
  VISUAL_VERIFY_RULE,
  writeFileNoFollowSync,
  type DocumentAST,
  type EngineErrorCode,
  type VisualImageFormat,
  type VisualPageMap,
  type VisualTelemetry,
} from '@attentionform/astra-engine';

/** Legacy argv helper; opt-in renders use the withVisual tool param (startup rejects unknown flags). */
export const MCP_VISUAL_FLAG = '--with-visual';

/** True when argv carries `--with-visual`; startup rejects unknown flags so this enables nothing. */
export function isVisualEnabled(argv: readonly string[] = process.argv): boolean {
  try {
    return argv.includes(MCP_VISUAL_FLAG);
  } catch {
    return false;
  }
}

// Q0b: full central engine set (previously 5; previously-dropped E_LIMIT/
// E_SECURITY/E_CONFLICT/E_NOT_FOUND/E_QUOTA/E_STATE/E_AUTH now pass through
// verbatim instead of collapsing to E_IO/E_SESSION).
export type McpVisualErrorCode = EngineErrorCode;

export interface McpVisualError {
  ok: false;
  code: McpVisualErrorCode;
  message: string;
  remedy: string;
}

export interface VisualImageRef {
  page: number;
  path: string;
  format: VisualImageFormat;
  bytes: number;
}

export interface VisualInspectOk {
  ok: true;
  sessionId: string;
  revision: number;
  /**
   * Phase C-1b: true when the in-memory session holds unsaved execute_batch
   * edits — visual renders in-memory session state, not the last saved file.
   */
  dirty: boolean;
  /** Always-on spatial telemetry (<500 tokens). */
  telemetry: VisualTelemetry;
  /** Prompt law + verify rule (vision-model instruction). */
  promptLaw: string;
  verifyRule: string;
  /** Opt-in renders (only when withVisual:true). File paths, not base64. */
  images?: VisualImageRef[];
  totalPages: number;
  outDir?: string;
  /**
   * Advisory stale-seed hint (additive, present only when untouched template
   * seed tables are detected; absent when clean). Lists table node ids whose
   * text still matches known financial-template seed values.
   */
  staleTableIds?: string[];
  /** One-line advisory naming staleTableIds + heuristic limits. */
  staleContentHint?: string;
  /** True when the stale list hit STALE_SEED_TABLE_MAX (ids truncated). */
  staleTruncated?: boolean;
}

function visualError(code: McpVisualErrorCode, message: string, remedy: string): McpVisualError {
  return { ok: false, code, message, remedy };
}

/**
 * Q0b: pure mapper for render-stage failures (exported for the E-code
 * contract test). Previously every non-DEPENDENCY failure collapsed to E_IO;
 * now the engine code survives verbatim, E_IO stays the uncoded fallback.
 */
export function mapVisualRenderError(err: unknown): McpVisualError {
  const preserved = getEngineErrorCode(err);
  if (preserved === 'E_DEPENDENCY') {
    const remedy =
      (err as { remedy?: unknown })?.remedy && typeof (err as { remedy?: unknown }).remedy === 'string'
        ? String((err as { remedy?: unknown }).remedy)
        : 'Install LibreOffice and poppler-utils, or inspect without withVisual.';
    const msg = err instanceof Error ? err.message : String(err);
    return visualError('E_DEPENDENCY', msg, remedy);
  }
  const msg = err instanceof Error ? err.message : String(err);
  if (preserved !== undefined) {
    const remedy =
      (err as { remedy?: unknown })?.remedy && typeof (err as { remedy?: unknown }).remedy === 'string'
        ? String((err as { remedy?: unknown }).remedy)
        : 'Retry telemetry-only (withVisual:false) or fix the document and retry.';
    return visualError(preserved, `visual_inspect: render failed: ${msg.slice(0, 1000)}`, remedy);
  }
  return visualError('E_IO', `visual_inspect: render failed: ${msg.slice(0, 1000)}`, 'Retry telemetry-only (withVisual:false) or fix the document and retry.');
}

/**
 * Q0b: pure mapper for outer (session-resolution) failures. Previously every
 * failure collapsed to E_SESSION; now the engine code survives verbatim,
 * E_SESSION stays the fallback for unknown/expired sessions.
 */
export function mapVisualOuterError(err: unknown): McpVisualError {
  const msg = err instanceof Error ? err.message : String(err);
  const preserved = getEngineErrorCode(err);
  if (preserved !== undefined) {
    const remedy =
      (err as { remedy?: unknown })?.remedy && typeof (err as { remedy?: unknown }).remedy === 'string'
        ? String((err as { remedy?: unknown }).remedy)
        : preserved === 'E_SESSION'
          ? 'Call open_document first, then retry with the returned sessionId.'
          : 'Call open_document first, then retry visual inspect.';
    return visualError(preserved, msg, remedy);
  }
  // withSession throws E_SESSION for unknown/expired sessions.
  if (msg.includes('unknown session') || msg.includes('expired') || msg.includes('session')) {
    return visualError('E_SESSION', msg, 'Call open_document first, then retry with the returned sessionId.');
  }
  return visualError('E_SESSION', msg, 'Call open_document first, then retry visual inspect.');
}

export const VisualInspectSchema = z
  .object({
    sessionId: z.string().min(1).max(128),
    /** Pages to inspect/render (1-indexed). Default: page 1 telemetry. */
    pages: z.array(z.number().int().min(1).max(1000)).max(4).optional(),
    /** Opt-in vision renders (file paths). Default false (telemetry only). */
    withVisual: z.boolean().optional(),
    /** PNG default (lossless); JPEG opt-in. */
    format: z.enum(['png', 'jpeg']).optional(),
    jpegQuality: z.number().int().min(1).max(100).optional(),
    dpi: z.number().int().min(72).max(300).optional(),
    /**
     * Folder to put the rendered images in (with withVisual:true), for clients that
     * cannot read the server's temp folder. Must be inside the allowed folders, like a
     * document path. Default: the server's temp artifacts folder.
     */
    outputDir: z.string().min(1).max(1024).optional(),
  })
  .strict();

export type VisualInspectInput = z.infer<typeof VisualInspectSchema>;

export const VISUAL_INSPECT_DESCRIPTION =
  'Inspect document visually with 2D spatial layout telemetry (ASCII wireframe, page breakdown, margins, defects) ' +
  'and optional headless rendered images (withVisual: true; add outputDir to write them to a folder inside the allowed directories, for example next to the document, instead of the server temp folder). Renders the in-memory session state ' +
  '(includes unsaved execute_batch edits, not the last saved file); success returns ' +
  '{ok:true,sessionId,revision,dirty,telemetry,...} (dirty:true = unsaved edits present).';

/**
 * Stale seed-table hint (advisory, additive, bounded).
 *
 * Detects tables that still carry untouched `financial`-template seed values
 * (Metric/Previous Q/Current Q/Net Revenue seed figures). A table counts as
 * stale when its concatenated text matches >=2 markers — a single generic
 * word (e.g. "Metric") never fires alone. Present only when stale tables
 * exist; clean docs stay silent (no fields emitted).
 *
 * Limits (honest): matches known financial-template seed strings only;
 * custom templates and paraphrased stale content are NOT covered; an edited
 * table still matches until ALL seed strings are replaced. Capped at
 * STALE_SEED_TABLE_MAX ids (staleTruncated:true when capped).
 */
export const STALE_SEED_TABLE_MARKERS = [
  'Net Revenue',
  'Operating Profit',
  '$1,250,000',
  '$1,520,000',
  '$340,000',
  '$415,000',
  'Previous Q',
  'Current Q',
  'Growth (%)',
] as const;

/** Cap on stale table ids listed in one visual_inspect response. */
export const STALE_SEED_TABLE_MAX = 10;

type StaleAst = {
  sections?: Array<{ nodes?: Array<{ id?: unknown; type?: unknown; metadata?: Record<string, unknown> }> }>;
};

function tableNodeText(meta: Record<string, unknown>): string {
  const parts: string[] = [];
  try {
    const t = meta['text'];
    if (typeof t === 'string' && t) parts.push(t);
    const headers = meta['headers'];
    if (Array.isArray(headers)) {
      for (const h of headers) if (typeof h === 'string' && h) parts.push(h);
    }
    const rows = meta['rows'];
    if (Array.isArray(rows)) {
      for (const r of rows) {
        if (Array.isArray(r)) {
          for (const c of r) if (typeof c === 'string' && c) parts.push(c);
        } else if (typeof r === 'string' && r) {
          parts.push(r);
        }
      }
    }
  } catch {
    /* best-effort: return what was collected */
  }
  return parts.join('\n');
}

/** Table node ids whose text still matches known template seed values (bounded). */
export function findStaleSeedTables(ast: StaleAst): { ids: string[]; truncated: boolean } {
  const ids: string[] = [];
  let truncated = false;
  try {
    for (const sec of ast.sections ?? []) {
      for (const n of sec.nodes ?? []) {
        if (!n || n.type !== 'table') continue;
        if (typeof n.id !== 'string' || !n.id) continue;
        const text = tableNodeText((n.metadata ?? {}) as Record<string, unknown>);
        if (!text) continue;
        let hits = 0;
        for (const m of STALE_SEED_TABLE_MARKERS) {
          if (text.includes(m)) {
            hits += 1;
            if (hits >= 2) break;
          }
        }
        if (hits >= 2) {
          if (ids.length < STALE_SEED_TABLE_MAX) {
            ids.push(n.id);
          } else {
            truncated = true;
          }
        }
      }
    }
  } catch {
    truncated = true;
  }
  return { ids, truncated };
}

/** One-line advisory hint naming stale table ids (call only when ids non-empty). */
export function buildStaleContentHint(ids: readonly string[], truncated: boolean): string {
  const named = ids.join(', ');
  return (
    `Advisory: ${ids.length} table(s) still contain template seed values (${named}). ` +
    'Replace seed figures before shipping. ' +
    `Limits: heuristic matches known financial-template seeds only (>=2 of ${STALE_SEED_TABLE_MARKERS.length} strings); ` +
    'edited tables still match until all seed text is replaced; custom templates not covered.' +
    (truncated ? ` List capped at ${STALE_SEED_TABLE_MAX} (truncated).` : '')
  );
}

/**
 * Pure telemetry enrichment for `inspect_document` (no MCP store needed).
 * Returns the wireframe + warnings + token count, budget-enforced <500.
 */
export function buildInspectVisualTelemetry(
  ast: Parameters<typeof buildVisualTelemetry>[0],
  opts: { page?: number; pageMap?: VisualPageMap | null; renderTotalPages?: number } = {}
): { telemetry: VisualTelemetry; promptLaw: string; verifyRule: string } {
  const telemetry = buildVisualTelemetry(ast, {
    page: opts.page ?? 1,
    pageMap: opts.pageMap ?? null,
    renderTotalPages: opts.renderTotalPages,
  });
  return { telemetry, promptLaw: VISUAL_PROMPT_LAW, verifyRule: VISUAL_VERIFY_RULE };
}

/** Capability probe passthrough (respects PATH at call time). */
export function visualBackendStatus(): ReturnType<typeof hasPdfBackend> {
  return hasPdfBackend();
}

/** Artifacts dir passthrough (session workspace folder). */
export function visualArtifactsDir(subDir?: string): string {
  return getVisualArtifactsDir(subDir);
}

export function visualTelemetryBudget(): number {
  return VISUAL_TELEMETRY_TOKEN_BUDGET;
}

export function estimateVisualTokens(text: string): number {
  return estimateTokenCount(text);
}

/**
 * Serializes the live session through an isolated parser copy.
 *
 * Read paths must never disturb live session state: the core rezip heals
 * the DOM it runs on (allowlist repairs, canonical ordering, timestamp
 * clamps, numbering syncs) and records what it did on the parser. Running
 * it on the live parser would heal the session silently, so a later save
 * could no longer report that work as its own. Snapshotting the package,
 * rebuilding a throwaway parser, and rezipping the copy keeps every
 * rewrite on the fork while the rendered bytes still reflect live edits.
 */
function serializeSessionCopy(parser: DocxParser, ast: unknown): Buffer {
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
  const fork = new DocxParser(zip.toBuffer());
  try {
    const live = ast as {
      numbering?: { abstractNum?: Record<string, unknown>; concreteNum?: Record<string, unknown> };
      relationships?: Record<string, unknown>;
    } | null;
    if (live && typeof live === 'object') {
      fork.adoptAstForNumbering({
        numbering: {
          abstractNum: { ...((live.numbering?.abstractNum ?? {}) as Record<string, unknown>) },
          concreteNum: { ...((live.numbering?.concreteNum ?? {}) as Record<string, unknown>) },
        },
        relationships: { ...((live.relationships ?? {}) as Record<string, unknown>) },
      } as unknown as DocumentAST);
    }
  } catch {
    // Fall through: a bare fork rezip preserves the snapshot as-is.
  }
  return fork.rezip();
}

/**
 * Moves rendered images into the caller's folder as `<document>-page-<n>.<ext>`.
 * The folder was already checked against the allowed roots. The write refuses a
 * symlink anywhere on the path and opens the name with O_NOFOLLOW, so a link
 * planted before or during the write cannot redirect it. A file already under
 * that name is replaced: it is the previous render of the same page.
 */
export function placeImages(images: VisualImageRef[], dir: string, docPath: string): VisualImageRef[] {
  fs.mkdirSync(dir, { recursive: true });
  const base = path.basename(docPath || 'document', path.extname(docPath || '')).replace(/[^a-zA-Z0-9._-]+/g, '-') || 'document';
  return images.map((img) => {
    const ext = path.extname(img.path) || (img.format === 'jpeg' ? '.jpg' : '.png');
    const dest = path.join(dir, `${base}-page-${img.page}${ext}`);
    writeFileNoFollowSync(dest, fs.readFileSync(img.path));
    try {
      fs.rmSync(img.path, { force: true });
    } catch {
      /* the temp copy is pruned with its run folder */
    }
    return { ...img, path: dest };
  });
}

/**
 * Store-aware handler: telemetry always, file-path images only when
 * `withVisual:true`. Never throws — maps every failure to the
 * `{ok:false, code, remedy}` contract (E_DEPENDENCY for missing binaries).
 */
export async function handleVisualInspect(
  store: McpSessionStore,
  input: unknown
): Promise<VisualInspectOk | McpVisualError> {
  const parsed = VisualInspectSchema.safeParse(input);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return visualError('E_VALIDATION', `visual_inspect: invalid input. ${detail}`, 'Retry with {sessionId} and optional {pages, withVisual, format}.');
  }
  const { sessionId, pages, withVisual, format, jpegQuality, dpi, outputDir } = parsed.data;
  if (outputDir !== undefined && !withVisual) {
    return visualError('E_VALIDATION', 'visual_inspect: outputDir only applies to rendered images.', 'Add withVisual:true, or drop outputDir for telemetry only.');
  }
  let canonicalOutputDir: string | undefined;
  if (outputDir !== undefined) {
    try {
      canonicalOutputDir = assertPathAllowed(outputDir, getMcpAllowedRoots());
    } catch (err: unknown) {
      return mapVisualOuterError(err);
    }
  }
  const requestedPage = pages && pages.length > 0 ? pages[0] : 1;
  try {
    return await store.withSession(sessionId, async (handle) => {
      // Prefer the renderer-compatible map when the parser exposes one;
      // otherwise the telemetry layer falls back to a labeled estimate.
      let embeddedMap: VisualPageMap | null = null;
      try {
        const summary = handle.parser.getStructuralSummary() as unknown as {
          nodePages?: Record<string, number>;
          totalPages?: number;
        } | null;
        if (summary && summary.nodePages && Object.keys(summary.nodePages).length > 0) {
          embeddedMap = {
            totalPages: typeof summary.totalPages === 'number' && summary.totalPages >= 1 ? Math.floor(summary.totalPages) : 1,
            nodePages: summary.nodePages,
            source: 'renderer',
            isEstimate: false,
          };
        }
      } catch {
        embeddedMap = null;
      }
      if (!withVisual) {
        // Telemetry-only: word-layout map only (no render count available),
        // byte-identical to the pre-render build this handler used for every
        // call before the renderer count became authoritative.
        const { telemetry, promptLaw, verifyRule } = buildInspectVisualTelemetry(handle.ast, {
          page: requestedPage,
          pageMap: embeddedMap,
        });
        const stale = findStaleSeedTables(handle.ast as unknown as StaleAst);
        const staleExtras =
          stale.ids.length > 0
            ? {
                staleTableIds: stale.ids,
                staleContentHint: buildStaleContentHint(stale.ids, stale.truncated),
                ...(stale.truncated ? { staleTruncated: true as const } : {}),
              }
            : {};
        return {
          ok: true as const,
          sessionId: handle.sessionId,
          revision: handle.revision,
          dirty: store.isDirty(handle),
          telemetry,
          promptLaw,
          verifyRule,
          totalPages: telemetry.totalPages,
          ...staleExtras,
        };
      }

      // Opt-in vision: headless render to file-path images. Telemetry is
      // built AFTER the render so the renderer's page count is the single
      // source of truth for `[PAGE n OF m]` when it disagrees with the
      // word-layout map (same rule as the CLI `visual` command); the
      // telemetry-only path above stays word-layout-based.
      let docx: Buffer;
      try {
        docx = serializeSessionCopy(handle.parser, handle.ast);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        return visualError('E_IO', `visual_inspect: cannot serialize document: ${msg}`, 'Retry inspect without withVisual (telemetry-only).');
      }
      try {
        const rendered = await renderDocxToImages(docx, {
          pages,
          format: (format ?? 'png') as VisualImageFormat,
          ...(jpegQuality !== undefined ? { jpegQuality } : {}),
          ...(dpi !== undefined ? { dpi } : {}),
        });
        // The marker must describe the image actually produced: page picks
        // clamp to the renderer's last page exactly like the CLI. Guard:
        // never emit `totalPages: 0` — the renderer contract is >=1 page (it
        // throws otherwise); fall back to the word-layout telemetry total if
        // a stubbed/future provider violates it.
        const renderTotalPages = rendered.totalPages >= 1 ? Math.floor(rendered.totalPages) : undefined;
        const picked = rendered.images.find((img) => img.page === requestedPage) ?? rendered.images[0];
        const { telemetry, promptLaw, verifyRule } = buildInspectVisualTelemetry(handle.ast, {
          page: picked?.page ?? requestedPage,
          pageMap: embeddedMap,
          renderTotalPages,
        });
        let images: VisualImageRef[] = rendered.images.map((img) => ({
          page: img.page,
          path: img.path,
          format: img.format,
          bytes: img.bytes,
        }));
        let outDir = rendered.outDir;
        if (canonicalOutputDir !== undefined) {
          images = placeImages(images, canonicalOutputDir, handle.filePath);
          outDir = canonicalOutputDir;
        }
        const stale = findStaleSeedTables(handle.ast as unknown as StaleAst);
        const staleExtras =
          stale.ids.length > 0
            ? {
                staleTableIds: stale.ids,
                staleContentHint: buildStaleContentHint(stale.ids, stale.truncated),
                ...(stale.truncated ? { staleTruncated: true as const } : {}),
              }
            : {};
        return {
          ok: true as const,
          sessionId: handle.sessionId,
          revision: handle.revision,
          dirty: store.isDirty(handle),
          telemetry,
          promptLaw,
          verifyRule,
          images,
          totalPages: renderTotalPages ?? telemetry.totalPages,
          outDir,
          ...staleExtras,
        };
      } catch (err: unknown) {
        return mapVisualRenderError(err);
      }
    });
  } catch (err: unknown) {
    return mapVisualOuterError(err);
  }
}
