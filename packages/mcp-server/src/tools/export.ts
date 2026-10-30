import { z } from 'zod';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import { McpSessionStore, canonicalizeForMcp, getMcpAllowedRoots } from '../lib/session-store';
import {
  AdmZip,
  assertValidDocx,
  assertWriteTargetNoSymlink,
  convertAstToHtml,
  convertAstToMarkdown,
  convertDocxToPdf,
  DocxParser,
  getEngineErrorCode,
  hasPdfBackend,
  runPreflightChecks,
  type DocumentAST,
  type PreflightIssue,
  writeFileNoFollow,
} from '@attentionform/astra-engine';

export const ExportDocumentSchema = z.object({
  sessionId: z.string().describe('Active document session ID returned by open_document'),
  format: z.enum(['html', 'markdown', 'md', 'pdf']).describe('Target export format: "html", "markdown", "md", or "pdf"'),
  outPath: z.string().optional().describe('Optional output destination path (must reside inside allowed directory)'),
  tagged_pdf: z.boolean().optional().describe('PDF only: tagged (accessible) render (default true)'),
  tagged: z.boolean().optional().describe('Alias of tagged_pdf'),
  strict: z.boolean().optional().describe('PDF only: fail closed on preflight blockers (default false, warn-only)'),
});

export type ExportDocumentArgs = z.infer<typeof ExportDocumentSchema>;

export const EXPORT_DOCUMENT_DESCRIPTION =
  'Export the active document to HTML, Markdown, or PDF format with sandbox security gating. ' +
  'Renders the in-memory session state (includes unsaved execute_batch edits, not the last saved file); ' +
  'success returns {ok:true,format,path|content,bytes,dirty,revision} (dirty:true = unsaved edits present). ' +
  'PDF renders via headless LibreOffice: a ~11s first-render cold start is expected behavior, not a bug. ' +
  'Failures return {ok:false,code,message,remedy}.';

/**
 * Q0b: pure mapper for export-stage failures (exported for the E-code
 * contract test). Every engine code survives verbatim via
 * getEngineErrorCode; E_IO stays the fallback for uncoded failures.
 */
export function mapExportError(err: unknown): { code: string; message: string; remedy: string } {
  const msg = err instanceof Error ? err.message : String(err);
  // Q0b: central engine table. Previously E_SANDBOX/E_DEPENDENCY/
  // E_VALIDATION passed via .code-or-message and any other .code passed
  // through, but a dropped code carried only in the message (E_LIMIT/
  // E_SECURITY without a .code prop) collapsed to E_IO. Now every engine
  // code survives verbatim via getEngineErrorCode; E_IO stays the fallback.
  const preserved = getEngineErrorCode(err);
  let code = 'E_IO';
  if (preserved !== undefined) {
    code = preserved;
  } else {
    const errCode = (err as { code?: string })?.code;
    if (errCode) code = errCode;
  }
  return {
    code,
    message: `Export failed: ${msg}`,
    remedy: (err as { remedy?: string })?.remedy || 'Check disk space, file permissions, and backend availability.',
  };
}

export function validateExportTarget(rawPath: unknown, expectedFormat: 'html' | 'markdown' | 'md' | 'pdf'): string {
  const normFormat = expectedFormat === 'md' ? 'markdown' : expectedFormat;
  if (typeof rawPath !== 'string' || !rawPath.trim()) {
    throw Object.assign(new Error('Export path must be a non-empty string'), {
      code: 'E_SANDBOX',
      remedy: 'Pass a valid output path inside an allowed directory.',
    });
  }
  if (rawPath.includes('\0') || /%00/i.test(rawPath)) {
    throw Object.assign(new Error('Export path contains a null byte'), {
      code: 'E_SANDBOX',
      remedy: 'Remove null bytes from the path and retry.',
    });
  }
  if (rawPath.includes('\n') || rawPath.includes('\r')) {
    throw Object.assign(new Error('Export path contains forbidden line break'), {
      code: 'E_SANDBOX',
      remedy: 'Remove line breaks from the path and retry.',
    });
  }
  const trimmed = rawPath.trim();
  const segments = trimmed.split(/[\\/]/);
  if (segments.includes('..')) {
    throw Object.assign(new Error('Export path contains forbidden ".." traversal'), {
      code: 'E_SANDBOX',
      remedy: 'Use a path inside the workspace documents directory.',
    });
  }

  const resolved = path.resolve(trimmed);
  const canonical = canonicalizeForMcp(resolved);
  const allowedRoots = getMcpAllowedRoots();

  const isInside = allowedRoots.some((root) => {
    let canonicalRoot = root;
    try {
      canonicalRoot = fs.realpathSync(root);
    } catch {
      // ignore
    }
    const normC = path.normalize(canonical);
    const normR = path.normalize(canonicalRoot);
    return normC === normR || normC.startsWith(normR + path.sep);
  });

  if (!isInside) {
    throw Object.assign(
      new Error(`Export path is outside allowed directories (${allowedRoots.join(', ')})`),
      {
        code: 'E_SANDBOX',
        remedy: 'Place output under documents/ or the system temp dir, or set ASTRA_EXTRA_DOC_ROOTS.',
      }
    );
  }

  const ext = path.extname(trimmed).toLowerCase();
  const validExts: Record<string, string[]> = {
    html: ['.html', '.htm'],
    markdown: ['.md', '.markdown'],
    pdf: ['.pdf'],
  };
  if (!validExts[normFormat]?.includes(ext)) {
    throw Object.assign(
      new Error(`Export path for format "${expectedFormat}" must use a valid extension (${validExts[normFormat]?.join(', ')}), got "${ext || '(none)'}"`),
      {
        code: 'E_VALIDATION',
        remedy: `Use a valid extension for ${expectedFormat}.`,
      }
    );
  }

  return resolved;
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
 * rewrite on the fork while the exported bytes still reflect live edits.
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

export async function handleExportDocument(
  store: McpSessionStore,
  rawArgs: unknown
  // Phase C-1a BREAKING: failure field renamed `error` -> `message` to match
  // the lean-6 {ok:false,code,message,remedy} envelope (visual_inspect already
  // used `message`). Clients reading `error` must migrate to `message`.
  // Phase C-1b: success envelopes carry dirty + revision (exports render the
  // in-memory session state, not the last saved file).
  // Preflight diagnostics ride along additively as warnings/blockers arrays.
): Promise<{ ok: boolean; format?: string; path?: string; bytes?: number; content?: string; dirty?: boolean; revision?: number; warnings?: PreflightIssue[]; blockers?: PreflightIssue[]; message?: string; code?: string; remedy?: string }> {
  const parsed = ExportDocumentSchema.safeParse(rawArgs);
  if (!parsed.success) {
    return {
      ok: false,
      code: 'E_VALIDATION',
      message: `Invalid export arguments: ${parsed.error.issues.map(i => `${i.path.join('.')}: ${i.message}`).join(', ')}`,
      remedy: 'Provide valid sessionId and format ("html" | "markdown" | "md" | "pdf").',
    };
  }

  const { sessionId, format: rawFormat, outPath } = parsed.data;
  const format: 'html' | 'markdown' | 'pdf' = rawFormat === 'md' ? 'markdown' : rawFormat;
  let session;
  try {
    session = store.get(sessionId);
  } catch {
    return {
      ok: false,
      code: 'E_SESSION',
      message: `Session "${sessionId}" not found or has been closed.`,
      remedy: 'Call open_document first to establish an active session.',
    };
  }
  // Phase C-1b: in-memory session state vs last saved file.
  const dirty = store.isDirty(session);
  const revision = session.revision;

  let targetPath: string | undefined;
  if (outPath) {
    try {
      targetPath = validateExportTarget(outPath, format);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      return {
        ok: false,
        code: (err as { code?: string })?.code || 'E_SANDBOX',
        message: msg,
        remedy: (err as { remedy?: string })?.remedy || 'Choose a valid path inside the allowed directory.',
      };
    }
  }

  let validatedPdfPath: string | undefined;
  if (format === 'pdf') {
    if (!hasPdfBackend()) {
      return {
        ok: false,
        code: 'E_DEPENDENCY',
        message: `[E_DEPENDENCY] Headless PDF export requires LibreOffice ('libreoffice' or 'soffice') which is not found in PATH.`,
        remedy: 'Install LibreOffice on the host system: sudo apt install libreoffice (Linux) or brew install --cask libreoffice (macOS).',
      };
    }
    const defaultPdf = targetPath || path.join(path.dirname(session.filePath), `${path.basename(session.filePath, path.extname(session.filePath))}.pdf`);
    try {
      validatedPdfPath = validateExportTarget(defaultPdf, 'pdf');
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      return {
        ok: false,
        code: (err as { code?: string })?.code || 'E_SANDBOX',
        message: msg,
        remedy: (err as { remedy?: string })?.remedy || 'Choose a valid path inside the allowed directory.',
      };
    }
  }

  try {
    if (format === 'pdf') {
      // M1: refuse dangling-symlink PDF targets before spawning LibreOffice.
      assertWriteTargetNoSymlink(validatedPdfPath!);
      const buf = serializeSessionCopy(session.parser, session.ast);
      await assertValidDocx(buf, 'PDF export source');
      let preflight: { warnings: PreflightIssue[]; blockers: PreflightIssue[] } | undefined;
      try {
        const report = runPreflightChecks(buf);
        preflight = { warnings: report.warnings, blockers: report.blockers };
      } catch {
        preflight = undefined;
      }
      try {
        await convertDocxToPdf(buf, validatedPdfPath!, {
          tagged: parsed.data.tagged_pdf ?? parsed.data.tagged ?? true,
          strict: parsed.data.strict ?? false,
        });
      } catch (err: unknown) {
        // Carry preflight diagnostics on the original error (additive: the
        // outer envelope picks up warnings/blockers/remedy from these fields).
        const original = err instanceof Error ? err : new Error(String(err));
        if (preflight) {
          const withDiag = original as Error & { warnings?: PreflightIssue[]; blockers?: PreflightIssue[] };
          withDiag.warnings = preflight.warnings;
          withDiag.blockers = (original as { blockers?: PreflightIssue[] }).blockers ?? preflight.blockers;
        }
        throw original;
      }
      const stat = await fsp.stat(validatedPdfPath!);
      return {
        ok: true,
        format: 'pdf',
        path: validatedPdfPath,
        bytes: stat.size,
        dirty,
        revision,
        warnings: preflight?.warnings ?? [],
        blockers: preflight?.blockers ?? [],
      };
    }

    const ast = session.ast;

    if (format === 'html') {
      const title = path.basename(session.filePath, path.extname(session.filePath));
      const html = convertAstToHtml(ast, title, undefined, session.parser);
      if (targetPath) {
        // M1: dangling-symlink leaf refused before mkdir/write (same guard
        // as save_document: lstat + O_NOFOLLOW write).
        assertWriteTargetNoSymlink(targetPath);
        await fsp.mkdir(path.dirname(targetPath), { recursive: true });
        assertWriteTargetNoSymlink(targetPath);
        await writeFileNoFollow(targetPath, html, 'utf-8');
        const stat = await fsp.stat(targetPath);
        return { ok: true, format: 'html', path: targetPath, bytes: stat.size, dirty, revision };
      }
      return { ok: true, format: 'html', content: html, bytes: Buffer.byteLength(html, 'utf-8'), dirty, revision };
    } else {
      const md = convertAstToMarkdown(ast, session.parser);
      if (targetPath) {
        // M1: same dangling-symlink guard as above.
        assertWriteTargetNoSymlink(targetPath);
        await fsp.mkdir(path.dirname(targetPath), { recursive: true });
        assertWriteTargetNoSymlink(targetPath);
        await writeFileNoFollow(targetPath, md, 'utf-8');
        const stat = await fsp.stat(targetPath);
        return { ok: true, format: 'markdown', path: targetPath, bytes: stat.size, dirty, revision };
      }
      return { ok: true, format: 'markdown', content: md, bytes: Buffer.byteLength(md, 'utf-8'), dirty, revision };
    }
  } catch (err: unknown) {
    const mapped = mapExportError(err);
    return {
      ok: false,
      ...mapped,
      ...((err as { warnings?: PreflightIssue[] })?.warnings !== undefined
        ? { warnings: (err as { warnings?: PreflightIssue[] }).warnings }
        : {}),
      ...((err as { blockers?: PreflightIssue[] })?.blockers !== undefined
        ? { blockers: (err as { blockers?: PreflightIssue[] }).blockers }
        : {}),
    };
  }
}
