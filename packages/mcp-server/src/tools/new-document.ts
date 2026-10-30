import { z } from 'zod';
import * as fsp from 'fs/promises';
import * as path from 'path';
import { DocxParser, getEngineErrorCode, assertWriteTargetNoSymlink, summarizeStyleGuide, writeFileNoFollow, type EngineErrorCode, type ResolvedStyleGuide, type StyleGuideSummary } from '@attentionform/astra-engine';
import {
  buildClientSessionIdRejection,
  validateSaveTarget,
  type McpSessionStore,
} from '../lib/session-store';
import { attachSessionStyleGuide, resolveMcpStyleGuide } from './compose-document';

/** Templates mirrored from the CLI `create` command (cli.ts CREATE_TEMPLATES). */
export const NEW_DOCUMENT_TEMPLATES = ['report', 'memo', 'blank'] as const;
export type NewDocumentTemplate = (typeof NEW_DOCUMENT_TEMPLATES)[number];

export const NewDocumentSchema = z
  .object({
    outPath: z.string().min(1).max(1024),
    template: z.enum(NEW_DOCUMENT_TEMPLATES).optional(),
    force: z.boolean().optional(),
    // Style guide attach-at-creation (bundled name or guide file path).
    // Persists for the session; batches inherit it unless they override.
    // Empty is rejected in the handler with a detach hint — no min(1) here
    // on purpose so the hint (not a bare schema error) reaches the caller.
    styleGuide: z.string().max(1024).optional(),
    // Default author for tracked changes and comments that do not name one (falls back to
    // ASTRA_AUTHOR on the server, then "ASTra"). An operation's own author always wins.
    author: z.string().trim().min(1).max(128).optional(),
    sessionId: z
      .string()
      .min(1)
      .max(128)
      .optional()
      .describe(
        'DEPRECATED: session ids are server-minted (sess_<uuid>); omit sessionId and use the returned id. ' +
          'Client-supplied values are rejected with E_VALIDATION (deterministic ids are test-only via --allow-client-session-ids).'
      ),
  })
  .strict();

export type NewDocumentInput = z.infer<typeof NewDocumentSchema>;

export const NEW_DOCUMENT_DESCRIPTION =
  'Create a new .docx from a template (report|memo|blank, default blank) at outPath inside an allowed directory ' +
  '(--allowed-dir / documents / system temp). Validates the fresh package before writing ' +
  '(E_VALIDATION when invalid, never writes corruption); refuses to overwrite without {force:true} (E_CONFLICT). ' +
  'Success returns {ok:true,path,sessionId,revision,template,bytes} with the file opened into a session ' +
  '(use sessionId for inspect/execute_batch/save). Example {"outPath":"documents/new-report.docx","template":"report"}. ' +
  'Session ids are server-minted (sess_<uuid>, DEPRECATED to supply sessionId: client-supplied values are ' +
  'rejected with E_VALIDATION; deterministic ids are test-only via --allow-client-session-ids). ' +
  'IDs are session-scoped: valid only in the issuing session, must not be stored across sessions. ' +
  'Optional styleGuide (bundled guide name or guide .md path, jailed to the document folder; remote URLs fail closed) ' +
  'attaches a house style to the session at creation: later batches inherit its table/chart/callout defaults, ' +
  'inspect_document shows its compact summary, save_document reports advisory styleWarnings. ' +
  'Guides cannot be detached afterwards: open a new session to work without a guide. ' +
  'Unknown names fail with E_VALIDATION naming the available set.';

type NewDocumentOk = {
  ok: true;
  path: string;
  sessionId: string;
  revision: number;
  template: NewDocumentTemplate;
  bytes: number;
  /** Compact summary of the guide attached at creation (absent when none). */
  styleGuide?: StyleGuideSummary;
  /** Parser diagnostics of the attached guide (present only when attached). */
  styleGuideWarnings?: string[];
};

type NewDocumentErr = {
  ok: false;
  code: EngineErrorCode;
  message: string;
  remedy: string;
};

function newDocError(
  code: NewDocumentErr['code'],
  message: string,
  remedy: string
): NewDocumentErr {
  return { ok: false, code, message, remedy };
}

/**
 * Q0b-gap: resolve the engine code verbatim via the shared table; fallback
 * applies only when the error carries no known code (previously E_LIMIT/
 * E_SECURITY/E_DEPENDENCY/E_AUTH and any other non-listed code collapsed to
 * the per-site fallback). Exported for the E-code contract test.
 */
export function resolveNewDocumentCode(err: unknown, fallback: EngineErrorCode): EngineErrorCode {
  return getEngineErrorCode(err) ?? fallback;
}

/**
 * Create a template document on disk and open it into the session store.
 * Mirrors CLI `create` semantics (template gate + validate-before-write);
 * the write target is jailed to the MCP allowed roots (same jail as save).
 * Never throws — every failure maps to {ok:false,code,remedy}.
 */
export async function handleNewDocument(
  store: McpSessionStore,
  rawArgs: unknown
): Promise<NewDocumentOk | NewDocumentErr> {
  const parsed = NewDocumentSchema.safeParse(rawArgs);
  if (!parsed.success) {
    const detail = parsed.error.issues.map((i) => `${i.path.join('.') || '(root)'}: ${i.message}`).join('; ');
    return newDocError(
      'E_VALIDATION',
      `new_document: invalid input. ${detail}`,
      'Retry with {"outPath":"<allowed-dir>/file.docx","template":"report|memo|blank"} (template optional, default blank).'
    );
  }
  const template: NewDocumentTemplate = parsed.data.template ?? 'blank';
  const force = parsed.data.force ?? false;

  // Fail before writing: a squatted sessionId would otherwise write the file
  // and then surface as E_SESSION from the open below.
  if (
    typeof parsed.data.sessionId === 'string' &&
    parsed.data.sessionId.trim() &&
    !store.getAllowClientSessionIds()
  ) {
    const rej = buildClientSessionIdRejection('new_document', parsed.data.sessionId);
    return newDocError(rej.code, rej.message, rej.remedy);
  }

  let target: string;
  try {
    target = validateSaveTarget(parsed.data.outPath);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return newDocError(
      resolveNewDocumentCode(err, 'E_SANDBOX'),
      msg,
      (err as { remedy?: string })?.remedy || 'Choose a .docx path inside an allowed directory.'
    );
  }

  // Guide attach resolves before any write: unknown names, jail escapes,
  // and remote URLs fail closed here (E_VALIDATION/E_SANDBOX + remedy) and
  // write nothing. File refs jail to the new document's own folder.
  let guideRef: string | undefined;
  let resolvedGuide: ResolvedStyleGuide | undefined;
  if (parsed.data.styleGuide !== undefined) {
    guideRef = parsed.data.styleGuide;
    if (guideRef.trim() === '') {
      return newDocError(
        'E_VALIDATION',
        'new_document: styleGuide is empty: guides cannot be detached — open a new session to work without a guide.',
        'Omit styleGuide to create a guide-free session, or pass a bundled guide name or guide file path.'
      );
    }
    try {
      resolvedGuide = resolveMcpStyleGuide(guideRef, { sessionFilePath: target });
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      return newDocError(
        resolveNewDocumentCode(err, 'E_VALIDATION'),
        msg,
        (err as { remedy?: string })?.remedy || 'Pass a bundled guide name or a guide file path inside the allowed directory.'
      );
    }
  }

  if (!force) {
    try {
      await fsp.access(target);
      return newDocError(
        'E_CONFLICT',
        `new_document: destination "${path.basename(target)}" already exists (refusing to overwrite without force)`,
        'Retry with {"force":true} to overwrite, or choose a different outPath.'
      );
    } catch {
      /* destination is free — proceed */
    }
  }

  let parser: DocxParser;
  try {
    parser = DocxParser.createNewDocument(template);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return newDocError(
      resolveNewDocumentCode(err, 'E_IO'),
      `new_document: template build failed: ${msg}`,
      (err as { remedy?: string })?.remedy || 'Retry with template blank and report the template name.'
    );
  }

  try {
    const validation = await parser.validateOOXML();
    if (!validation.ok) {
      const first = (validation.errors?.[0] ?? {}) as { description?: unknown; message?: unknown; xPath?: unknown; path?: unknown };
      const desc =
        (typeof first.description === 'string' && first.description) ||
        (typeof first.message === 'string' && first.message) ||
        'OOXML validation failed';
      const xp = typeof first.xPath === 'string' ? first.xPath : typeof first.path === 'string' ? first.path : undefined;
      const at = typeof xp === 'string' ? ` at ${xp}` : '';
      return newDocError(
        'E_VALIDATION',
        `new_document: fresh "${template}" template failed validation${at}: ${String(desc).slice(0, 500)}`,
        'Retry with template blank and report the template name.'
      );
    }
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return newDocError(
      resolveNewDocumentCode(err, 'E_VALIDATION'),
      `new_document: validation failed: ${msg.slice(0, 500)}`,
      (err as { remedy?: string })?.remedy || 'Retry with template blank and report the template name.'
    );
  }

  let buf: Buffer;
  try {
    buf = parser.rezip();
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return newDocError(
      resolveNewDocumentCode(err, 'E_IO'),
      `new_document: cannot serialize document: ${msg}`,
      (err as { remedy?: string })?.remedy || 'Retry new_document; re-open if the store is at capacity.'
    );
  }

  try {
    // M1: dangling-symlink leaf refused before mkdir/write (same guard as
    // save_document: lstat the full component chain + O_NOFOLLOW write).
    assertWriteTargetNoSymlink(target);
    await fsp.mkdir(path.dirname(target), { recursive: true });
    assertWriteTargetNoSymlink(target);
    await writeFileNoFollow(target, buf);
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    return newDocError(
      resolveNewDocumentCode(err, 'E_IO'),
      `new_document: write failed for "${target}": ${msg}`,
      (err as { remedy?: string })?.remedy || 'Check the outPath directory is writable and retry new_document.'
    );
  }

  try {
    const handle = await store.openDocument(target, parsed.data.sessionId);
    if (parsed.data.author !== undefined) handle.ast.settings.defaultAuthor = parsed.data.author;
    let bytes = buf.length;
    try {
      bytes = (await fsp.stat(target)).size;
    } catch {
      /* keep in-memory size */
    }
    if (guideRef !== undefined && resolvedGuide !== undefined) {
      attachSessionStyleGuide(handle, guideRef, resolvedGuide);
    }
    return {
      ok: true,
      path: target,
      sessionId: handle.sessionId,
      revision: handle.revision,
      template,
      bytes,
      ...(resolvedGuide !== undefined
        ? {
            styleGuide: summarizeStyleGuide(resolvedGuide.config),
            styleGuideWarnings: [...resolvedGuide.warnings],
          }
        : {}),
    };
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    const code = resolveNewDocumentCode(err, 'E_SESSION');
    return newDocError(
      code,
      `new_document: file written but session open failed (${code}): ${msg.slice(0, 500)}`,
      (err as { remedy?: string })?.remedy || 'Call open_document with the returned path to establish a session.'
    );
  }
}
