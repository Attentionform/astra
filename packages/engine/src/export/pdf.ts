import { spawn } from 'child_process';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import { pathToFileURL } from 'url';
import { DocumentAST } from '../ast';
import { assertWriteTargetNoSymlink, writeFileNoFollow } from '../sandbox/paths';
import { keepForForensics, secureFile, secureMkdtemp } from '../tmp';
import { PreflightError, runPreflightChecks } from './preflight';
import { findOnPath } from '../visual/backend';
import { runPlan } from '../visual/render-plan';
import {
  AUTO_PREFERS_WORD,
  effectiveRendererPreference,
  planRenderers,
  wordBreaker,
  type RendererPreference,
  type RendererReport,
} from '../visual/renderer-preference';
import {
  analyseReferences,
  buildReferenceReport,
  effectiveAllowNetworkReferences,
  emptyReferenceReport,
  neutraliseReferences,
  removeUpdateFields,
  type ReferenceReport,
} from '../visual/references';
import { findWord, wordDocxToPdf, type WordInstall } from '../visual/word';
import {
  assertLibreOfficeMarkupSupport,
  buildMarkupReport,
  effectiveShowMarkup,
  libreOfficeFilterSpec,
  libreOfficeMarkupWarnings,
  markupFailureError,
  type MarkupReport,
} from '../visual/markup';
import type { VisualTestHooks } from '../visual/provider';

export class DependencyError extends Error {
  code = 'E_DEPENDENCY';
  constructor(message: string) {
    super(message);
    this.name = 'DependencyError';
  }
}

export class RenderError extends Error {
  code = 'E_RENDER';
  taggedDetail: string;
  untaggedDetail: string;
  /** Which render run failed: 'tagged' for a tagged-only failure, 'untagged' for an untagged-only failure, 'both' when both runs failed. */
  stage: 'tagged' | 'untagged' | 'both';
  /** LibreOffice process exit code when parsable from the failure detail. */
  exitCode?: number;
  /** Underlying failure text (the stage detail that best explains the failure). */
  cause?: string;
  /** Document repair hint returned to callers (MCP surfaces this as its remedy). */
  remedy: string;
  constructor(taggedDetail: string, untaggedDetail: string, stage?: 'tagged' | 'untagged' | 'both') {
    super(
      `[E_RENDER] PDF render failed (tagged: ${taggedDetail}; untagged: ${untaggedDetail}). ` +
        `Remedy: repair the document in Word (accept revisions, fix charts/notes) and retry, ` +
        `or export to HTML instead.`
    );
    this.name = 'RenderError';
    this.taggedDetail = taggedDetail;
    this.untaggedDetail = untaggedDetail;
    this.stage =
      stage ?? (taggedDetail && untaggedDetail ? 'both' : taggedDetail ? 'tagged' : untaggedDetail ? 'untagged' : 'both');
    this.exitCode =
      this.stage === 'tagged'
        ? parseLoExitCode(taggedDetail) ?? parseLoExitCode(untaggedDetail)
        : parseLoExitCode(untaggedDetail) ?? parseLoExitCode(taggedDetail);
    this.cause =
      (this.stage === 'tagged' ? taggedDetail || untaggedDetail : untaggedDetail || taggedDetail) || undefined;
    this.remedy =
      'Repair the document in Word (accept revisions, fix charts/notes) and retry, ' +
      'or export to HTML instead.';
  }
}

/** Extract the LibreOffice exit code from a `LibreOffice exited with code N` detail. */
function parseLoExitCode(detail: string): number | undefined {
  const m = /exited with code (\d+)/.exec(detail);
  return m ? Number(m[1]) : undefined;
}

// Re-exported so MCP-layer callers importing from '@attentionform/astra-engine' can run
// preflight and carry its diagnostics without depending on deep paths.
export { PreflightError, runPreflightChecks } from './preflight';
export type { PreflightIssue, PreflightReport } from './preflight';

export interface PdfExportOptions {
  timeoutMs?: number;
  binary?: string;
  /** Tagged (accessible) PDF render. Default true. */
  tagged?: boolean;
  /** Fail closed on preflight blockers. Default false (warn-only). */
  strict?: boolean;
  /** DOCX -> PDF program: auto (default), word (Windows only) or libreoffice. Falls back to ASTRA_RENDERER, then auto. */
  renderer?: RendererPreference;
  /**
   * Let the renderer load network locations the document names (default false: such references are retargeted
   * on the temporary copy). An explicit value, including `false`, beats ASTRA_ALLOW_NETWORK_REFERENCES.
   */
  allowNetworkReferences?: boolean;
  /**
   * Draw tracked changes and comments (default false). Falls back to ASTRA_SHOW_MARKUP. Word: the export with
   * markup; LibreOffice: comments in the margin. Pages can be scaled and pagination can change.
   */
  showMarkup?: boolean;
  /** Tests only: replace what the machine would report and the PowerShell host. */
  testHooks?: VisualTestHooks;
}

/** What a PDF export did, beyond the file it wrote. */
export interface PdfConversionResult {
  path: string;
  renderer: RendererReport;
  references: ReferenceReport;
  /** Present only when markup was requested: what was asked and what the renderer was told to do. */
  markup?: MarkupReport;
}

/**
 * LibreOffice PDF filter spec for a tagged or untagged render.
 * The tagged path uses the plain filter (tagged output carries StructTreeRoot
 * + MarkInfo<</Marked true>>, verified on LO 25.2). The untagged path must
 * pass JSON filter data: LibreOffice 25.2 rejects the semicolon form
 * (`SelectPdfVersion=0;UseTaggedPDF=false`) with Unspecified Application Error.
 */
export function pdfFilterSpec(tagged: boolean, extra: { markup?: boolean } = {}): string {
  return libreOfficeFilterSpec(tagged, extra.markup === true);
}

let cachedBackendBinary: string | null | undefined = undefined;

/**
 * Checks if a headless PDF backend (LibreOffice / soffice) is available on the system.
 */
export function getPdfBackendBinary(): string | null {
  if (cachedBackendBinary !== undefined) {
    return cachedBackendBinary;
  }
  // Shared resolver: PATH search (Windows .exe/.com aware), then the
  // standard install folders. Returns an absolute path, so the spawn does
  // not depend on PATH lookup rules of the platform.
  const found = findOnPath('soffice') ?? findOnPath('libreoffice');
  if (found) {
    cachedBackendBinary = found;
    return found;
  }
  cachedBackendBinary = null;
  return null;
}

export function hasPdfBackend(): boolean {
  return getPdfBackendBinary() !== null;
}

export const hasExportPdfBackend = hasPdfBackend;

export function resetPdfBackendCache(): void {
  cachedBackendBinary = undefined;
}

export function setPdfBackendBinaryForTesting(bin: string | null): void {
  cachedBackendBinary = bin;
}

/**
 * Converts a raw .docx buffer into a PDF file on disk and returns its path.
 * See `convertDocxToPdfDetailed` for the behaviour and for what else a conversion reports.
 */
export async function convertDocxToPdf(
  docxBuffer: Buffer,
  outPdfPath: string,
  options?: PdfExportOptions
): Promise<string> {
  return (await convertDocxToPdfDetailed(docxBuffer, outPdfPath, options)).path;
}

/**
 * Converts a raw .docx buffer into a PDF file on disk and says how.
 * Uses headless LibreOffice, or Word on Windows when asked (`renderer: 'word'`); `auto` is LibreOffice in 0.5.2.
 * Runs export preflight first: warn-only by default, fail-closed with
 * `strict: true` when blockers are present. Network references are retargeted on the render copy unless
 * `allowNetworkReferences` is true (the source bytes are never changed).
 * With LibreOffice, a failed tagged render is retried once as an untagged render; if both
 * fail a structured RenderError (code 'E_RENDER') is thrown. A Word failure is thrown as a WordRenderError
 * (`code`, `remedy` and `failureClass` set), never wrapped.
 * Throws DependencyError (code 'E_DEPENDENCY') if no renderer is installed.
 */
export async function convertDocxToPdfDetailed(
  docxBuffer: Buffer,
  outPdfPath: string,
  options?: PdfExportOptions
): Promise<PdfConversionResult> {
  const hooks: VisualTestHooks = options?.testHooks ?? {};
  const platform = hooks.platform ?? process.platform;
  // Choices first: an invalid value fails loudly before anything runs.
  const rendererPref = effectiveRendererPreference(options?.renderer);
  const allowRefs = effectiveAllowNetworkReferences(options?.allowNetworkReferences);
  const showMarkup = effectiveShowMarkup(options?.showMarkup);

  const binary = options?.binary || getPdfBackendBinary();
  const wordMayRun = rendererPref === 'word' || (rendererPref === 'auto' && (hooks.autoPrefersWord ?? AUTO_PREFERS_WORD));
  const word: WordInstall | null = wordMayRun && platform === 'win32' ? (hooks.word !== undefined ? hooks.word : findWord()) : null;
  const plan = planRenderers(rendererPref, {
    platform,
    word: word !== null,
    libreoffice: Boolean(binary),
    wordUsable: !wordBreaker.isOpen(),
    autoPrefersWord: hooks.autoPrefersWord,
  });
  if (plan.engines.length === 0) {
    throw new DependencyError(
      `[E_DEPENDENCY] Headless PDF export requires LibreOffice ('libreoffice' or 'soffice'). ` +
      `Please install LibreOffice:\n` +
      `  Debian/Ubuntu: sudo apt install libreoffice\n` +
      `  Fedora/RHEL:   sudo dnf install libreoffice\n` +
      `  macOS:         brew install --cask libreoffice\n` +
      `  Windows:       winget install TheDocumentFoundation.LibreOffice (or use Microsoft Word: --renderer word)\n` +
      `ASTra also looks in the standard install folders (for example C:\\Program Files\\LibreOffice\\program ` +
      `and /Applications/LibreOffice.app/Contents/MacOS).`
    );
  }

  const preflight = runPreflightChecks(docxBuffer);
  for (const w of preflight.warnings) {
    process.stderr.write(`[preflight] warning ${w.code}: ${w.message}\n`);
  }
  for (const b of preflight.blockers) {
    process.stderr.write(`[preflight] blocker ${b.code}: ${b.message}\n`);
  }
  if ((options?.strict ?? false) && preflight.blockers.length > 0) {
    throw new PreflightError(preflight.blockers);
  }

  // Network references: retargeted on the render copy unless allowed. Failing to check fails the export
  // (fail closed); in allowed mode the scan only reports, so its failure is not fatal.
  let renderBytes = docxBuffer;
  let references: ReferenceReport;
  if (allowRefs) {
    try {
      references = buildReferenceReport(true, analyseReferences(docxBuffer, { mutate: false }));
    } catch {
      references = { ...emptyReferenceReport(true), scanFailed: true };
    }
  } else {
    const neutralised = neutraliseReferences(docxBuffer);
    renderBytes = neutralised.docx;
    references = buildReferenceReport(false, neutralised);
  }
  renderBytes = removeUpdateFields(renderBytes);

  // Failure dirs are staged (not yet logged as forensics) so a successful render by any engine or filter
  // leaves no temp dirs behind.
  const staged: string[] = [];
  try {
    const run = await runPlan(
      plan.engines,
      { canFallBack: plan.canFallBack, auto: plan.requested === 'auto' },
      (engine) =>
        engine === 'word'
          ? convertWithWord(renderBytes, outPdfPath, options, staged, hooks, showMarkup)
          : convertWithLibreOffice(renderBytes, outPdfPath, options, binary as string, staged, showMarkup),
      { word: wordBreaker }
    );
    for (const dir of staged) {
      await removeStagedDir(dir);
    }
    const renderer: RendererReport = {
      requested: plan.requested,
      used: run.used,
      ...(run.used === 'word' && word ? { version: word.version } : {}),
      ...(run.fallbackFrom ? { fallbackFrom: run.fallbackFrom, fallbackReason: run.fallbackReason } : {}),
    };
    return { path: run.value.path, renderer, references, ...(showMarkup ? { markup: buildMarkupReport(run.used, run.value.markupWarnings) } : {}) };
  } catch (err: unknown) {
    for (const dir of staged) {
      keepForForensics(dir, 'PDF export failed');
    }
    throw err;
  }
}

/** LibreOffice: tagged render, retried once untagged; both failing is a structured RenderError. */
async function convertWithLibreOffice(
  docxBuffer: Buffer,
  outPdfPath: string,
  options: PdfExportOptions | undefined,
  binary: string,
  staged: string[],
  markup: boolean
): Promise<ExportStep> {
  if (markup) return convertWithLibreOfficeMarkup(docxBuffer, outPdfPath, options, binary, staged);
  return { path: await convertWithLibreOfficePlain(docxBuffer, outPdfPath, options, binary, staged) };
}

/**
 * With markup: one render with the margin-comments option, at the requested tagging. The option is only sent to a
 * tested LibreOffice version. A failure is never retried silently without the option; one diagnostic render without
 * it (its output is discarded) lets the error say what this call showed.
 */
async function convertWithLibreOfficeMarkup(
  docxBuffer: Buffer,
  outPdfPath: string,
  options: PdfExportOptions | undefined,
  binary: string,
  staged: string[]
): Promise<ExportStep> {
  const markupWarnings = (assertLibreOfficeMarkupSupport(binary), libreOfficeMarkupWarnings());
  const tagged = options?.tagged ?? true;
  try {
    return { path: await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(tagged, { markup: true }), options, binary, staged), markupWarnings };
  } catch (markupErr: unknown) {
    const markupMsg = markupErr instanceof Error ? markupErr.message : String(markupErr);
    const diagnostic = path.join(secureMkdtemp('astra-pdf-diag-'), 'diagnostic.pdf');
    let plain: { ok: true } | { ok: false; error: string };
    try {
      await convertWithFilter(docxBuffer, diagnostic, pdfFilterSpec(tagged), options, binary, undefined);
      plain = { ok: true };
    } catch (plainErr: unknown) {
      plain = { ok: false, error: plainErr instanceof Error ? plainErr.message : String(plainErr) };
    } finally {
      try {
        await fsp.rm(path.dirname(diagnostic), { recursive: true, force: true });
      } catch {
        /* best effort: the diagnostic file is never returned */
      }
    }
    throw markupFailureError(markupMsg, plain);
  }
}

/** LibreOffice: tagged render, retried once untagged; both failing is a structured RenderError. */
async function convertWithLibreOfficePlain(
  docxBuffer: Buffer,
  outPdfPath: string,
  options: PdfExportOptions | undefined,
  binary: string,
  staged: string[]
): Promise<string> {
  const tagged = options?.tagged ?? true;
  if (!tagged) {
    try {
      return await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(false), options, binary, staged);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new RenderError('', msg, 'untagged');
    }
  }
  try {
    return await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(true), options, binary, staged);
  } catch (taggedErr: unknown) {
    const taggedMsg = taggedErr instanceof Error ? taggedErr.message : String(taggedErr);
    process.stderr.write(`[pdf] tagged render failed (${taggedMsg}); retrying with untagged render\n`);
    try {
      return await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(false), options, binary, staged);
    } catch (untaggedErr: unknown) {
      const untaggedMsg = untaggedErr instanceof Error ? untaggedErr.message : String(untaggedErr);
      throw new RenderError(taggedMsg, untaggedMsg, 'both');
    }
  }
}

/** What one engine returns: the file it wrote, and any best-effort Word view settings that could not be set. */
interface ExportStep {
  path: string;
  markupWarnings?: string[];
}

/** Word (Windows): one export; the tagged flag maps to Word's document-structure-tags switch. */
async function convertWithWord(
  docxBuffer: Buffer,
  outPdfPath: string,
  options: PdfExportOptions | undefined,
  staged: string[],
  hooks: VisualTestHooks,
  markup: boolean
): Promise<ExportStep> {
  const tmpDir = secureMkdtemp('astra-pdf-');
  const tmpDocxPath = path.join(tmpDir, `doc_${Date.now()}.docx`);
  let succeeded = false;
  try {
    await fsp.writeFile(tmpDocxPath, docxBuffer);
    secureFile(tmpDocxPath);
    const generated = path.join(tmpDir, 'word-output.pdf');
    const outcome = await wordDocxToPdf(tmpDocxPath, generated, {
      tagged: options?.tagged ?? true,
      timeoutMs: options?.timeoutMs || 60000,
      ...(markup ? { markup: true, ...(hooks.wordStartState ? { startState: hooks.wordStartState } : {}) } : {}),
      powershellBin: hooks.powershellBin,
      powershellArgs: hooks.powershellArgs,
      ...(hooks.wordRetryPauseMs ? { retryPauseMs: hooks.wordRetryPauseMs } : {}),
    });
    const out = await copyOutPdf(generated, outPdfPath);
    succeeded = true;
    return { path: out, ...(outcome.markup ? { markupWarnings: outcome.markup.warnings } : {}) };
  } finally {
    if (succeeded) {
      try {
        await fsp.rm(tmpDir, { recursive: true, force: true });
      } catch {
        keepForForensics(tmpDir, 'could not remove tmp dir after success');
      }
    } else {
      staged.push(tmpDir);
    }
  }
}

/**
 * Per-filter LibreOffice retry cap. Total LO spawns per convertDocxToPdf call
 * are bounded at 2 filter stages x MAX_CONVERT_ATTEMPTS = 6 processes.
 */
const MAX_CONVERT_ATTEMPTS = 3;

/** Best-effort removal of a staged failure dir after fallback success (one retry; logs when retained). */
async function removeStagedDir(dir: string): Promise<void> {
  for (let attempt = 1; attempt <= 2; attempt++) {
    try {
      await fsp.rm(dir, { recursive: true, force: true });
      return;
    } catch {
      if (attempt < 2) {
        await new Promise((r) => setTimeout(r, 100));
      }
    }
  }
  keepForForensics(dir, 'could not remove staged tmp dir after fallback success');
}

/** Copy a generated PDF to its destination with the hardened write (no symlink follow, 0644). */
async function copyOutPdf(generatedPdfPath: string, outPdfPath: string): Promise<string> {
  const targetDir = path.dirname(path.resolve(outPdfPath));
  // M1: refuse dangling-symlink PDF targets before the copy-out lands
  // (lstat chain + O_NOFOLLOW write closes the check-then-use window).
  assertWriteTargetNoSymlink(path.resolve(outPdfPath));
  await fsp.mkdir(targetDir, { recursive: true });
  assertWriteTargetNoSymlink(path.resolve(outPdfPath));

  await writeFileNoFollow(path.resolve(outPdfPath), await fsp.readFile(generatedPdfPath));
  // Normalize the output mode with the sibling exporters (0644): the
  // hardened write above creates 0600 (LO-temp inheritance) while
  // docx/html/md outputs land 0644. Explicit chmod, not umask-dependent.
  fs.chmodSync(path.resolve(outPdfPath), 0o644);
  return path.resolve(outPdfPath);
}

async function convertWithFilter(
  docxBuffer: Buffer,
  outPdfPath: string,
  filterSpec: string,
  options: PdfExportOptions | undefined,
  binary: string,
  stagedForensics?: string[]
): Promise<string> {
  const tmpDir = secureMkdtemp('astra-pdf-');
  const tmpDocxPath = path.join(tmpDir, `doc_${Date.now()}.docx`);

  let succeeded = false;
  try {
    await fsp.writeFile(tmpDocxPath, docxBuffer);
    secureFile(tmpDocxPath);

    let converted = false;
    let lastError: Error | null = null;
    for (let attempt = 1; attempt <= MAX_CONVERT_ATTEMPTS; attempt++) {
      try {
        await new Promise<void>((resolve, reject) => {
          const child = spawn(
            binary,
            [`-env:UserInstallation=${pathToFileURL(path.join(tmpDir, 'lo_profile')).href}`, '--headless', '--convert-to', filterSpec, tmpDocxPath, '--outdir', tmpDir],
            { stdio: ['ignore', 'pipe', 'pipe'] }
          );

          let stderr = '';
          child.stderr.on('data', (d) => { stderr += d.toString(); });

          const timeout = setTimeout(() => {
            child.kill('SIGKILL');
            reject(new Error(`PDF conversion timed out after ${options?.timeoutMs || 30000}ms`));
          }, options?.timeoutMs || 30000);

          child.on('close', (code) => {
            clearTimeout(timeout);
            if (code === 0) {
              resolve();
            } else {
              reject(new Error(`LibreOffice exited with code ${code}: ${stderr.trim()}`));
            }
          });

          child.on('error', (err) => {
            clearTimeout(timeout);
            reject(err);
          });
        });
        converted = true;
        break;
      } catch (err: unknown) {
        lastError = err instanceof Error ? err : new Error(String(err));
        if (attempt < 3) {
          await new Promise((r) => setTimeout(r, 250 * attempt));
        }
      }
    }
    if (!converted && lastError) {
      throw lastError;
    }

    const expectedPdfName = path.basename(tmpDocxPath).replace(/\.docx$/i, '.pdf');
    const generatedPdfPath = path.join(tmpDir, expectedPdfName);

    if (!fs.existsSync(generatedPdfPath)) {
      throw new Error(`PDF conversion failed: output file ${generatedPdfPath} was not created`);
    }

    const written = await copyOutPdf(generatedPdfPath, outPdfPath);
    succeeded = true;
    return written;
  } finally {
    if (succeeded) {
      try {
        await fsp.rm(tmpDir, { recursive: true, force: true });
      } catch {
        keepForForensics(tmpDir, 'could not remove tmp dir after success');
      }
    } else if (stagedForensics) {
      stagedForensics.push(tmpDir);
    } else {
      keepForForensics(tmpDir, 'PDF export failed');
    }
  }
}

/**
 * Converts a DocumentAST into a PDF file on disk.
 * If docxBuffer is supplied, it is converted directly.
 */
export async function convertAstToPdf(
  ast: DocumentAST,
  outPdfPath: string,
  docxBuffer?: Buffer,
  options?: PdfExportOptions
): Promise<string> {
  if (docxBuffer) {
    return convertDocxToPdf(docxBuffer, outPdfPath, options);
  }
  throw new Error('convertAstToPdf requires docxBuffer of the OOXML package to export to PDF');
}
