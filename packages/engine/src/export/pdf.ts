import { execSync, spawn } from 'child_process';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import { pathToFileURL } from 'url';
import { DocumentAST } from '../ast';
import { assertWriteTargetNoSymlink, writeFileNoFollow } from '../sandbox/paths';
import { keepForForensics, secureFile, secureMkdtemp } from '../tmp';
import { PreflightError, runPreflightChecks } from './preflight';

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
}

/**
 * LibreOffice PDF filter spec for a tagged or untagged render.
 * The tagged path uses the plain filter (tagged output carries StructTreeRoot
 * + MarkInfo<</Marked true>>, verified on LO 25.2). The untagged path must
 * pass JSON filter data: LibreOffice 25.2 rejects the semicolon form
 * (`SelectPdfVersion=0;UseTaggedPDF=false`) with Unspecified Application Error.
 */
export function pdfFilterSpec(tagged: boolean): string {
  return tagged
    ? 'pdf:writer_pdf_Export'
    : 'pdf:writer_pdf_Export:{"UseTaggedPDF":{"type":"boolean","value":"false"}}';
}

let cachedBackendBinary: string | null | undefined = undefined;

/**
 * Checks if a headless PDF backend (LibreOffice / soffice) is available on the system.
 */
export function getPdfBackendBinary(): string | null {
  if (cachedBackendBinary !== undefined) {
    return cachedBackendBinary;
  }
  const candidates = ['libreoffice', 'soffice'];
  for (const bin of candidates) {
    try {
      execSync(`which ${bin}`, { stdio: 'ignore' });
      cachedBackendBinary = bin;
      return bin;
    } catch {
      // ignore
    }
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
 * Converts a raw .docx buffer into a PDF file on disk.
 * Uses headless LibreOffice conversion (tagged render by default).
 * Runs export preflight first: warn-only by default, fail-closed with
 * `strict: true` when blockers are present.
 * A failed tagged render is retried once as an untagged render; if both
 * fail a structured RenderError (code 'E_RENDER') is thrown.
 * Throws DependencyError (code 'E_DEPENDENCY') if LibreOffice is not installed.
 */
export async function convertDocxToPdf(
  docxBuffer: Buffer,
  outPdfPath: string,
  options?: PdfExportOptions
): Promise<string> {
  const binary = options?.binary || getPdfBackendBinary();
  if (!binary) {
    throw new DependencyError(
      `[E_DEPENDENCY] Headless PDF export requires LibreOffice ('libreoffice' or 'soffice'). ` +
      `Please install LibreOffice:\n` +
      `  Debian/Ubuntu: sudo apt install libreoffice\n` +
      `  Fedora/RHEL:   sudo dnf install libreoffice\n` +
      `  macOS:         brew install --cask libreoffice\n` +
      `  Windows:       winget install TheDocumentFoundation.LibreOffice`
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

  const tagged = options?.tagged ?? true;
  if (!tagged) {
    try {
      return await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(false), options, binary);
    } catch (err: unknown) {
      const msg = err instanceof Error ? err.message : String(err);
      throw new RenderError('', msg, 'untagged');
    }
  }
  // Failure dirs from the tagged stage are staged (not yet logged as
  // forensics) so a successful untagged fallback leaves no temp dirs behind.
  const stagedForensics: string[] = [];
  try {
    return await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(true), options, binary, stagedForensics);
  } catch (taggedErr: unknown) {
    const taggedMsg = taggedErr instanceof Error ? taggedErr.message : String(taggedErr);
    process.stderr.write(`[pdf] tagged render failed (${taggedMsg}); retrying with untagged render\n`);
    try {
      const result = await convertWithFilter(docxBuffer, outPdfPath, pdfFilterSpec(false), options, binary, stagedForensics);
      for (const dir of stagedForensics) {
        await removeStagedDir(dir);
      }
      return result;
    } catch (untaggedErr: unknown) {
      const untaggedMsg = untaggedErr instanceof Error ? untaggedErr.message : String(untaggedErr);
      for (const dir of stagedForensics) {
        keepForForensics(dir, 'PDF export failed');
      }
      throw new RenderError(taggedMsg, untaggedMsg, 'both');
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
    succeeded = true;
    return path.resolve(outPdfPath);
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
