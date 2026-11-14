/**
 * Headless document renderer: DOCX -> PDF (LibreOffice, or Word on Windows) -> PNG/JPEG (Windows built-in PDF renderer or pdftoppm).
 *
 * PNG default (lossless glyph edges for vision models); JPEG opt-in via
 * `{format:'jpeg', jpegQuality}` for bandwidth-sensitive hosts.
 * Writes page images to the artifacts dir and returns FILE PATHS
 * (never base64 — MCP/CLI pass paths + metadata to avoid token bloat).
 *
 * Pure Node. Both trees keep IDENTICAL copies.
 */

import * as fs from 'fs';
import * as path from 'path';
import { pathToFileURL } from 'url';
import { execFile } from 'child_process';
import { promisify } from 'util';
import { createDependencyError, resolvePoppler, resolveSoffice } from './backend';
import {
  AUTO_PREFERS_WORD,
  effectiveRasterizerPreference,
  effectiveRendererPreference,
  planRasterizers,
  planRenderers,
  windowsRasterizerBreaker,
  wordBreaker,
  type RasterizerReport,
  type RendererReport,
} from './renderer-preference';
import { runPlan } from './render-plan';
import {
  analyseReferences,
  buildReferenceReport,
  effectiveAllowNetworkReferences,
  emptyReferenceReport,
  neutraliseReferences,
  removeUpdateFields,
  type ReferenceReport,
} from './references';
import { findWord, wordDocxToPdf, type WordInstall } from './word';
import {
  assertLibreOfficeMarkupSupport,
  buildMarkupReport,
  effectiveShowMarkup,
  libreOfficeFilterSpec,
  libreOfficeMarkupWarnings,
  markupFailureError,
  type PlainRunOutcome,
} from './markup';
import { windowsPdfToImages, windowsRendererAvailable } from './windows-pdf';
import { canonicalizeForAsset as canonicalizeVisualPath, isPathInsideDir as isVisualInsideDir } from '../sandbox/paths';
import { keepForForensics, removeDirBestEffort, resolveAstraTmpBase, secureDir, secureFile, secureMkdtemp } from '../tmp';
import type {
  ScreenshotProvider,
  VisualImageFormat,
  VisualPageImage,
  VisualRenderOptions,
  VisualRenderResult,
  VisualTestHooks,
} from './provider';

const execFileAsync = promisify(execFile);

export const VISUAL_DEFAULT_DPI = 150;
export const VISUAL_DEFAULT_JPEG_QUALITY = 90;
export const VISUAL_DEFAULT_MAX_PAGES = 4;
export const VISUAL_MIN_DPI = 72;
export const VISUAL_MAX_DPI = 300;

/** Base artifacts dir: `<base>/ASTra/visual` (base honors $ASTRA_TMPDIR). Created 0700 on demand. */
export function getVisualArtifactsBaseDir(): string {
  const dir = path.join(resolveAstraTmpBase(), 'ASTra', 'visual');
  try {
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  } catch {
    /* best-effort */
  }
  // Repair pre-existing world-readable dirs (0700) without touching parents.
  try {
    secureDir(path.join(resolveAstraTmpBase(), 'ASTra'));
  } catch {
    /* best-effort */
  }
  secureDir(dir);
  return dir;
}

/** Run-scoped artifacts dir: `<base>/<sub>`. Created 0700 on demand. */
export function getVisualArtifactsDir(subDir?: string): string {
  const base = getVisualArtifactsBaseDir();
  if (!subDir || subDir.trim() === '') return base;
  // Contain subDir to a single path segment (no traversal out of base).
  const safe = subDir.trim().replace(/[^a-zA-Z0-9._-]+/g, '-').slice(0, 64) || 'run';
  const dir = path.join(base, safe);
  try {
    fs.mkdirSync(dir, { recursive: true, mode: 0o700 });
  } catch {
    /* best-effort */
  }
  secureDir(dir);
  return dir;
}

function clampInt(n: number, lo: number, hi: number, fallback: number): number {
  if (!Number.isFinite(n)) return fallback;
  const i = Math.floor(n);
  if (!Number.isFinite(i)) return fallback;
  return Math.max(lo, Math.min(hi, i));
}

function normalizePages(pages: unknown, totalPages: number, maxPages: number): number[] | null {
  if (pages === undefined) return null;
  if (!Array.isArray(pages) || pages.length === 0) return null;
  const seen = new Set<number>();
  for (const raw of pages) {
    const n = typeof raw === 'number' ? Math.floor(raw) : Number(raw);
    if (!Number.isFinite(n)) continue;
    const p = Math.max(1, Math.min(totalPages, Math.floor(n)));
    seen.add(p);
  }
  const sorted = [...seen].sort((a, b) => a - b).slice(0, Math.max(1, maxPages));
  return sorted.length > 0 ? sorted : null;
}

function parsePageNumber(fileName: string): number | null {
  const m = fileName.match(/-(\d+)\.(png|jpg|jpeg|ppm|pbm|pgm)$/i);
  if (!m) return null;
  const n = Number(m[1]);
  return Number.isInteger(n) && n >= 1 ? n : null;
}

function listRenderedPages(workDir: string, prefixBase: string): Array<{ page: number; file: string }> {
  let entries: string[] = [];
  try {
    entries = fs.readdirSync(workDir);
  } catch {
    return [];
  }
  const out: Array<{ page: number; file: string }> = [];
  for (const name of entries) {
    if (!name.startsWith(prefixBase)) continue;
    const page = parsePageNumber(name);
    if (page === null) continue;
    out.push({ page, file: path.join(workDir, name) });
  }
  out.sort((a, b) => a.page - b.page);
  return out;
}

function validationError(message: string, remedy: string): Error & { code: string; remedy: string } {
  return Object.assign(new Error(message), { code: 'E_VALIDATION', remedy });
}

function sandboxError(message: string, remedy: string): Error & { code: string; remedy: string } {
  return Object.assign(new Error(message), { code: 'E_SANDBOX', remedy });
}

function canonicalVisualBaseDir(): string {
  const base = getVisualArtifactsBaseDir();
  try {
    return fs.realpathSync(base);
  } catch {
    return base; // just created above; fall back to lexical form
  }
}

function assertVisualInsideBase(canonicalOutDir: string, rawOutDir: string, canonicalBase: string): void {
  if (!isVisualInsideDir(canonicalOutDir, canonicalBase)) {
    throw sandboxError(
      `E_SANDBOX: visual outDir escapes the visual artifacts dir: "${rawOutDir}" (resolves to "${canonicalOutDir}", outside "${canonicalBase}").`,
      'Pass an outDir inside the visual artifacts dir (use getVisualArtifactsDir(subDir) for a run-scoped dir) and retry.'
    );
  }
}

/**
 * Jail a caller-supplied `outDir` to the visual artifacts base dir
 * (`getVisualArtifactsBaseDir()`). Realpath-canonicalizes (symlink-parent
 * escapes resolve outside the base and are refused) and throws `E_SANDBOX`
 * pre-write on any escape. Returns the canonical dir to write under.
 */
export function resolveJailedVisualOutDir(rawOutDir: string): string {
  const trimmed = rawOutDir.trim();
  if (!trimmed) {
    throw sandboxError('E_SANDBOX: visual outDir must be a non-empty string.', 'Pass an outDir inside the visual artifacts dir (use getVisualArtifactsDir(subDir) for a run-scoped dir) and retry.');
  }
  if (trimmed.includes('\0') || /%00/i.test(trimmed)) {
    throw sandboxError('E_SANDBOX: visual outDir contains a null byte.', 'Remove null bytes from the path and retry.');
  }
  if (trimmed.includes('\n') || trimmed.includes('\r')) {
    throw sandboxError('E_SANDBOX: visual outDir contains a forbidden line break.', 'Remove line breaks from the path and retry.');
  }
  const resolved = path.resolve(trimmed);
  const canonical = canonicalizeVisualPath(resolved);
  const canonicalBase = canonicalVisualBaseDir();
  assertVisualInsideBase(canonical, trimmed, canonicalBase);
  return canonical;
}

interface PdfStep {
  pdfPath: string;
  /** Word major version when Word rendered. */
  wordVersion?: number;
  /** Best-effort Word view settings that could not be set (markup only). */
  markupWarnings?: string[];
}

interface RasterStep {
  files: Array<{ page: number; file: string }>;
  totalPages: number;
  /** The rasterizer already applied the page selection (the Windows renderer does); otherwise `normalizePages` runs. */
  selected: boolean;
}

/** DOCX -> PDF with LibreOffice headless (retry for profile lock contention). */
async function renderWithLibreOffice(sofficePath: string, workDir: string, inputDocx: string, timeoutMs: number, markup: boolean): Promise<PdfStep> {
  // The margin-comments option is only sent to a version that was tested: LibreOffice ignores an unknown option.
  const markupWarnings = markup ? (assertLibreOfficeMarkupSupport(sofficePath), libreOfficeMarkupWarnings()) : [];
  // With markup off the literal `pdf` is kept, so the default command line is the one from before.
  const filter = markup ? libreOfficeFilterSpec(true, true) : 'pdf';
  let lastErr: unknown;
  let converted = false;
  for (let attempt = 1; attempt <= 3; attempt++) {
    try {
      await execFileAsync(sofficePath, [`-env:UserInstallation=${pathToFileURL(path.join(workDir, 'lo_profile')).href}`, '--headless', '--convert-to', filter, '--outdir', workDir, inputDocx], {
        timeout: timeoutMs,
      });
      converted = true;
      break;
    } catch (err: unknown) {
      lastErr = err;
      if (attempt < 3) {
        await new Promise((r) => setTimeout(r, 250 * attempt));
      }
    }
  }
  if (!converted && markup) {
    // No silent retry without the option. One diagnostic conversion without it, into its own folder and discarded,
    // lets the message say what this call showed.
    const msg = lastErr instanceof Error ? lastErr.message : String(lastErr);
    throw markupFailureError(msg, await diagnosticPlainRun(sofficePath, workDir, inputDocx, timeoutMs));
  }
  if (!converted) {
    const msg = lastErr instanceof Error ? lastErr.message : String(lastErr);
    throw validationError(`E_VALIDATION: libreoffice PDF conversion failed: ${msg.slice(0, 500)}`, 'Verify the .docx opens in Word/LibreOffice and retry.');
  }
  let pdfPath: string | null = null;
  try {
    const names = fs.readdirSync(workDir).filter((n) => n.toLowerCase().endsWith('.pdf') && n !== WORD_PDF_NAME);
    // Prefer input.pdf; fall back to the first PDF (LO version drift).
    names.sort((a, b) => (a === 'input.pdf' ? -1 : b === 'input.pdf' ? 1 : a.localeCompare(b)));
    if (names.length > 0) pdfPath = path.join(workDir, names[0]);
  } catch {
    pdfPath = null;
  }
  if (!pdfPath || !fs.existsSync(pdfPath)) {
    throw validationError('E_VALIDATION: libreoffice produced no PDF output.', 'Verify the .docx opens in Word/LibreOffice and retry.');
  }
  return { pdfPath, ...(markup ? { markupWarnings } : {}) };
}

const WORD_PDF_NAME = 'word-output.pdf';

/** One conversion without the markup option, for the failure message only. Its output is removed. */
async function diagnosticPlainRun(sofficePath: string, workDir: string, inputDocx: string, timeoutMs: number): Promise<PlainRunOutcome> {
  const dir = path.join(workDir, 'markup-diagnostic');
  try {
    fs.mkdirSync(dir, { recursive: true });
    await execFileAsync(sofficePath, [`-env:UserInstallation=${pathToFileURL(path.join(dir, 'lo_profile')).href}`, '--headless', '--convert-to', 'pdf', '--outdir', dir, inputDocx], { timeout: timeoutMs });
    const made = fs.readdirSync(dir).some((n) => n.toLowerCase().endsWith('.pdf'));
    return made ? { ok: true } : { ok: false, error: 'LibreOffice produced no PDF output' };
  } catch (err: unknown) {
    return { ok: false, error: err instanceof Error ? err.message : String(err) };
  } finally {
    removeDirBestEffort(dir);
  }
}

/** DOCX -> PDF with Word (Windows). A failed attempt leaves no partial PDF behind for the fallback to trip over. */
async function renderWithWord(
  word: WordInstall,
  workDir: string,
  inputDocx: string,
  timeoutMs: number,
  hooks: VisualTestHooks,
  markup: boolean
): Promise<PdfStep> {
  const out = path.join(workDir, WORD_PDF_NAME);
  try {
    // Page images do not need a tagged (accessible) PDF; skipping the structure tree is faster and smaller.
    const outcome = await wordDocxToPdf(inputDocx, out, {
      tagged: false,
      timeoutMs,
      ...(markup ? { markup: true, ...(hooks.wordStartState ? { startState: hooks.wordStartState } : {}) } : {}),
      powershellBin: hooks.powershellBin,
      powershellArgs: hooks.powershellArgs,
      ...(hooks.wordRetryPauseMs ? { retryPauseMs: hooks.wordRetryPauseMs } : {}),
    });
    return { pdfPath: out, wordVersion: word.version, ...(outcome.markup ? { markupWarnings: outcome.markup.warnings } : {}) };
  } catch (err: unknown) {
    try {
      fs.rmSync(out, { force: true });
    } catch {
      /* best-effort */
    }
    throw err;
  }
}

/** PDF -> page images with Poppler's pdftoppm (renders every page; selection happens afterwards). */
async function rasterizeWithPoppler(
  pdftoppmPath: string,
  workDir: string,
  pdfPath: string,
  format: VisualImageFormat,
  dpi: number,
  jpegQuality: number,
  timeoutMs: number
): Promise<RasterStep> {
  const prefixBase = 'page';
  const prefix = path.join(workDir, prefixBase);
  const ppmArgs: string[] =
    format === 'png'
      ? ['-png', '-r', String(dpi), pdfPath, prefix]
      : ['-jpeg', '-jpegopt', `quality=${jpegQuality}`, '-r', String(dpi), pdfPath, prefix];
  try {
    await execFileAsync(pdftoppmPath, ppmArgs, { timeout: Math.min(timeoutMs, 60000) });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw validationError(`E_VALIDATION: pdftoppm rasterization failed: ${msg.slice(0, 500)}`, 'Verify the converted PDF is readable and retry.');
  }
  const rendered = listRenderedPages(workDir, prefixBase);
  if (rendered.length === 0) {
    throw validationError('E_VALIDATION: pdftoppm produced no page images.', 'Verify the document has at least one page and retry.');
  }
  return { files: rendered, totalPages: rendered.length, selected: false };
}

/** PDF -> page images with the Windows built-in renderer (renders only the selected pages, into its own folder). */
async function rasterizeWithWindows(
  workDir: string,
  pdfPath: string,
  format: VisualImageFormat,
  dpi: number,
  jpegQuality: number,
  pages: number[] | undefined,
  maxPages: number,
  timeoutMs: number,
  hooks: VisualTestHooks
): Promise<RasterStep> {
  const dir = path.join(workDir, 'windows-pages');
  fs.mkdirSync(dir, { recursive: true });
  const out = await windowsPdfToImages(pdfPath, dir, {
    format,
    dpi,
    jpegQuality,
    pages,
    maxPages,
    timeoutMs: Math.min(timeoutMs, 60000),
    powershellBin: hooks.powershellBin,
    powershellArgs: hooks.powershellArgs,
  });
  return { files: out.images.sort((a, b) => a.page - b.page), totalPages: out.totalPages, selected: true };
}

/**
 * Render a DOCX buffer to page images.
 *
 * DOCX -> PDF uses Word (Windows; `auto` prefers it there) or LibreOffice; PDF -> images uses the Windows built-in renderer or
 * Poppler. `renderer`, `pdfToImage` and `allowNetworkReferences` (or their environment variables) choose.
 * Throws `E_DEPENDENCY` when no usable program exists (mock PATH in tests by clearing `process.env.PATH` or
 * passing bogus `sofficeBin`/`pdftoppmBin`).
 */
export async function renderDocxToImages(
  docxBuffer: Buffer,
  opts: VisualRenderOptions = {}
): Promise<VisualRenderResult> {
  if (!Buffer.isBuffer(docxBuffer) || docxBuffer.length === 0) {
    throw validationError('E_VALIDATION: docxBuffer must be a non-empty Buffer.', 'Pass the raw .docx bytes and retry.');
  }
  const format: VisualImageFormat = opts.format ?? 'png';
  if (format !== 'png' && format !== 'jpeg') {
    throw validationError(`E_VALIDATION: format must be 'png' or 'jpeg' (got '${String(format)}').`, "Use format 'png' (default, lossless) or opt into 'jpeg'.");
  }
  const dpi = clampInt(opts.dpi ?? VISUAL_DEFAULT_DPI, VISUAL_MIN_DPI, VISUAL_MAX_DPI, VISUAL_DEFAULT_DPI);
  const jpegQuality = clampInt(opts.jpegQuality ?? VISUAL_DEFAULT_JPEG_QUALITY, 1, 100, VISUAL_DEFAULT_JPEG_QUALITY);
  const maxPages = clampInt(opts.maxPages ?? VISUAL_DEFAULT_MAX_PAGES, 1, 32, VISUAL_DEFAULT_MAX_PAGES);
  const timeoutMs = clampInt(opts.timeoutMs ?? 60000, 1000, 300000, 60000);
  const hooks: VisualTestHooks = opts.testHooks ?? {};
  const platform = hooks.platform ?? process.platform;

  // Choices first: an invalid flag, parameter or environment value fails loudly before any temp I/O.
  const rendererPref = effectiveRendererPreference(opts.renderer);
  // An explicit pdftoppm binary means "use that Poppler" unless the caller chose a rasterizer.
  // The environment value is validated even when the binary override makes it moot: a typo must not hide.
  const rasterFromEnvOrParam = effectiveRasterizerPreference(opts.pdfToImage);
  const rasterPref = opts.pdfToImage === undefined && opts.pdftoppmBin && opts.pdftoppmBin.trim() ? 'poppler' : rasterFromEnvOrParam;
  const allowRefs = effectiveAllowNetworkReferences(opts.allowNetworkReferences);
  const showMarkup = effectiveShowMarkup(opts.showMarkup);

  // Capability gate (fail-closed E_DEPENDENCY before any temp I/O).
  const sofficePath = resolveSoffice(opts.sofficeBin);
  const pdftoppmPath = resolvePoppler({ pdftoppmBin: opts.pdftoppmBin });
  const wordMayRun = rendererPref === 'word' || (rendererPref === 'auto' && (hooks.autoPrefersWord ?? AUTO_PREFERS_WORD));
  const word: WordInstall | null = wordMayRun && platform === 'win32' ? (hooks.word !== undefined ? hooks.word : findWord()) : null;
  const rendererPlan = planRenderers(rendererPref, {
    platform,
    word: word !== null,
    libreoffice: sofficePath !== null,
    wordUsable: !wordBreaker.isOpen(),
    autoPrefersWord: hooks.autoPrefersWord,
  });
  const rasterizerPlan = planRasterizers(rasterPref, {
    platform,
    windowsRenderer: hooks.windowsRenderer ?? windowsRendererAvailable({ platform }),
    poppler: pdftoppmPath !== null,
    windowsUsable: !windowsRasterizerBreaker.isOpen(),
  });
  const missing: string[] = [];
  if (rendererPlan.engines.length === 0) missing.push('soffice');
  if (rasterizerPlan.engines.length === 0) missing.push('pdftoppm');
  if (missing.length > 0) throw createDependencyError(missing);

  // Network references: neutralised on the render copy unless the caller allowed them. A failure to check
  // fails the render (fail closed); in allowed mode the scan only reports, so its failure is not fatal.
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

  const runId = `run-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
  // Jail custom outDirs to the visual artifacts base dir (E_SANDBOX pre-write
  // on escape). Default callers (CLI --visual, MCP visual tool) take the
  // run-scoped branch, unchanged.
  let outDir: string;
  let canonicalBase: string | null = null;
  if (opts.outDir && opts.outDir.trim()) {
    outDir = resolveJailedVisualOutDir(opts.outDir);
    canonicalBase = canonicalVisualBaseDir();
  } else {
    outDir = getVisualArtifactsDir(runId);
  }
  try {
    fs.mkdirSync(outDir, { recursive: true });
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    throw validationError(`E_VALIDATION: cannot create artifacts dir "${outDir}": ${msg}`, 'Pass a writable outDir and retry.');
  }
  if (canonicalBase !== null) {
    // Post-mkdir re-assert (a symlink planted at the mkdir races the
    // pre-write check): the live canonical dir must still be inside base.
    const live = canonicalizeVisualPath(outDir);
    assertVisualInsideBase(live, opts.outDir as string, canonicalBase);
    outDir = live;
  }

  const workDir = secureMkdtemp('astra-visual-');
  const inputDocx = path.join(workDir, 'input.docx');
  let succeeded = false;
  try {
    fs.writeFileSync(inputDocx, renderBytes);
    secureFile(inputDocx);

    // 1) DOCX -> PDF.
    const pdfRun = await runPlan(
      rendererPlan.engines,
      { canFallBack: rendererPlan.canFallBack, auto: rendererPlan.requested === 'auto' },
      (engine) =>
        engine === 'word'
          ? renderWithWord(word as WordInstall, workDir, inputDocx, timeoutMs, hooks, showMarkup)
          : renderWithLibreOffice(sofficePath as string, workDir, inputDocx, timeoutMs, showMarkup),
      { word: wordBreaker }
    );
    const pdfPath = pdfRun.value.pdfPath;

    // 2) PDF -> images.
    const rasterRun = await runPlan(
      rasterizerPlan.engines,
      { canFallBack: rasterizerPlan.canFallBack, auto: rasterizerPlan.requested === 'auto' },
      (engine) =>
        engine === 'windows'
          ? rasterizeWithWindows(workDir, pdfPath, format, dpi, jpegQuality, opts.pages, maxPages, timeoutMs, hooks)
          : rasterizeWithPoppler(pdftoppmPath as string, workDir, pdfPath, format, dpi, jpegQuality, timeoutMs),
      { windows: windowsRasterizerBreaker }
    );
    const raster = rasterRun.value;
    const totalPages = raster.totalPages;
    let selected: Array<{ page: number; file: string }>;
    if (raster.selected) {
      selected = raster.files;
    } else {
      const keep = normalizePages(opts.pages, totalPages, maxPages);
      selected = keep === null ? raster.files.slice(0, maxPages) : raster.files.filter((r) => keep.includes(r.page));
    }

    const images: VisualPageImage[] = [];
    const ext = format === 'png' ? 'png' : 'jpg';
    for (const entry of selected) {
      const destName = `page-${entry.page}.${ext}`;
      const dest = path.join(outDir, destName);
      // Pre-write pin: the destination must stay inside outDir.
      if (!isVisualInsideDir(path.resolve(dest), path.resolve(outDir))) {
        throw sandboxError(
          `E_SANDBOX: visual artifact destination escapes outDir: "${dest}".`,
          'Pass an outDir inside the visual artifacts dir (use getVisualArtifactsDir(subDir) for a run-scoped dir) and retry.'
        );
      }
      try {
        fs.copyFileSync(entry.file, dest);
        secureFile(dest);
      } catch (err: unknown) {
        const msg = err instanceof Error ? err.message : String(err);
        throw validationError(`E_VALIDATION: cannot write artifact "${dest}": ${msg}`, 'Pass a writable outDir and retry.');
      }
      // Post-write assert: resolve symlinks at the destination (a planted
      // symlink at dest races the copy) and re-check containment.
      const canonicalDest = canonicalizeVisualPath(dest);
      if (!isVisualInsideDir(canonicalDest, canonicalizeVisualPath(outDir))) {
        try {
          fs.unlinkSync(dest);
        } catch {
          /* best-effort removal of the escaped artifact */
        }
        throw sandboxError(
          `E_SANDBOX: visual artifact destination escapes outDir: "${dest}".`,
          'Pass an outDir inside the visual artifacts dir (use getVisualArtifactsDir(subDir) for a run-scoped dir) and retry.'
        );
      }
      let bytes = 0;
      try {
        bytes = fs.statSync(dest).size;
      } catch {
        bytes = 0;
      }
      images.push({ page: entry.page, path: dest, format, bytes });
    }
    images.sort((a, b) => a.page - b.page);
    const renderer: RendererReport = {
      requested: rendererPlan.requested,
      used: pdfRun.used,
      ...(pdfRun.value.wordVersion !== undefined ? { version: pdfRun.value.wordVersion } : {}),
      ...(pdfRun.fallbackFrom ? { fallbackFrom: pdfRun.fallbackFrom, fallbackReason: pdfRun.fallbackReason } : {}),
    };
    const rasterizer: RasterizerReport = {
      requested: rasterizerPlan.requested,
      used: rasterRun.used,
      ...(rasterRun.fallbackFrom ? { fallbackFrom: rasterRun.fallbackFrom, fallbackReason: rasterRun.fallbackReason } : {}),
    };
    const markup = showMarkup ? buildMarkupReport(pdfRun.used, pdfRun.value.markupWarnings) : undefined;
    const result = { ok: true as const, images, totalPages, outDir, format, backend: pdfRun.used, renderer, rasterizer, references, ...(markup ? { markup } : {}) };
    succeeded = true;
    return result;
  } finally {
    if (succeeded) {
      removeDirBestEffort(workDir);
    } else {
      keepForForensics(workDir, 'visual render failed');
    }
  }
}

/** Default production provider (LibreOffice -> PDF -> pdftoppm). */
export class HeadlessVisualRenderer implements ScreenshotProvider {
  public readonly id = 'headless-lo';
  public async render(docxBuffer: Buffer, opts?: VisualRenderOptions): Promise<VisualRenderResult> {
    return renderDocxToImages(docxBuffer, opts);
  }
}

export const headlessScreenshotProvider: ScreenshotProvider = new HeadlessVisualRenderer();
