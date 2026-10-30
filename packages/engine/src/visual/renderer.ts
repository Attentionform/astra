/**
 * Headless document renderer: DOCX -> PDF (LibreOffice) -> PNG/JPEG (pdftoppm).
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
import { assertPdfBackend } from './backend';
import { canonicalizeForAsset as canonicalizeVisualPath, isPathInsideDir as isVisualInsideDir } from '../sandbox/paths';
import { keepForForensics, removeDirBestEffort, resolveAstraTmpBase, secureDir, secureFile, secureMkdtemp } from '../tmp';
import type {
  ScreenshotProvider,
  VisualImageFormat,
  VisualPageImage,
  VisualRenderOptions,
  VisualRenderResult,
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

/**
 * Render a DOCX buffer to page images.
 * Throws `E_DEPENDENCY` when LibreOffice/pdftoppm are missing (mock PATH in
 * tests by clearing `process.env.PATH` or passing bogus `sofficeBin`).
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

  // Capability gate FIRST (fail-closed E_DEPENDENCY before any temp I/O).
  const { sofficePath, pdftoppmPath } = assertPdfBackend({
    sofficeBin: opts.sofficeBin,
    pdftoppmBin: opts.pdftoppmBin,
  });

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
    fs.writeFileSync(inputDocx, docxBuffer);
    secureFile(inputDocx);

    // 1) DOCX -> PDF via LibreOffice headless (retry for profile lock contention).
    let lastErr: unknown;
    let converted = false;
    for (let attempt = 1; attempt <= 3; attempt++) {
      try {
        await execFileAsync(sofficePath, [`-env:UserInstallation=${pathToFileURL(path.join(workDir, 'lo_profile')).href}`, '--headless', '--convert-to', 'pdf', '--outdir', workDir, inputDocx], {
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
    if (!converted) {
      const msg = lastErr instanceof Error ? lastErr.message : String(lastErr);
      throw validationError(`E_VALIDATION: libreoffice PDF conversion failed: ${msg.slice(0, 500)}`, 'Verify the .docx opens in Word/LibreOffice and retry.');
    }
    let pdfPath: string | null = null;
    try {
      const names = fs.readdirSync(workDir).filter((n) => n.toLowerCase().endsWith('.pdf'));
      // Prefer input.pdf; fall back to the first PDF (LO version drift).
      names.sort((a, b) => (a === 'input.pdf' ? -1 : b === 'input.pdf' ? 1 : a.localeCompare(b)));
      if (names.length > 0) pdfPath = path.join(workDir, names[0]);
    } catch {
      pdfPath = null;
    }
    if (!pdfPath || !fs.existsSync(pdfPath)) {
      throw validationError('E_VALIDATION: libreoffice produced no PDF output.', 'Verify the .docx opens in Word/LibreOffice and retry.');
    }

    // 2) PDF -> images via pdftoppm.
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
    const totalPages = rendered.length;
    const keep = normalizePages(opts.pages, totalPages, maxPages);
    const selected = keep === null ? rendered.slice(0, maxPages) : rendered.filter((r) => keep.includes(r.page));

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
    const result = { ok: true as const, images, totalPages, outDir, format, backend: 'libreoffice' as const };
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
