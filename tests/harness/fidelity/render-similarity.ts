/**
 * Render-similarity scorer and oracle trust boundary (advisory only).
 *
 * Nothing in this file blocks a merge. Every consumer treats the output as
 * data: a per-fixture advisory entry that is either `measured` (scores
 * recorded) or `unavailable` (the renderer could not run; reason recorded).
 * There is deliberately no pass/fail verdict here.
 *
 * Metric (what is compared, threshold, why):
 * - Render: the before/after `.docx` bytes are each converted to PDF with the
 *   same local LibreOffice binary (`soffice --headless --convert-to pdf`),
 *   each run under an isolated throwaway profile so parallel workers never
 *   share a profile lock. PDF bytes are NEVER compared: the exporter stamps
 *   CreationDate/producer counters, so byte equality is meaningless.
 * - Structural projection (the oracle output): page count from `pdfinfo` plus
 *   layout text from `pdftotext -layout`. Page-count equality and token
 *   Jaccard similarity are reported per fixture.
 * - Pixel score: each PDF page is rasterized with `pdftoppm -gray -r 36` and
 *   parsed as PGM (no image dependency). Per page the score is
 *   mean(|a-b|)/255 over the shared area; pages beyond the shorter document
 *   are reported through the page-count flag, not folded into the score. The
 *   fixture score is the mean over shared pages; the worst page is reported
 *   alongside it so one broken page cannot hide behind clean siblings.
 * - Thresholds (advisory bands, not gates): pixel <= 0.02 `consistent`,
 *   <= 0.10 `localized-change`, above that or on page-count change
 *   `needs-eyes`. Rationale: the pipeline mutates two sentences, which at
 *   36 dpi (a letter page is 306x396 px) moves roughly one to three percent
 *   of pixels on the touched pages; the bands keep that expected signal
 *   below the human-review line while a reflowed table or dropped page lands
 *   far above it. The bands were set against the measured distribution over
 *   the corpus, not guessed (see the variance report for the calibration
 *   campaign).
 * - Renderer-noise tolerance by design: a missing binary, a conversion
 *   timeout, a renderer crash, or an unrasterizable PDF yields `unavailable`
 *   with the reason, never a failure. Version drift is attributable, not
 *   silent: every report records the LibreOffice version string.
 *
 * Trust boundary (borrowed shape): the oracle compares structural
 * projections (page count, extracted text, pixel score) only. It never
 * asserts what the engine must emit, never reads engine internals, and is
 * binary-gated: without a working renderer on PATH the whole advisory suite
 * skips cleanly. Caveats: 36 dpi grayscale rasters are blind to sub-point
 * shifts and to color-only changes; `pdftotext` order follows the PDF
 * content stream, not visual reading order, so bidi/complex-script text
 * similarity is approximate; the renderer has known-bad inputs (shapes that
 * crash the exporter are recorded `unavailable`, see the findings log).
 */

import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';

/** Render resolution for rasters. Keeps a letter page at 306x396 px. */
export const RENDER_DPI = 36;
/** Per-conversion ceiling; exceeding it records `unavailable`, never fails. */
export const CONVERT_TIMEOUT_MS = 90_000;
/** Pixel score at or under this is routine mutation noise. */
export const BAND_CONSISTENT = 0.02;
/** Pixel score at or under this is a plausible localized change. */
export const BAND_LOCALIZED = 0.10;
/**
 * Promotion budget: identical-input self-variance must stay under this for
 * three consecutive full campaigns before render similarity may gate merges.
 */
export const VARIANCE_BUDGET = 0.005;

export type AdvisoryBand = 'consistent' | 'localized-change' | 'needs-eyes';

export interface FixtureAdvisory {
  fixture: string;
  status: 'measured' | 'unavailable';
  /** Renderer version that produced the scores (or that was probed). */
  renderer: string | null;
  reason?: string;
  pagesBefore?: number;
  pagesAfter?: number;
  pageCountEqual?: boolean;
  /** Mean pixel score over shared pages. */
  pixelMean?: number;
  /** Worst shared page score. */
  pixelWorstPage?: number;
  /** Whitespace-token Jaccard similarity of layout text. */
  textSimilarity?: number;
  band?: AdvisoryBand;
}

/** Thrown when the renderer cannot produce output. Never a test failure. */
export class OracleUnavailableError extends Error {
  constructor(message: string) {
    super(message);
    this.name = 'OracleUnavailableError';
  }
}

const BIN_CANDIDATES = ['soffice', 'libreoffice'];

function toolOnPath(bin: string): boolean {
  const probe = spawnSync('sh', ['-c', `command -v ${bin}`], { encoding: 'utf8' });
  return probe.status === 0 && (probe.stdout ?? '').trim().length > 0;
}

/** Renderer binary name when present, null when the oracle must skip. */
export function oracleBinary(): string | null {
  for (const bin of BIN_CANDIDATES) {
    if (toolOnPath(bin)) return bin;
  }
  return null;
}

/** True when every external renderer tool the scorer needs is on PATH. */
export function isOracleAvailable(): boolean {
  if (oracleBinary() === null) return false;
  for (const bin of ['pdfinfo', 'pdftotext', 'pdftoppm']) {
    if (!toolOnPath(bin)) return false;
  }
  return true;
}

/** Version string for attribution in reports, null when not runnable. */
export function libreOfficeVersion(): string | null {
  const bin = oracleBinary();
  if (bin === null) return null;
  const probe = spawnSync(bin, ['--version'], { encoding: 'utf8', timeout: 30_000 });
  const line = (probe.stdout ?? '').split('\n')[0]?.trim() ?? '';
  return line.length > 0 ? line : null;
}

export interface ConvertOptions {
  timeoutMs?: number;
  /**
   * Caller-owned renderer profile directory. Reusing one directory across
   * sequential conversions in a process amortizes first-run setup (a fresh
   * profile costs ~20 s once, later conversions under a second) and keeps
   * parallel workers isolated when each owns its directory. Omit for a
   * throwaway profile (slow, but free of lifecycle obligations). A crashed
   * conversion does not poison a shared directory: the next conversion
   * reuses it normally (verified against the exporter crash on endnote
   * shapes).
   */
  profileDir?: string;
}

/**
 * Create an isolated renderer profile directory owned by the caller.
 * The caller removes it when done. One per test worker is the intended
 * granularity: sequential conversions share it, parallel workers never do.
 */
export function createIsolatedProfile(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'astra-lo-prof-'));
}

/**
 * Convert one `.docx` to PDF in `outDir` under an isolated renderer profile.
 * Returns the PDF path. Throws OracleUnavailableError on a missing binary,
 * a timeout, a nonzero exit, or a missing output file (renderer crash).
 */
export function convertDocxToPdf(docxPath: string, outDir: string, opts: ConvertOptions = {}): string {
  const bin = oracleBinary();
  if (bin === null) {
    throw new OracleUnavailableError('no renderer binary on PATH (soffice/libreoffice)');
  }
  const stem = path.basename(docxPath, '.docx');
  const ownedProfile = opts.profileDir === undefined;
  const profileDir = ownedProfile ? createIsolatedProfile() : (opts.profileDir as string);
  try {
    const run = spawnSync(
      bin,
      [
        '--headless',
        `-env:UserInstallation=file://${profileDir}`,
        '--convert-to',
        'pdf',
        '--outdir',
        outDir,
        docxPath,
      ],
      { encoding: 'utf8', timeout: opts.timeoutMs ?? CONVERT_TIMEOUT_MS },
    );
    if (run.error !== undefined) {
      const code = (run.error as NodeJS.ErrnoException).code;
      throw new OracleUnavailableError(
        code === 'ETIMEDOUT' ? `conversion timed out for ${stem}` : `conversion spawn failed for ${stem}: ${String(run.error.message ?? run.error).slice(0, 160)}`,
      );
    }
    const pdfPath = path.join(outDir, `${stem}.pdf`);
    if (run.status !== 0 || !fs.existsSync(pdfPath)) {
      const tail = `${run.stdout ?? ''}\n${run.stderr ?? ''}`.trim().split('\n').pop() ?? '';
      throw new OracleUnavailableError(
        `renderer exit=${run.status ?? '?'} without PDF for ${stem}: ${tail.slice(0, 160)}`,
      );
    }
    return pdfPath;
  } finally {
    if (ownedProfile) fs.rmSync(profileDir, { recursive: true, force: true });
  }
}

/** Page count from pdfinfo. Throws OracleUnavailableError when unreadable. */
export function pdfPageCount(pdfPath: string): number {
  const run = spawnSync('pdfinfo', [pdfPath], { encoding: 'utf8', timeout: 30_000 });
  const match = /Pages:\s+(\d+)/.exec(run.stdout ?? '');
  if (run.status !== 0 || match === null) {
    throw new OracleUnavailableError(`pdfinfo could not read ${path.basename(pdfPath)}`);
  }
  return Number(match[1]);
}

/** Layout text via pdftotext. Throws OracleUnavailableError when unreadable. */
export function pdfLayoutText(pdfPath: string): string {
  const run = spawnSync('pdftotext', ['-layout', pdfPath, '-'], { encoding: 'utf8', timeout: 60_000 });
  if (run.status !== 0) {
    throw new OracleUnavailableError(`pdftotext could not read ${path.basename(pdfPath)}`);
  }
  return run.stdout ?? '';
}

export interface RasterPage {
  width: number;
  height: number;
  pixels: Buffer;
}

/** Parse a binary P5 PGM (the pdftoppm -gray output). Pure, no dependencies. */
export function parsePgm(buf: Buffer): RasterPage {
  let pos = 0;
  const readToken = (): string => {
    while (pos < buf.length) {
      const ch = buf[pos];
      if (ch === 0x23) {
        while (pos < buf.length && buf[pos] !== 0x0a) pos++;
        continue;
      }
      if (ch === 0x20 || ch === 0x0a || ch === 0x0d || ch === 0x09) {
        pos++;
        continue;
      }
      break;
    }
    const start = pos;
    while (pos < buf.length && ![0x20, 0x0a, 0x0d, 0x09].includes(buf[pos])) pos++;
    return buf.subarray(start, pos).toString('ascii');
  };
  const magic = readToken();
  if (magic !== 'P5') throw new Error(`not a binary PGM (magic ${JSON.stringify(magic)})`);
  const width = Number(readToken());
  const height = Number(readToken());
  const maxval = Number(readToken());
  if (!Number.isInteger(width) || width <= 0 || !Number.isInteger(height) || height <= 0 || maxval !== 255) {
    throw new Error(`unsupported PGM header ${width}x${height}/${maxval}`);
  }
  pos += 1; // single whitespace byte before the raster
  const pixels = buf.subarray(pos, pos + width * height);
  if (pixels.length !== width * height) throw new Error('PGM raster truncated');
  return { width, height, pixels: Buffer.from(pixels) };
}

/**
 * Rasterize every PDF page to grayscale at RENDER_DPI. Returns one entry per
 * page in order. Throws OracleUnavailableError when rasterization fails.
 */
export function renderPdfToRasters(pdfPath: string, workDir: string): RasterPage[] {
  const prefix = path.join(workDir, 'page');
  const run = spawnSync('pdftoppm', ['-gray', '-r', String(RENDER_DPI), pdfPath, prefix], {
    encoding: 'utf8',
    timeout: 120_000,
  });
  if (run.status !== 0) {
    throw new OracleUnavailableError(`pdftoppm could not rasterize ${path.basename(pdfPath)}`);
  }
  const pages: RasterPage[] = [];
  for (let n = 1; ; n++) {
    const pgm = `${prefix}-${n}.pgm`;
    if (!fs.existsSync(pgm)) break;
    pages.push(parsePgm(fs.readFileSync(pgm)));
  }
  if (pages.length === 0) {
    throw new OracleUnavailableError(`pdftoppm produced no pages for ${path.basename(pdfPath)}`);
  }
  return pages;
}

export interface PixelDiff {
  mean: number;
  /** Fraction of pixels differing by more than 16 gray levels. */
  coarseFraction: number;
  dimMismatch: boolean;
}

/** Normalized mean absolute pixel difference. Pure. */
export function pixelDifference(a: RasterPage, b: RasterPage): PixelDiff {
  if (a.width !== b.width || a.height !== b.height) {
    return { mean: 1, coarseFraction: 1, dimMismatch: true };
  }
  let sum = 0;
  let coarse = 0;
  const n = a.pixels.length;
  for (let i = 0; i < n; i++) {
    const d = Math.abs(a.pixels[i] - b.pixels[i]);
    sum += d;
    if (d > 16) coarse++;
  }
  return { mean: sum / n / 255, coarseFraction: coarse / n, dimMismatch: false };
}

/** Whitespace-token Jaccard similarity in [0, 1]. Pure. */
export function tokenSimilarity(aText: string, bText: string): number {
  const tokens = (s: string): Set<string> => new Set(s.split(/\s+/).filter((t) => t.length > 0));
  const a = tokens(aText);
  const b = tokens(bText);
  if (a.size === 0 && b.size === 0) return 1;
  let inter = 0;
  for (const t of a) if (b.has(t)) inter++;
  return inter / (a.size + b.size - inter);
}

/** Advisory band from a pixel score and page-count equality. Pure. */
export function advisoryBand(pixelMean: number, pageCountEqual: boolean): AdvisoryBand {
  if (!pageCountEqual || pixelMean > BAND_LOCALIZED) return 'needs-eyes';
  if (pixelMean > BAND_CONSISTENT) return 'localized-change';
  return 'consistent';
}

export interface SelfVariance {
  runs: number;
  pages: number[];
  pagesStable: boolean;
  textStable: boolean;
  /** Largest mean pixel score over any pair of runs. */
  maxPixelMean: number;
}

/**
 * Convert identical input bytes `runs` times and compare the projections
 * pairwise. Each run converts in its own directory with its own renderer
 * profile, so profile or output reuse cannot hide variance. Temporary
 * directories are removed as the study proceeds.
 */
export function measureSelfVariance(
  docxPath: string,
  runs: number,
  workRoot: string,
  profileDir?: string,
): SelfVariance {
  const stem = path.basename(docxPath, '.docx');
  const rasters: RasterPage[][] = [];
  const pages: number[] = [];
  const texts: string[] = [];
  for (let i = 0; i < runs; i++) {
    const dir = fs.mkdtempSync(path.join(workRoot, `astra-variance-${stem}-${i}-`));
    try {
      const input = path.join(dir, `${stem}.docx`);
      fs.copyFileSync(docxPath, input);
      const pdf = convertDocxToPdf(input, dir, { profileDir });
      pages.push(pdfPageCount(pdf));
      rasters.push(renderPdfToRasters(pdf, dir));
      texts.push(pdfLayoutText(pdf));
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  }
  let maxPixelMean = 0;
  for (let a = 0; a < runs; a++) {
    for (let b = a + 1; b < runs; b++) {
      const shared = Math.min(rasters[a].length, rasters[b].length);
      let sum = 0;
      for (let p = 0; p < shared; p++) sum += pixelDifference(rasters[a][p], rasters[b][p]).mean;
      const score = shared === 0 ? 1 : sum / shared;
      if (score > maxPixelMean) maxPixelMean = score;
    }
  }
  return {
    runs,
    pages,
    pagesStable: pages.every((n) => n === pages[0]),
    textStable: texts.every((t) => t === texts[0]),
    maxPixelMean,
  };
}
