/**
 * Screenshot provider hook interface (Sprint 6 Track B).
 *
 * Pure types only — no Node imports so both trees stay IDENTICAL.
 * The engine never hard-codes a renderer: callers inject a provider via
 * `ReviewOptions.screenshotProvider`. The headless LibreOffice renderer
 * (`./renderer.ts`) is the default production provider.
 */

export type VisualImageFormat = 'png' | 'jpeg';

export interface VisualPageImage {
  /** 1-indexed page number. */
  page: number;
  /** Absolute file path to the rendered image (server-side artifact). */
  path: string;
  /** Image encoding. PNG default (lossless glyphs); JPEG opt-in. */
  format: VisualImageFormat;
  /** File size in bytes. */
  bytes: number;
  /** Pixel width when known (pdftoppm output). */
  width?: number;
  /** Pixel height when known. */
  height?: number;
}

export interface VisualRenderOptions {
  /** 1-indexed pages to keep. Default: all pages (capped by maxPages). */
  pages?: number[];
  /** PNG default (lossless); JPEG opt-in for bandwidth-sensitive hosts. */
  format?: VisualImageFormat;
  /** JPEG quality 1..100. Default 90. Ignored for PNG. */
  jpegQuality?: number;
  /** Render DPI. Default 150. Clamped 72..300. */
  dpi?: number;
  /** Artifacts output dir. Default: getVisualArtifactsDir(runId). */
  outDir?: string;
  /** Explicit binary paths (tests use these to simulate missing deps). */
  sofficeBin?: string;
  pdftoppmBin?: string;
  /** Per-binary timeout ms. Default 60000 (soffice) / 30000 (pdftoppm). */
  timeoutMs?: number;
  /** Max pages returned. Default 4 (matches visual_review cap). */
  maxPages?: number;
}

export interface VisualRenderResult {
  ok: true;
  images: VisualPageImage[];
  totalPages: number;
  outDir: string;
  format: VisualImageFormat;
  backend: 'libreoffice';
}

/**
 * Pluggable screenshot provider. Implementations must write image files
 * to disk and return their paths (never inline base64 — token bloat).
 */
export interface ScreenshotProvider {
  readonly id: string;
  render(docxBuffer: Buffer, opts?: VisualRenderOptions): Promise<VisualRenderResult>;
}

/**
 * Engine review options carrying the screenshot hook.
 * `screenshotProvider` unset = telemetry-only mode (no images).
 */
export interface ReviewOptions {
  screenshotProvider?: ScreenshotProvider;
  visualFormat?: VisualImageFormat;
  visualDpi?: number;
}
