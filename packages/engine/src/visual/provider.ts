/**
 * Screenshot provider hook interface (Sprint 6 Track B).
 *
 * Pure types only — no Node imports so both trees stay IDENTICAL.
 * The engine never hard-codes a renderer: callers inject a provider via
 * `ReviewOptions.screenshotProvider`. The headless LibreOffice renderer
 * (`./renderer.ts`) is the default production provider.
 */

import type { RasterizerPreference, RasterizerReport, RendererPreference, RendererReport } from './renderer-preference';
import type { ReferenceReport } from './references';
import type { WordInstall, WordStartState } from './word';
import type { MarkupReport } from './markup';

/**
 * The format of a page image: `png` or `jpeg`.
 */
export type VisualImageFormat = 'png' | 'jpeg';

/**
 * One rendered page: its number, the file that holds the image, the format and the size.
 */
export interface VisualPageImage {
  /** 1-indexed page number. */
  page: number;
  /** The absolute path of the image file. */
  path: string;
  /** The image format. The default is PNG, and JPEG is optional. */
  format: VisualImageFormat;
  /** File size in bytes. */
  bytes: number;
  /** The width in pixels, when it is known. */
  width?: number;
  /** The height in pixels, when it is known. */
  height?: number;
}

/**
 * Options for `renderDocxToImages`: which pages, the format and quality, the resolution, the output folder, the program paths, the time limit, the page limit, the renderer, network references and markup.
 */
export interface VisualRenderOptions {
  /** 1-indexed pages to keep. Default: all pages (capped by maxPages). */
  pages?: number[];
  /** The image format. The default is PNG, which is lossless. JPEG files are smaller. */
  format?: VisualImageFormat;
  /** JPEG quality 1..100. Default 90. Ignored for PNG. */
  jpegQuality?: number;
  /** Render DPI. Default 150. Clamped 72..300. */
  dpi?: number;
  /** The folder for the images. The default is a new folder from `getVisualArtifactsDir`. */
  outDir?: string;
  /** The paths of the LibreOffice (`sofficeBin`) and Poppler (`pdftoppmBin`) programs, when they are not on the PATH. */
  sofficeBin?: string;
  pdftoppmBin?: string;
  /** The time limit for each program, in milliseconds. The default is 60000 for LibreOffice and 30000 for Poppler. */
  timeoutMs?: number;
  /** The most pages to return. The default is 4. */
  maxPages?: number;
  /** DOCX -> PDF program: auto (default), word (Windows only) or libreoffice. Falls back to ASTRA_RENDERER, then auto. */
  renderer?: RendererPreference;
  /** PDF -> image program: auto (default), windows (Windows only) or poppler. Falls back to ASTRA_PDF_TO_IMAGE, then auto. */
  pdfToImage?: RasterizerPreference;
  /**
   * Let the renderer load network locations the document names (default false: such references are
   * retargeted on the temporary copy). An explicit value, including `false`, beats ASTRA_ALLOW_NETWORK_REFERENCES.
   */
  allowNetworkReferences?: boolean;
  /**
   * Draw tracked changes and comments (default false). Falls back to ASTRA_SHOW_MARKUP. A markup render is produced
   * by a different layout mode: pages can be scaled and pagination can change, so page numbers refer to this render.
   */
  showMarkup?: boolean;
  /**
   * @internal Tests only: replace what the machine would report (platform, Word, Windows renderer) and the PowerShell host.
   */
  testHooks?: VisualTestHooks;
}

export interface VisualTestHooks {
  platform?: NodeJS.Platform;
  /** `null` = Word not installed. */
  word?: WordInstall | null;
  windowsRenderer?: boolean;
  powershellBin?: string;
  powershellArgs?: string[];
  /** Pause range between Word retries (ms). */
  wordRetryPauseMs?: [number, number];
  /** Word review-view starting state applied on the render's own instance before the pin (smoke tests). */
  wordStartState?: WordStartState;
  /** Plan `auto` as Word-first (the behaviour a later release may turn on). */
  autoPrefersWord?: boolean;
}

/**
 * What `renderDocxToImages` returns: the page `images`, the total number of pages, the output folder, and reports on the renderer, the references and the markup.
 */
export interface VisualRenderResult {
  ok: true;
  images: VisualPageImage[];
  totalPages: number;
  outDir: string;
  format: VisualImageFormat;
  /** The program that produced the PDF (same as `renderer.used`). */
  backend: 'libreoffice' | 'word';
  renderer?: RendererReport;
  rasterizer?: RasterizerReport;
  /** What the render did about network references in the document. */
  references?: ReferenceReport;
  /** Present only when markup was requested: what was asked and what the renderer was told to do. */
  markup?: MarkupReport;
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
