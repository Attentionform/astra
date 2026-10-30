import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import {
  DocxParser,
  buildVisualTelemetry,
  renderDocxToImages,
  VISUAL_PROMPT_LAW,
  type DocumentAST,
  type VisualImageFormat,
  type VisualTelemetry,
} from '@attentionform/astra-engine';
import { EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import {
  macroFlagsUsageError,
  macroOptsFromFlags,
  parseMaxRatioFlag,
  parseMaxUncompressedBytesFlag,
  readDocx,
  withZipCeilingEnv,
  type ZipCeilingOverrides,
} from './io';

/**
 * Single CLI visual-telemetry builder (two entry points: `visual` below and
 * `inspect --visual` in inspect-cmd.ts — both call this, never the engine
 * directly, so the wireframe/warnings stay identical across surfaces).
 * `renderTotalPages` is the renderer's (LibreOffice) actual page count: the
 * single source of truth for the `[PAGE n OF m]` marker whenever it differs
 * from the word-layout map embedded in the structural summary.
 */
export function buildCliVisualTelemetry(ast: DocumentAST, summary: unknown, page = 1, renderTotalPages?: number): VisualTelemetry {
  const nodePages = (summary as unknown as { nodePages?: Record<string, number> })?.nodePages;
  const totalPages = (summary as unknown as { totalPages?: number })?.totalPages;
  const pageMap =
    nodePages && typeof totalPages === 'number'
      ? { totalPages, nodePages, source: 'renderer' as const, isEstimate: false }
      : undefined;
  return buildVisualTelemetry(ast, { page, pageMap, renderTotalPages });
}

/**
 * Single CLI visual-render path (two entry points: `visual --out` below and
 * `inspect --visual` in inspect-cmd.ts — both call this, never
 * `renderDocxToImages` directly, so DPI/quality/page-selection stay
 * identical: same `pages: [page]` + `format` args on every invocation).
 */
export function buildCliVisualRenderArgs(page: number, format: VisualImageFormat): { pages: [number]; format: VisualImageFormat } {
  return { pages: [page], format };
}

export async function renderCliVisualPage(
  buf: Buffer,
  page: number,
  format: VisualImageFormat,
): Promise<Awaited<ReturnType<typeof renderDocxToImages>>> {
  const args = buildCliVisualRenderArgs(page, format);
  return renderDocxToImages(buf, args);
}

/**
 * Single CLI visual-telemetry printer (both entry points call this, so the
 * wireframe/warnings/tokens lines stay identical across surfaces).
 * Callers pass the EFFECTIVE page (the renderer's clamped pick, not the raw
 * request) so the `[PAGE n OF m]` marker always describes the image shown,
 * and the renderer's actual page count as `renderTotalPages` so `OF m` is
 * the renderer's count — the image came from the renderer, the marker must
 * name it. Divergence from the word-layout page map and any --page beyond
 * that map warn on stderr (never silent; stdout stays JSON-clean).
 */
export function printCliVisualTelemetry(
  out: GlobalOpts,
  ast: DocumentAST,
  summary: unknown,
  page: number,
  renderTotalPages?: number,
): VisualTelemetry {
  const telemetry = buildCliVisualTelemetry(ast, summary, page, renderTotalPages);
  print(out, telemetry.wireframe);
  if (telemetry.warnings.length > 0) {
    for (const w of telemetry.warnings.slice(0, 10)) print(out, `visual: ${w}`);
  }
  print(out, `visual: tokens~${telemetry.tokens} budget=500 law="${VISUAL_PROMPT_LAW.slice(0, 80)}..."`);
  // Honest-pagination stderr warnings (both entry points, always emitted —
  // these survive --json so machine consumers see them without stdout mixing).
  if (telemetry.wordLayoutTotalPages !== undefined) {
    // The "nodes beyond the map" clause only has content when the renderer
    // paginated past the word-layout map; on a reverse mismatch (renderer
    // stopped before the map's last page) nothing exists beyond the map,
    // so the clause is dropped as vacuous.
    const beyondMapClause =
      telemetry.totalPages > telemetry.wordLayoutTotalPages
        ? `, wireframe nodes beyond page ${telemetry.wordLayoutTotalPages} are unmapped`
        : '';
    printStderr(
      `visual: pagination mismatch: renderer (LibreOffice) produced ${pageCountLabel(telemetry.totalPages)}; word-layout page map totals ${telemetry.wordLayoutTotalPages} — [PAGE n OF m] uses the renderer count${beyondMapClause}`,
    );
  }
  const wordTotal = wordLayoutTotalFromSummary(summary);
  if (page > wordTotal) {
    printStderr(
      `visual: --page ${page} exceeds the word-layout page map (${pageCountLabel(wordTotal)}) — wireframe node selection is unmapped for it`,
    );
  }
  return telemetry;
}

/** Singular/plural label for page counts in stderr messages. */
function pageCountLabel(count: number): string {
  return count === 1 ? '1 page' : `${count} pages`;
}

/**
 * Word-layout page total from the structural summary (same derivation the
 * pageMap above uses; defensive: missing/invalid falls back to 1 page).
 */
function wordLayoutTotalFromSummary(summary: unknown): number {
  const totalPages = (summary as unknown as { totalPages?: number })?.totalPages;
  return typeof totalPages === 'number' && Number.isFinite(totalPages) && totalPages >= 1
    ? Math.floor(totalPages)
    : 1;
}

export async function cmdVisual(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'visual');
  let file: string | undefined;
  let pageRaw: string | undefined;
  let outPath: string | undefined;
  let formatRaw: string | undefined;
  let maxRatioRaw: string | undefined;
  let maxBytesRaw: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  let force = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--page') {
      const value = args[++i];
      if (value === undefined) return usage(out, 'visual --page requires a value (integer >= 1)');
      pageRaw = value;
    } else if (a.startsWith('--page=')) pageRaw = a.slice('--page='.length);
    else if (a === '--out' || a === '--output' || a === '-o') outPath = args[++i];
    else if (a.startsWith('--out=')) outPath = a.slice('--out='.length);
    else if (a.startsWith('--output=')) outPath = a.slice('--output='.length);
    else if (a.startsWith('-o=')) outPath = a.slice('-o='.length);
    else if (a === '--format') {
      const value = args[++i];
      if (value === undefined) return usage(out, 'visual --format requires a value (png|jpeg)');
      formatRaw = value;
    } else if (a.startsWith('--format=')) formatRaw = a.slice('--format='.length);
    else if (a === '--max-ratio') {
      const value = args[++i];
      if (value === undefined) return usage(out, 'visual --max-ratio requires a value (positive number, e.g. --max-ratio 200)');
      maxRatioRaw = value;
    } else if (a.startsWith('--max-ratio=')) maxRatioRaw = a.slice('--max-ratio='.length);
    else if (a === '--max-uncompressed-bytes') maxBytesRaw = args[++i];
    else if (a.startsWith('--max-uncompressed-bytes=')) maxBytesRaw = a.slice('--max-uncompressed-bytes='.length);
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a === '--force') force = true;
    else if (a.startsWith('--')) return usage(out, `visual: unknown flag "${a}"`);
    else if (!file) file = a;
    else return usage(out, 'visual takes exactly one <docx> argument');
  }
  if (!file) return usage(out, 'visual requires <docx>');
  if (!outPath) return usage(out, 'visual requires --out <render.png>');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  const ceilings: ZipCeilingOverrides = {};
  try {
    if (maxRatioRaw !== undefined) ceilings.maxRatio = parseMaxRatioFlag(maxRatioRaw);
    if (maxBytesRaw !== undefined) ceilings.maxUncompressedBytes = parseMaxUncompressedBytesFlag(maxBytesRaw);
  } catch (err: unknown) {
    return usage(out, err instanceof Error ? err.message : String(err));
  }
  let page = 1;
  if (pageRaw !== undefined) {
    const n = Number(pageRaw);
    if (!Number.isInteger(n) || n < 1) {
      return usage(out, 'visual --page must be an integer >= 1');
    }
    page = n;
  }
  if (formatRaw !== undefined && formatRaw !== 'png' && formatRaw !== 'jpeg') {
    return usage(out, 'visual --format must be png|jpeg');
  }
  const format = formatRaw === 'jpeg' ? 'jpeg' : 'png';
  if (outPath !== undefined && !force) {
    try {
      await fsp.access(outPath);
      return usage(out, `visual: --out target "${path.basename(outPath)}" already exists (refusing to overwrite without --force)`);
    } catch {
      /* destination is free — proceed */
    }
  }
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  try {
    // Alias-real parser construction: same macroOptsFromFlags +
    // withZipCeilingEnv wiring as `inspect` (flag ceilings scope the
    // DocxParser zip-bomb gate for this load).
    const macroOpts = macroOptsFromFlags({ allowMacros, stripMacros });
    const { parser, ast } = await withZipCeilingEnv(ceilings, async () => {
      const p = new DocxParser(buf, macroOpts);
      return { parser: p, ast: await p.getAST() };
    });
    const summary = parser.getStructuralSummary();
    // Render first, then describe the EFFECTIVE page: an over-range --page
    // clamps to the last rendered page, and the telemetry marker must name
    // the image actually written (never the unclamped request). The
    // renderer's page count (`rendered.totalPages`) is the single source of
    // truth for the marker's `OF m` — the word-layout map can disagree on
    // real docs, and the printer warns on stderr when it does.
    let rendered: Awaited<ReturnType<typeof renderDocxToImages>>;
    try {
      rendered = await renderCliVisualPage(buf, page, format);
    } catch (err: unknown) {
      return failHuman(out, err, { file, fallbackCode: 'E_DEPENDENCY' });
    }
    const picked = rendered.images.find((img) => img.page === page) ?? rendered.images[0];
    if (!picked) {
      return failHuman(out, 'visual: renderer produced no page images', { file, fallbackCode: 'E_VALIDATION' });
    }
    const telemetry = printCliVisualTelemetry(out, ast, summary, picked.page, rendered.totalPages);
    try {
      await fsp.mkdir(path.dirname(path.resolve(outPath)), { recursive: true });
      if (!force && fs.existsSync(outPath)) {
        return usage(out, `visual: --out target "${path.basename(outPath)}" already exists (refusing to overwrite without --force)`);
      }
      await fsp.copyFile(picked.path, outPath);
    } catch (err: unknown) {
      return failHuman(out, err, { file: outPath, fallbackCode: 'E_IO' });
    }
    let bytes = picked.bytes;
    try {
      bytes = (await fsp.stat(outPath)).size;
    } catch {
      /* keep the renderer-reported size */
    }
    print(out, `VISUAL OK ${file} page ${picked.page} -> ${outPath} (${format.toUpperCase()}, ${bytes} bytes)`);
    printJson(out, {
      ok: true,
      command: 'visual',
      file,
      page: picked.page,
      format,
      path: outPath,
      bytes,
      // Guard: never emit `totalPages: 0` — the renderer contract is >=1 page
      // (it throws otherwise); fall back to the telemetry total if violated.
      totalPages: rendered.totalPages >= 1 ? Math.floor(rendered.totalPages) : telemetry.totalPages,
      telemetry,
      promptLaw: VISUAL_PROMPT_LAW,
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
