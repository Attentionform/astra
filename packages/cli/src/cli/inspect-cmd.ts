import { DocxParser, ToonGenerator, VISUAL_PROMPT_LAW } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_OK, type Fidelity, type GlobalOpts, print, printJson, printStderr, resolveInspectWidth } from './output';
import { cleanCliMessage, failHuman, rewriteMcpRemedies } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { computeDocStats, macroFlagsUsageError, macroOptsFromFlags, parseMaxRatioFlag, parseMaxUncompressedBytesFlag, readDocx, withZipCeilingEnv, type ZipCeilingOverrides } from './io';
import { printCliVisualTelemetry, renderCliVisualPage } from './visual-cmd';

export function parseFidelity(raw: string | undefined): Fidelity | undefined {
  if (raw === undefined) return undefined;
  if (raw === 'compact' || raw === 'standard' || raw === 'full') return raw;
  return undefined;
}

export async function cmdInspect(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'inspect');
  let file: string | undefined;
  let fidelityRaw: string | undefined;
  let sectionId: string | undefined;
  let maxTextRaw: string | undefined;
  let widthRaw: string | undefined;
  let wrap = false;
  let withVisual = false;
  let pageRaw: string | undefined;
  let formatRaw: string | undefined;
  let maxRatioRaw: string | undefined;
  let maxBytesRaw: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--fidelity') fidelityRaw = args[++i];
    else if (a === '--section') sectionId = args[++i];
    else if (a === '--max-text') maxTextRaw = args[++i];
    else if (a === '--width') widthRaw = args[++i];
    else if (a.startsWith('--width=')) widthRaw = a.slice('--width='.length);
    else if (a === '--wrap') wrap = true;
    else if (a === '--visual') withVisual = true;
    else if (a === '--page') {
      const value = args[++i];
      if (value === undefined) return usage(out, 'inspect --page requires a value (integer >= 1)');
      pageRaw = value;
    } else if (a.startsWith('--page=')) pageRaw = a.slice('--page='.length);
    else if (a === '--format') {
      const value = args[++i];
      if (value === undefined) return usage(out, 'inspect --format requires a value (png|jpeg)');
      formatRaw = value;
    } else if (a === '--max-ratio') {
      const value = args[++i];
      if (value === undefined) return usage(out, 'inspect --max-ratio requires a value (positive number, e.g. --max-ratio 200)');
      maxRatioRaw = value;
    } else if (a.startsWith('--max-ratio=')) maxRatioRaw = a.slice('--max-ratio='.length);
    else if (a === '--max-uncompressed-bytes') maxBytesRaw = args[++i];
    else if (a.startsWith('--max-uncompressed-bytes=')) maxBytesRaw = a.slice('--max-uncompressed-bytes='.length);
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a.startsWith('--')) return usage(out, `inspect: unknown flag "${a}"`);
    else if (!file) file = a;
    else return usage(out, 'inspect takes exactly one <docx> argument');
  }
  if (!file) return usage(out, 'inspect requires <docx>');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  const ceilings: ZipCeilingOverrides = {};
  try {
    if (maxRatioRaw !== undefined) ceilings.maxRatio = parseMaxRatioFlag(maxRatioRaw);
    if (maxBytesRaw !== undefined) ceilings.maxUncompressedBytes = parseMaxUncompressedBytesFlag(maxBytesRaw);
  } catch (err: unknown) {
    return usage(out, err instanceof Error ? err.message : String(err));
  }
  const fidelity = parseFidelity(fidelityRaw);
  if (fidelityRaw !== undefined && fidelity === undefined) {
    return usage(out, 'inspect --fidelity must be compact|standard|full');
  }
  if (formatRaw !== undefined && formatRaw !== 'png' && formatRaw !== 'jpeg') {
    return usage(out, 'inspect --format must be png|jpeg');
  }
  const visualFormat = formatRaw === 'jpeg' ? 'jpeg' : 'png';
  let visualPage = 1;
  if (pageRaw !== undefined) {
    const n = Number(pageRaw);
    if (!Number.isInteger(n) || n < 1) {
      return usage(out, 'inspect --page must be an integer >= 1');
    }
    visualPage = n;
  }
  let maxTextLength: number | undefined;
  if (maxTextRaw !== undefined) {
    const n = Number(maxTextRaw);
    if (!Number.isInteger(n) || n < 1 || n > 10000) {
      return usage(out, 'inspect --max-text must be an integer 1..10000');
    }
    maxTextLength = n;
  }
  let widthExplicit: number | undefined;
  if (widthRaw !== undefined) {
    const n = Number(widthRaw);
    if (!Number.isInteger(n) || n < 1 || n > 10000) {
      return usage(out, 'inspect --width must be an integer 1..10000');
    }
    widthExplicit = n;
  }
  const width = resolveInspectWidth({ width: widthExplicit, wrap });
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  try {
    // R-8: flag ceilings scope the DocxParser zip-bomb gate for this load.
    const macroOpts = macroOptsFromFlags({ allowMacros, stripMacros });
    const { parser, ast } = await withZipCeilingEnv(ceilings, async () => {
      const p = new DocxParser(buf, macroOpts);
      return { parser: p, ast: await p.getAST() };
    });
    const toon = new ToonGenerator();
    const table = toon.toon(ast, { fidelity, sectionId, maxTextLength });
    // Wrap is human presentation only: human stdout uses the wrapped form,
    // the machine `toon` field below always uses the canonical unwrapped
    // single-line-per-node form. Wrap breaks only at whitespace runs and
    // rejoins with single spaces, so text with multi-space/tab runs does not
    // unwrap byte-identically (single-space text round-trips exactly via
    // unwrapToonText).
    const canonicalText = toon.renderToonText(table);
    const text = width === undefined ? canonicalText : toon.renderToonText(table, { width });
    const summary = parser.getStructuralSummary();
    const stats = computeDocStats(ast);
    // Pipe contract: stdout stays pure TOON / pure JSON. The engine's
    // `> fidelity: lossy tier: <tier>` first line is a human banner, not a
    // row — strip it from stdout/JSON and re-emit on stderr (human mode
    // only). --json keeps the machine signal as the `fidelity` field below.
    const bannerMatch = /^> fidelity:.*(?:\r?\n|$)/.exec(text);
    const fidelityBanner = bannerMatch ? bannerMatch[0].trimEnd() : `> fidelity: lossy tier: ${table.tier}`;
    const pureToon = bannerMatch ? text.slice(bannerMatch[0].length) : text;
    const canonicalBannerMatch = /^> fidelity:.*(?:\r?\n|$)/.exec(canonicalText);
    const pureCanonical = canonicalBannerMatch ? canonicalText.slice(canonicalBannerMatch[0].length) : canonicalText;
    if (!out.json) printStderr(fidelityBanner);
    print(out, pureToon.trimEnd());
    print(out, `rows=${table.rows.length} sections=${ast.sections.length} nodes=${summary.nodes.length}`);
    print(out, `words=${stats.words} paragraphs=${stats.paragraphs} pages~${stats.pages} (est, 500 wpp)`);
    // Would-be normalizations: the allowlist healing a save would apply
    // (dry-run preview over a DOM clone, nothing written). Clean documents
    // print nothing here; the --json envelope always carries the array.
    let pendingNormalizations: string[] = [];
    try {
      pendingNormalizations = parser.getPendingNormalizations();
    } catch {
      pendingNormalizations = [];
    }
    for (const n of pendingNormalizations.slice(0, 10)) {
      print(out, `would-be normalize: ${String(n).slice(0, 500)}`);
    }
    if (pendingNormalizations.length > 10) {
      print(out, `would-be normalize: +${pendingNormalizations.length - 10} more (total ${pendingNormalizations.length}; save applies them)`);
    }
    print(out, `Flags: [B]=bold [I]=italic [C:hex]=color [SZ:n]=half-points [RS:n]=run style [J:v]=just [N:]=numbering [TBL]=table [IMG]=image [EQ]=equation [CITE]=citation [CMT]=comment [TRK]=tracked [LNK]=link [BKMK]=bookmark [L#]=list level (from 1) [+ADDED]/[~MODIFIED]=diff [HDR]/[ZEBRA]=table look [CH:t]/[CH:?]=chart type/? [RID:id]=chart rel-id (spanN)/hidden-cover grid; placeholders [Image]/[Equation]/[Field]/[Page Break]/[Table RxC]/[Chart]/[3D Chart].`);
    print(out, `  Units: [SZ:28]=14pt (SZ is w:sz half-points); fontSize/sz bare numbers are half-points (24->w:sz 24); points only via "Npt" strings ("12pt"->w:sz 24).`);
    // Sprint 6 Track B: --visual adds spatial telemetry (<500 tok) always,
    // plus headless PNG renders (JPEG opt-in) as file paths. Missing
    // LO/pdftoppm degrades gracefully (telemetry still prints, images skip
    // with an E_DEPENDENCY hint on stderr) — exit stays 0.
    let visual: unknown = undefined;
    if (withVisual) {
      try {
        // Alias-real render: same shared path as `visual --out` (same
        // `pages: [page]` + `format` args), so the same --page/--format pair
        // yields byte-identical images. Telemetry names the EFFECTIVE
        // (clamped) page and the renderer's own page count, matching
        // `visual` exactly.
        const rendered = await renderCliVisualPage(buf, visualPage, visualFormat);
        const picked = rendered.images.find((img) => img.page === visualPage) ?? rendered.images[0];
        const telemetry = printCliVisualTelemetry(out, ast, summary, picked?.page ?? visualPage, rendered.totalPages);
        for (const img of rendered.images) {
          print(out, `visual: page-${img.page} ${img.path} (${img.bytes}B ${img.format})`);
        }
        visual = {
          telemetry,
          promptLaw: VISUAL_PROMPT_LAW,
          images: rendered.images,
          // Guard: never emit `totalPages: 0` — the renderer contract is >=1
          // page (it throws otherwise); fall back to the telemetry total if
          // violated.
          totalPages: rendered.totalPages >= 1 ? Math.floor(rendered.totalPages) : telemetry.totalPages,
          outDir: rendered.outDir,
        };
      } catch (err: unknown) {
        const code = (err as { code?: unknown })?.code;
        const msg = err instanceof Error ? err.message : String(err);
        if (code === 'E_DEPENDENCY') {
          const clean = cleanCliMessage(rewriteMcpRemedies(msg));
          // Degrade path still describes the requested page (nothing was
          // rendered, so there is no clamped pick and no renderer page
          // count to name: the marker stays word-layout-based, which is
          // honest here because no image contradicts it).
          const telemetry = printCliVisualTelemetry(out, ast, summary, visualPage);
          printStderr(`${CLI_NAME}: error: [E_DEPENDENCY] ${clean}. Remedy: install LibreOffice (soffice in PATH) and retry; telemetry above still applies`);
          visual = { telemetry, promptLaw: VISUAL_PROMPT_LAW, images: [], skipped: 'E_DEPENDENCY' };
        } else {
          return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
        }
      }
    }
    printJson(out, {
      ok: true,
      file,
      fidelity: table.tier,
      rowCount: table.rows.length,
      sections: ast.sections.length,
      nodes: summary.nodes.length,
      words: stats.words,
      paragraphs: stats.paragraphs,
      pages: stats.pages,
      toon: pureCanonical,
      pendingNormalizations,
      ...(withVisual && visual !== undefined ? { visual } : {}),
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}

