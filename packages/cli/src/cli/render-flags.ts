import {
  describeReferenceReport,
  parseRasterizerPreference,
  parseRendererPreference,
  type ReferenceReport,
  type RasterizerReport,
  type RendererReport,
} from '@attentionform/astra-engine';
import { printStderr } from './output';

/**
 * The flags that choose how a page or PDF is rendered. Shared by `visual`, `inspect --visual` and `export`,
 * so the three commands read them identically:
 *   --renderer auto|word|libreoffice      (env ASTRA_RENDERER)
 *   --pdf-to-image auto|windows|poppler   (env ASTRA_PDF_TO_IMAGE; page images only)
 *   --allow-network-references | --no-allow-network-references   (env ASTRA_ALLOW_NETWORK_REFERENCES)
 * Each value flag takes `--flag value` or `--flag=value`. The last of the two network flags wins.
 */
export interface RenderFlags {
  renderer?: string;
  pdfToImage?: string;
  allowNetworkReferences?: boolean;
}

export type RenderFlagResult =
  | { matched: false }
  | { matched: true; next: number }
  | { matched: true; error: string };

/**
 * Try to read one render flag at `args[i]`. Returns `{ matched: false }` for any other argument, so the
 * caller's own parsing continues. On a match, `next` is the index of the last argument consumed.
 */
export function readRenderFlag(
  command: string,
  args: string[],
  i: number,
  flags: RenderFlags,
  opts: { pdfToImage: boolean }
): RenderFlagResult {
  const a = args[i];
  if (a === '--allow-network-references') {
    flags.allowNetworkReferences = true;
    return { matched: true, next: i };
  }
  if (a === '--no-allow-network-references') {
    flags.allowNetworkReferences = false;
    return { matched: true, next: i };
  }
  for (const name of opts.pdfToImage ? (['renderer', 'pdf-to-image'] as const) : (['renderer'] as const)) {
    const flag = `--${name}`;
    let value: string | undefined;
    let next = i;
    if (a === flag) {
      value = args[i + 1];
      next = i + 1;
      if (value === undefined) {
        return { matched: true, error: `${command} ${flag} requires a value (${name === 'renderer' ? 'auto|word|libreoffice' : 'auto|windows|poppler'})` };
      }
    } else if (a.startsWith(`${flag}=`)) {
      value = a.slice(flag.length + 1);
    } else {
      continue;
    }
    try {
      if (name === 'renderer') {
        const parsed = parseRendererPreference(value, 'flag');
        if (parsed === undefined) return { matched: true, error: `${command} ${flag} requires a value (auto|word|libreoffice)` };
        flags.renderer = parsed;
      } else {
        const parsed = parseRasterizerPreference(value, 'flag');
        if (parsed === undefined) return { matched: true, error: `${command} ${flag} requires a value (auto|windows|poppler)` };
        flags.pdfToImage = parsed;
      }
    } catch (err: unknown) {
      return { matched: true, error: `${command}: ${err instanceof Error ? err.message.replace(/^E_VALIDATION:\s*/, '') : String(err)}` };
    }
    return { matched: true, next };
  }
  return { matched: false };
}

/** The engine options for the parsed flags (a flag that was not given is left out, so the environment applies). */
export function renderOptionsFromFlags(flags: RenderFlags): {
  renderer?: 'auto' | 'word' | 'libreoffice';
  pdfToImage?: 'auto' | 'windows' | 'poppler';
  allowNetworkReferences?: boolean;
} {
  return {
    ...(flags.renderer !== undefined ? { renderer: flags.renderer as 'auto' | 'word' | 'libreoffice' } : {}),
    ...(flags.pdfToImage !== undefined ? { pdfToImage: flags.pdfToImage as 'auto' | 'windows' | 'poppler' } : {}),
    ...(flags.allowNetworkReferences !== undefined ? { allowNetworkReferences: flags.allowNetworkReferences } : {}),
  };
}

export interface RenderFacts {
  renderer?: RendererReport;
  rasterizer?: RasterizerReport;
  references?: ReferenceReport;
}

/**
 * `note:` lines on stderr (never stdout, so `--json` output stays pure) for what the caller should know:
 * a fallback to another engine, and network references that were not loaded or may be loaded.
 */
export function printRenderNotes(facts: RenderFacts): string[] {
  const notes: string[] = [];
  const r = facts.renderer;
  if (r?.fallbackFrom) notes.push(`renderer: ${r.fallbackFrom} could not render this document (${r.fallbackReason ?? 'no reason given'}); used ${r.used} instead`);
  const p = facts.rasterizer;
  if (p?.fallbackFrom) notes.push(`pdf-to-image: ${p.fallbackFrom} could not draw the pages (${p.fallbackReason ?? 'no reason given'}); used ${p.used} instead`);
  const ref = facts.references ? describeReferenceReport(facts.references) : null;
  if (ref) notes.push(ref);
  for (const n of notes) printStderr(`note: ${n}`);
  return notes;
}

/** Plain name of the engine that rendered, for messages that used to say "LibreOffice". */
export function rendererLabel(r: RendererReport | undefined): string {
  return r?.used === 'word' ? 'Word' : 'LibreOffice';
}
