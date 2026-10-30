// MS-Q2b PART1_QUALITY Q2.2: verbatim split from lib/mcp-tool-adapter.ts (lines 35-74).
// Zero logic edits. Display-only normalizer reads.
// ── Normalizer display surface ────────────────────────────────────────────
// Display-only: save/batch envelopes carry `normalized[]` (counts +
// transforms); the rewrites live in the engine core (allowlist normalizer
// inside `DocxParser.rezip()`, surfaced via its public `lastNormalizations`
// log). No repair logic here — only defensive reads of the core's own log
// ([] when the seam is absent or mid-flight, never a throw).
export type NormalizedRewrite = string;

/** Display-only read of an engine result's `normalized` report ([] when absent). */
export function readEngineNormalized(res: unknown): NormalizedRewrite[] {
  const list = (res as { normalized?: unknown } | null | undefined)?.normalized;
  if (Array.isArray(list)) return list.filter((v): v is string => typeof v === 'string');
  return [];
}

/** Display-only read of the core's own normalizer log ([] when absent). */
export function readParserNormalizations(parser: unknown): NormalizedRewrite[] {
  const list = (parser as { lastNormalizations?: unknown } | null | undefined)?.lastNormalizations;
  if (Array.isArray(list)) return list.filter((v): v is string => typeof v === 'string');
  return [];
}

/**
 * Merge core logs across the rezips one save performs (gate rezip + save
 * rezip), dropping exact duplicates. The gate's internal rezip already heals
 * and logs, so reading only the final rezip would report a dishonest zero.
 */
export function mergeNormalized(...lists: string[][]): string[] {
  const seen = new Set<string>();
  const out: string[] = [];
  for (const list of lists) {
    for (const entry of list) {
      if (!seen.has(entry)) {
        seen.add(entry);
        out.push(entry);
      }
    }
  }
  return out;
}
