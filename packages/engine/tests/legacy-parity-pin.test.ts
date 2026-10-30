/**
 * Legacy parity pins for typed-cast batch (Q3).
 *
 * Proves the typed guards preserve exact legacy truth tables on edge inputs.
 * Each case runs old logic (from HEAD, byte-faithful) side-by-side with the fixed logic.
 *
 * DIVERGENT pins (marked INTENTIONAL below) show old truthfully even where new
 * differs; each carries a one-line reason. All divergences are garbage-in
 * crash-avoidance behind schema gates (unreachable via CLI/MCP).
 *
 * P2 OWNER-DECIDED divergences (do NOT revert — crash-avoidance, schema-gated):
 * - chart strOr/arrOr (html.ts): old `meta.title || fallback` kept non-string
 *   garbage (e.g. title=42 stayed 42 and poisoned SVG render); new coerces to
 *   fallback. Pinned below proving new safe behavior.
 * - caption guards (html.ts renderTable): old `node.caption || ... || undefined`
 *   chain passed non-strings through (caption=42 stayed 42); new narrows to
 *   string|undefined. Pinned below.
 * - agent getStats/getDagStore optional chaining (agent.ts): old
 *   `session.getStats().steps` threw when getStats() returned undefined/null;
 *   new `getStats?.()?.steps ?? 0` yields 0. Pinned below.
 */
import { describe, it, expect } from 'vitest';

// B1: legacy `e?.message || String(e)` vs fixed typed guard
function oldOr(e: unknown): unknown {
  const v = e as { message?: unknown } | null | undefined;
  return (v as { message?: string })?.message || String(e);
}
function fixedOr(err: unknown): string {
  if (err instanceof Error) return err.message !== '' ? err.message : String(err);
  if (typeof err === 'object' && err !== null && 'message' in err) {
    const m = (err as { message?: unknown }).message;
    if (typeof m === 'string' && m !== '') return m;
  }
  return String(err);
}

// B5: legacy `??` chain vs fixed typed `??`
function oldNullish(rec: { text?: unknown; textDelta?: unknown; delta?: unknown }): unknown {
  return rec.text ?? rec.textDelta ?? rec.delta ?? '';
}
function fixedNullish(rec: { text?: unknown; textDelta?: unknown; delta?: unknown }): string {
  const raw: unknown = rec.text ?? rec.textDelta ?? rec.delta ?? '';
  return typeof raw === 'string' ? raw : String(raw);
}

// B2/B3: legacy `(text || '').trim()` vs fixed metaTextOf
function fixedMetaTextOf(metadata: unknown): string {
  const rec = metadata as Record<string, unknown> | undefined;
  const t: unknown = (rec as Record<string, unknown> | undefined)?.['text'];
  if (typeof t === 'string') return t;
  if (t) return t as unknown as string;
  return '';
}
function oldAlt(text: unknown, id: string): string | null {
  try {
    const alt = ((text as string) || '').trim() || id || 'Image';
    return alt;
  } catch {
    return null;
  }
}
function fixedAlt(text: unknown, id: string): string | null {
  try {
    const alt = fixedMetaTextOf({ text }).trim() || id || 'Image';
    return alt;
  } catch {
    return null;
  }
}

// ingest (html.ts extractHtmlNoteBodies Map/object path):
// HEAD byte-faithful: `typeof v === 'string' ? v : (v?.text || String(v ?? ''))`
// vs fixed `String(rec?.['text'] ?? v ?? '')`. Full path adds `text.trim()` on store.
function oldIngestRaw(v: unknown): unknown {
  const rec = v as Record<string, unknown> | null | undefined;
  return typeof v === 'string' ? v : (rec?.['text'] || String(v ?? ''));
}
function oldIngestFull(v: unknown): string | null {
  const text = oldIngestRaw(v) as { trim?: unknown } | string | number | null | undefined;
  if (text) return (text as string).trim();
  return null;
}
function fixedIngest(v: unknown): string {
  const rec = v as Record<string, unknown> | string;
  return typeof v === 'string' ? v : String((rec as Record<string, unknown>)?.['text'] ?? v ?? '');
}
function fixedIngestFull(v: unknown): string | null {
  const text = fixedIngest(v);
  if (text) return text.trim();
  return null;
}

// ingest array path:
// HEAD byte-faithful: `if (item && item.id != null) bodies.set(id, (item.text || '').trim())`
// vs fixed: `typeof raw === 'string' ? raw : ''` then trim (never throws).
function oldArrayFull(item: Record<string, unknown>): string | null {
  if (item && item['id'] != null) {
    return ((item['text'] || '') as string).trim();
  }
  return null;
}
function fixedArrayFull(item: Record<string, unknown>): string | null {
  if (item && item['id'] != null) {
    const raw: unknown = item['text'];
    return ((typeof raw === 'string' ? raw : '') as string).trim();
  }
  return null;
}

// table cells (html.ts renderTable extractedRows):
// HEAD byte-faithful: `typeof c === 'string' ? c : c?.text ?? ''`
// vs fixed: object + typeof-text-string check, else ''.
function oldCell(c: unknown): unknown {
  return typeof c === 'string' ? c : (c as { text?: unknown } | null | undefined)?.text ?? '';
}
function fixedCell(c: unknown): string {
  if (typeof c === 'string') return c;
  if (c !== null && typeof c === 'object') {
    const t: unknown = (c as Record<string, unknown>)['text'];
    return typeof t === 'string' ? t : '';
  }
  return '';
}

// unknown-path renderTable fallbacks:
// HEAD: `node.metadata?.headers || node.headers || []`,
//   `node.metadata?.rows || (Array.isArray(...) ? ... : [])`,
//   `node.caption || node.metadata?.caption || node.metadata?.captionText`,
//   `node.captionPosition || ... || 'above'`
// vs fixed: Array.isArray / typeof-string narrowing.
type TableNode = {
  headers?: unknown;
  rows?: unknown;
  caption?: unknown;
  captionPosition?: unknown;
  metadata?: Record<string, unknown>;
};
function oldHeadersOf(node: TableNode): unknown {
  return node.metadata?.['headers'] || node.headers || [];
}
function fixedHeadersOf(node: TableNode): string[] {
  const mh: unknown = node.metadata?.['headers'];
  const nh: unknown = node.headers;
  if (Array.isArray(mh)) return (mh as unknown[]).map(String);
  if (Array.isArray(nh)) return (nh as unknown[]).map(String);
  return [];
}
function oldRowsOf(node: TableNode): unknown {
  return (
    node.metadata?.['rows'] ||
    (Array.isArray(node.rows) &&
    ((node.rows as unknown[]).length === 0 || Array.isArray((node.rows as unknown[])[0])
      ? node.rows
      : []))
  );
}
function fixedRowsOf(node: TableNode): string[][] {
  const coerce = (v: unknown): string[][] => {
    if (!Array.isArray(v)) return [];
    const arr = v as unknown[];
    if (arr.length === 0) return [];
    if (Array.isArray(arr[0])) return (arr as unknown[][]).map((r) => (r as unknown[]).map(String));
    return [];
  };
  const mr: unknown = node.metadata?.['rows'];
  if (Array.isArray(mr)) return coerce(mr);
  if (Array.isArray(node.rows)) return coerce(node.rows);
  return [];
}
function oldCaptionOf(node: TableNode): unknown {
  return node.caption || node.metadata?.['caption'] || node.metadata?.['captionText'];
}
function fixedCaptionOf(node: TableNode): string | undefined {
  const raw: unknown = node.caption || node.metadata?.['caption'] || node.metadata?.['captionText'];
  return typeof raw === 'string' ? raw : undefined;
}

// P2 chart strOr/arrOr (html.ts): old `meta.title || fallback` vs new strOr/arrOr.
function oldChartTitle(meta: Record<string, unknown>): unknown {
  return meta['title'] || 'Data Analysis Chart';
}
function fixedChartTitle(meta: Record<string, unknown>): string {
  const v: unknown = meta['title'];
  return typeof v === 'string' ? v : 'Data Analysis Chart';
}
function oldChartCategories(meta: Record<string, unknown>): unknown {
  return meta['categories'] || ['Q1', 'Q2', 'Q3', 'Q4'];
}
function fixedChartCategories(meta: Record<string, unknown>): string[] {
  const v: unknown = meta['categories'];
  return Array.isArray(v) ? (v as string[]) : ['Q1', 'Q2', 'Q3', 'Q4'];
}

// P2 agent optional chaining: old `session.getStats().steps` vs new `getStats?.()?.steps ?? 0`.
function oldStepsUsed(session: {
  stepCount?: number;
  getStats?: () => { steps: number } | null | undefined;
}): number {
  return session?.stepCount ?? (typeof session?.getStats === 'function' ? (session.getStats() as { steps: number }).steps : 0);
}
function fixedStepsUsed(session: {
  stepCount?: number;
  getStats?: () => { steps: number } | null | undefined;
}): number {
  return (
    session?.stepCount ??
    (typeof session?.getStats === 'function' ? (session.getStats?.()?.steps ?? 0) : 0)
  );
}

describe('legacy parity pins', () => {
  it('B1: empty-string message falls through to String(e) like ||', () => {
    expect(fixedOr({ message: '' })).toBe(oldOr({ message: '' }));
    expect(fixedOr(new Error(''))).toBe(oldOr(new Error('')));
    expect(fixedOr({ message: 'oops' })).toBe('oops');
    expect(fixedOr(new Error('oops'))).toBe('oops');
    expect(fixedOr(null)).toBe('null');
    expect(fixedOr(undefined)).toBe('undefined');
  });
  it('B1: falsy non-string message (0) falls through identically', () => {
    expect(oldOr({ message: 0 })).toBe('[object Object]');
    expect(fixedOr({ message: 0 })).toBe('[object Object]');
  });
  it('B1: numeric message diverges (old leaks number, new stringifies)', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(oldOr({ message: 42 })).toBe(42);
    expect(fixedOr({ message: 42 })).toBe('[object Object]');
  });
  it('B5: ?? preserves empty-string winner unlike ||', () => {
    const edge = { text: '', textDelta: 'fallback', delta: 'd2' };
    expect(fixedNullish(edge)).toBe(String(oldNullish(edge)));
    expect(fixedNullish(edge)).toBe('');
    const edge2 = { text: undefined, textDelta: '', delta: 'fallback' };
    expect(fixedNullish(edge2)).toBe('');
    expect(fixedNullish({ text: 'hi' })).toBe('hi');
  });
  it('B2/B3: numeric/object metadata.text flows like HEAD (throw-inside-try -> null)', () => {
    for (const v of ['hello', '', 0, false, null, undefined]) {
      expect(fixedAlt(v, 'n1')).toBe(oldAlt(v, 'n1'));
    }
    // Malformed truthy non-strings throw inside try -> null, exactly as HEAD
    expect(fixedAlt(42, 'n1')).toBeNull();
    expect(oldAlt(42, 'n1')).toBeNull();
    expect(fixedAlt({}, 'n1')).toBeNull();
    expect(oldAlt({}, 'n1')).toBeNull();
  });
  it('B2 bare-divider String() path preserves HEAD for numeric/object', () => {
    const oldBare = (t: unknown): string => String((t as string) || '').trim();
    const fixedBare = (t: unknown): string => String(fixedMetaTextOf({ text: t }) || '').trim();
    for (const v of ['hello', '', 42, {}, 0]) {
      expect(fixedBare(v)).toBe(oldBare(v));
    }
    expect(fixedBare(42)).toBe('42');
    expect(fixedBare({})).toBe('[object Object]');
  });
  it('ingest: string and string-text converge', () => {
    expect(oldIngestFull('hi')).toBe('hi');
    expect(fixedIngestFull('hi')).toBe('hi');
    expect(oldIngestFull({ text: 'hi' })).toBe('hi');
    expect(fixedIngestFull({ text: 'hi' })).toBe('hi');
  });
  it('ingest {text:42}: old raw leaks number then THROWS on trim, new yields "42"', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(oldIngestRaw({ text: 42 })).toBe(42);
    expect(() => oldIngestFull({ text: 42 })).toThrow(TypeError);
    expect(fixedIngestFull({ text: 42 })).toBe('42');
  });
  it('ingest falsy non-string text: old writes "[object Object]" junk, new drops/coerces', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(oldIngestFull({ text: '' })).toBe('[object Object]');
    expect(fixedIngestFull({ text: '' })).toBeNull();
    expect(oldIngestFull({ text: 0 })).toBe('[object Object]');
    expect(fixedIngestFull({ text: 0 })).toBe('0');
  });
  it('ingest array path converges on string/missing/falsy text', () => {
    expect(oldArrayFull({ id: 1, text: 'hi' })).toBe('hi');
    expect(fixedArrayFull({ id: 1, text: 'hi' })).toBe('hi');
    expect(oldArrayFull({ id: 1, text: 0 })).toBe('');
    expect(fixedArrayFull({ id: 1, text: 0 })).toBe('');
    expect(oldArrayFull({ id: 1 })).toBe('');
    expect(fixedArrayFull({ id: 1 })).toBe('');
  });
  it('ingest array {id:1,text:42}: old THROWS on trim, new yields ""', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(() => oldArrayFull({ id: 1, text: 42 })).toThrow(TypeError);
    expect(fixedArrayFull({ id: 1, text: 42 })).toBe('');
  });
  it('table-cell converges on string/bare/object-empty shapes', () => {
    expect(fixedCell('hi')).toBe(oldCell('hi'));
    expect(fixedCell({ text: 'hi' })).toBe('hi');
    expect(fixedCell(42)).toBe('');
    expect(fixedCell(42)).toBe(oldCell(42));
    expect(fixedCell(null)).toBe('');
    expect(fixedCell({ text: '' })).toBe('');
  });
  it('table-cell {text:42}: old leaks number 42 into string[][], new drops to ""', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(oldCell({ text: 42 })).toBe(42);
    expect(fixedCell({ text: 42 })).toBe('');
  });
  it('unknown-path headers/rows/caption converge on well-formed shapes', () => {
    const node: TableNode = {
      headers: ['a'],
      rows: [['b']],
      caption: 'cap',
      metadata: { headers: ['a'], rows: [['b']], caption: 'cap' },
    };
    expect(fixedHeadersOf(node)).toEqual(oldHeadersOf(node));
    expect(fixedRowsOf(node)).toEqual(oldRowsOf(node));
    expect(fixedCaptionOf(node)).toBe(oldCaptionOf(node));
  });
  it('unknown-path non-array headers/rows: old leaks garbage, new coerces to []', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(oldHeadersOf({ metadata: { headers: 42 } })).toBe(42);
    expect(fixedHeadersOf({ metadata: { headers: 42 } })).toEqual([]);
    expect(oldRowsOf({ metadata: { rows: 'x' } })).toBe('x');
    expect(fixedRowsOf({ metadata: { rows: 'x' } })).toEqual([]);
    expect(oldRowsOf({ rows: [{ foo: 1 }] })).toEqual([]);
    expect(fixedRowsOf({ rows: [{ foo: 1 }] })).toEqual([]);
  });
  it('unknown-path non-string caption: old leaks it, new narrows to undefined', () => {
    // INTENTIONAL divergence: schema-gated, unreachable via CLI/MCP — divergent inputs rejected at schema gates.
    expect(oldCaptionOf({ metadata: { caption: 42 } })).toBe(42);
    expect(fixedCaptionOf({ metadata: { caption: 42 } })).toBeUndefined();
  });
  it('P2 chart strOr/arrOr: strings/arrays converge, garbage falls back instead of poisoning render', () => {
    // INTENTIONAL divergence (owner-decided, crash-avoidance): schema-gated, unreachable via CLI/MCP.
    expect(fixedChartTitle({ title: 'T' })).toBe(oldChartTitle({ title: 'T' }));
    expect(oldChartTitle({ title: 42 })).toBe(42);
    expect(fixedChartTitle({ title: 42 })).toBe('Data Analysis Chart');
    expect(oldChartCategories({ categories: ['A'] })).toEqual(['A']);
    expect(fixedChartCategories({ categories: ['A'] })).toEqual(['A']);
    expect(oldChartCategories({ categories: 'x' })).toBe('x');
    expect(fixedChartCategories({ categories: 'x' })).toEqual(['Q1', 'Q2', 'Q3', 'Q4']);
  });
  it('P2 agent getStats: healthy sessions converge, undefined store yields 0 instead of throwing', () => {
    // INTENTIONAL divergence (owner-decided, crash-avoidance): schema-gated, unreachable via CLI/MCP.
    const healthy = { getStats: () => ({ steps: 3 }) };
    expect(fixedStepsUsed(healthy)).toBe(oldStepsUsed(healthy));
    expect(fixedStepsUsed({ stepCount: 5, getStats: () => ({ steps: 3 }) })).toBe(5);
    const broken = { getStats: () => undefined };
    expect(() => oldStepsUsed(broken)).toThrow(TypeError);
    expect(fixedStepsUsed(broken)).toBe(0);
  });
});
