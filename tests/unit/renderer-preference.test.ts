import { describe, it, expect } from 'vitest';
import {
  AUTO_PREFERS_WORD,
  CircuitBreaker,
  effectiveRasterizerPreference,
  effectiveRendererPreference,
  parseRasterizerPreference,
  parseRendererPreference,
  planRasterizers,
  planRenderers,
} from '../../packages/engine/src/visual/renderer-preference';
import { runPlan, failureClassOf } from '../../packages/engine/src/visual/render-plan';

const win = 'win32' as const;
const mac = 'darwin' as const;
const linux = 'linux' as const;

describe('renderer preference parsing', () => {
  it('accepts the three values case-insensitively and trims', () => {
    expect(parseRendererPreference('AUTO')).toBe('auto');
    expect(parseRendererPreference(' Word ')).toBe('word');
    expect(parseRendererPreference('libreoffice')).toBe('libreoffice');
  });
  it('treats undefined, null and an empty string as "not set"', () => {
    expect(parseRendererPreference(undefined)).toBeUndefined();
    expect(parseRendererPreference(null)).toBeUndefined();
    expect(parseRendererPreference('   ')).toBeUndefined();
  });
  it('rejects anything else with E_VALIDATION naming the source and the valid values', () => {
    expect(() => parseRendererPreference('lo', 'flag')).toThrowError(/flag renderer "lo" is not one of auto, word, libreoffice/);
    try {
      parseRendererPreference('x', 'ASTRA_RENDERER');
    } catch (e) {
      expect((e as { code: string }).code).toBe('E_VALIDATION');
      expect((e as Error).message).toContain('ASTRA_RENDERER');
      expect((e as { remedy: string }).remedy).toMatch(/auto/);
    }
    expect(() => parseRendererPreference(5 as unknown as string)).toThrow(/must be one of/);
  });
  it('truncates a very long invalid value in the message', () => {
    try {
      parseRendererPreference('x'.repeat(500));
    } catch (e) {
      expect((e as Error).message.length).toBeLessThan(200);
    }
  });
  it('resolves explicit > environment > auto, and a bad environment value fails loudly', () => {
    expect(effectiveRendererPreference(undefined, {})).toBe('auto');
    expect(effectiveRendererPreference(undefined, { ASTRA_RENDERER: 'word' })).toBe('word');
    expect(effectiveRendererPreference('libreoffice', { ASTRA_RENDERER: 'word' })).toBe('libreoffice');
    expect(effectiveRendererPreference(undefined, { ASTRA_RENDERER: '' })).toBe('auto');
    expect(() => effectiveRendererPreference(undefined, { ASTRA_RENDERER: 'nope' })).toThrow(/ASTRA_RENDERER/);
    // an explicit call value wins even over an invalid environment value
    expect(effectiveRendererPreference('word', { ASTRA_RENDERER: 'nope' })).toBe('word');
  });
  it('parses and resolves the pdf-to-image choice the same way', () => {
    expect(parseRasterizerPreference('Windows')).toBe('windows');
    expect(parseRasterizerPreference('poppler')).toBe('poppler');
    expect(() => parseRasterizerPreference('gs', 'flag')).toThrow(/flag pdf-to-image "gs"/);
    expect(effectiveRasterizerPreference(undefined, { ASTRA_PDF_TO_IMAGE: 'poppler' })).toBe('poppler');
    expect(effectiveRasterizerPreference('windows', { ASTRA_PDF_TO_IMAGE: 'poppler' })).toBe('windows');
    expect(() => effectiveRasterizerPreference(undefined, { ASTRA_PDF_TO_IMAGE: 'x' })).toThrow(/ASTRA_PDF_TO_IMAGE/);
  });
});

describe('planRenderers', () => {
  it('0.5.2: Word is opt-in, so auto never starts with Word', () => {
    expect(AUTO_PREFERS_WORD).toBe(false);
    expect(planRenderers('auto', { platform: win, word: true, libreoffice: true }).engines).toEqual(['libreoffice']);
    expect(planRenderers('auto', { platform: win, word: true, libreoffice: false }).engines).toEqual([]);
  });
  it('explicit word on Windows with Word installed uses only Word and never falls back', () => {
    const p = planRenderers('word', { platform: win, word: true, libreoffice: true });
    expect(p).toEqual({ requested: 'word', engines: ['word'], canFallBack: false });
  });
  it('explicit word without Word, or off Windows, is an E_DEPENDENCY with a one-sentence remedy', () => {
    for (const platform of [mac, linux]) {
      try {
        planRenderers('word', { platform, word: false, libreoffice: true });
        throw new Error('should have thrown');
      } catch (e) {
        expect((e as { code: string }).code).toBe('E_DEPENDENCY');
        expect((e as Error).message).toMatch(/only available on Windows/);
        expect((e as { remedy: string }).remedy).toMatch(/WSL/);
      }
    }
    expect(() => planRenderers('word', { platform: win, word: false, libreoffice: true })).toThrow(/Word is not installed/);
  });
  it('libreoffice never uses Word; missing LibreOffice gives an empty plan for the caller to report', () => {
    expect(planRenderers('libreoffice', { platform: win, word: true, libreoffice: true }).engines).toEqual(['libreoffice']);
    expect(planRenderers('libreoffice', { platform: win, word: true, libreoffice: false }).engines).toEqual([]);
    expect(planRenderers('libreoffice', { platform: linux, word: false, libreoffice: false }).engines).toEqual([]);
  });
  it('auto on macOS and Linux is always LibreOffice', () => {
    for (const platform of [mac, linux]) {
      expect(planRenderers('auto', { platform, word: false, libreoffice: true }).engines).toEqual(['libreoffice']);
      expect(planRenderers('auto', { platform, word: true, libreoffice: true }).engines).toEqual(['libreoffice']);
      expect(planRenderers('auto', { platform, word: false, libreoffice: false }).engines).toEqual([]);
    }
  });
});

describe('planRasterizers', () => {
  it('auto on Windows prefers the built-in renderer and falls back to Poppler', () => {
    const p = planRasterizers('auto', { platform: win, windowsRenderer: true, poppler: true });
    expect(p.engines).toEqual(['windows', 'poppler']);
    expect(p.canFallBack).toBe(true);
  });
  it('auto on Windows without Poppler uses only the built-in renderer; without both is empty', () => {
    expect(planRasterizers('auto', { platform: win, windowsRenderer: true, poppler: false }).engines).toEqual(['windows']);
    expect(planRasterizers('auto', { platform: win, windowsRenderer: false, poppler: true }).engines).toEqual(['poppler']);
    expect(planRasterizers('auto', { platform: win, windowsRenderer: false, poppler: false }).engines).toEqual([]);
  });
  it('auto elsewhere is Poppler', () => {
    for (const platform of [mac, linux]) {
      expect(planRasterizers('auto', { platform, windowsRenderer: true, poppler: true }).engines).toEqual(['poppler']);
      expect(planRasterizers('auto', { platform, windowsRenderer: false, poppler: false }).engines).toEqual([]);
    }
  });
  it('poppler is the opt-out: never the Windows renderer', () => {
    expect(planRasterizers('poppler', { platform: win, windowsRenderer: true, poppler: true }).engines).toEqual(['poppler']);
    expect(planRasterizers('poppler', { platform: win, windowsRenderer: true, poppler: false }).engines).toEqual([]);
  });
  it('explicit windows off Windows or without the renderer is E_DEPENDENCY', () => {
    expect(() => planRasterizers('windows', { platform: linux, windowsRenderer: false, poppler: true })).toThrow(/only available on Windows/);
    expect(() => planRasterizers('windows', { platform: win, windowsRenderer: false, poppler: true })).toThrow(/not available/);
    expect(planRasterizers('windows', { platform: win, windowsRenderer: true, poppler: true })).toEqual({ requested: 'windows', engines: ['windows'], canFallBack: false });
  });
  it('an open breaker removes the Windows renderer from auto only', () => {
    expect(planRasterizers('auto', { platform: win, windowsRenderer: true, poppler: true, windowsUsable: false }).engines).toEqual(['poppler']);
    expect(planRasterizers('windows', { platform: win, windowsRenderer: true, poppler: true, windowsUsable: false }).engines).toEqual(['windows']);
  });
});

describe('CircuitBreaker', () => {
  it('opens for the cool-down and closes after it, on an injected clock', () => {
    let now = 1000;
    const b = new CircuitBreaker(300_000, () => now);
    expect(b.isOpen()).toBe(false);
    b.trip();
    expect(b.isOpen()).toBe(true);
    now += 299_999;
    expect(b.isOpen()).toBe(true);
    now += 2;
    expect(b.isOpen()).toBe(false);
    b.trip();
    b.reset();
    expect(b.isOpen()).toBe(false);
  });
});

describe('runPlan', () => {
  const infra = (m: string) => Object.assign(new Error(`E_RENDER: ${m}`), { failureClass: 'infrastructure' as const });
  const doc = (m: string) => Object.assign(new Error(`E_VALIDATION: ${m}`), { failureClass: 'document' as const });

  it('returns the first success and reports no fallback', async () => {
    const r = await runPlan(['a', 'b'] as const, { canFallBack: true, auto: true }, async (e) => e);
    expect(r).toEqual({ value: 'a', used: 'a' });
  });
  it('falls through on failure with a reason, and trips an infrastructure breaker only in auto', async () => {
    const now = 0;
    const breaker = new CircuitBreaker(1000, () => now);
    const r = await runPlan(
      ['a', 'b'] as const,
      { canFallBack: true, auto: true },
      async (e) => {
        if (e === 'a') throw infra('word hung');
        return 'ok';
      },
      { a: breaker }
    );
    expect(r.used).toBe('b');
    expect(r.fallbackFrom).toBe('a');
    expect(r.fallbackReason).toBe('word hung');
    expect(breaker.isOpen()).toBe(true);
  });
  it('a document failure falls back but does not trip the breaker', async () => {
    const breaker = new CircuitBreaker(1000);
    const r = await runPlan(['a', 'b'] as const, { canFallBack: true, auto: true }, async (e) => { if (e === 'a') throw doc('password'); return 1; }, { a: breaker });
    expect(r.fallbackFrom).toBe('a');
    expect(breaker.isOpen()).toBe(false);
  });
  it('an explicit single engine never falls back and does not trip the breaker', async () => {
    const breaker = new CircuitBreaker(1000);
    await expect(runPlan(['a'] as const, { canFallBack: false, auto: false }, async () => { throw infra('x'); }, { a: breaker })).rejects.toThrow(/x/);
    expect(breaker.isOpen()).toBe(false);
  });
  it('auto with a single engine still trips the breaker on an infrastructure failure', async () => {
    const breaker = new CircuitBreaker(1000);
    await expect(runPlan(['a'] as const, { canFallBack: false, auto: true }, async () => { throw infra('x'); }, { a: breaker })).rejects.toThrow();
    expect(breaker.isOpen()).toBe(true);
  });
  it('when every engine fails the last error is thrown with the earlier cause named', async () => {
    const err = await runPlan(['a', 'b'] as const, { canFallBack: true, auto: true }, async (e) => { throw e === 'a' ? infra('first cause') : doc('second cause'); }).catch((e) => e);
    expect(err.message).toContain('second cause');
    expect(err.message).toContain('a failed first: first cause');
    expect(failureClassOf(err)).toBe('document');
    expect(err.earlierFailures).toEqual(['a: first cause']);
  });
  it('a success resets that engine\'s breaker', async () => {
    const breaker = new CircuitBreaker(1000);
    breaker.trip();
    await runPlan(['a'] as const, { canFallBack: false, auto: false }, async () => 1, { a: breaker });
    expect(breaker.isOpen()).toBe(false);
  });
});
