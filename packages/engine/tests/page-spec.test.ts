import { describe, expect, it } from 'vitest';
import { MIN_TEXT_HEIGHT, MIN_TEXT_WIDTH, PAPERS, PAPER_NAMES, checkPage, defaultMarginsFor, detectPaper, findPaper, parseLength, resolveMargins, resolvePageSize } from '../src/page';

describe('lengths', () => {
  it('reads every unit as whole twips', () => {
    expect(parseLength('1in', 'x')).toBe(1440);
    expect(parseLength('25.4mm', 'x')).toBe(1440);
    expect(parseLength('2.54cm', 'x')).toBe(1440);
    expect(parseLength('72pt', 'x')).toBe(1440);
    expect(parseLength('1440tw', 'x')).toBe(1440);
    expect(parseLength('1440', 'x')).toBe(1440);
    expect(parseLength(1440, 'x')).toBe(1440);
    expect(parseLength(' 0.5 IN ', 'x')).toBe(720);
    expect(parseLength('.5in', 'x')).toBe(720);
    expect(parseLength('210mm', 'x')).toBe(11906);
  });
  it('rounds to the nearest twip, so a file never holds a fraction', () => {
    expect(parseLength(720.5, 'x')).toBe(721);
    expect(parseLength('10.3pt', 'x')).toBe(206);
    expect(Number.isInteger(parseLength('12.345678mm', 'x'))).toBe(true);
  });
  it('refuses what is not a length, with a remedy', () => {
    for (const bad of ['abc', '', '12 parsecs', '1..5in', 'in', NaN, Infinity, null, undefined, {}, [1]]) {
      expect(() => parseLength(bad, 'width'), String(bad)).toThrow(/E_VALIDATION|width/);
    }
    expect(() => parseLength('12 parsecs', 'width')).toThrow(/unknown unit 'parsecs'.*mm, cm, in, pt, tw/s);
  });
});

describe('papers', () => {
  it('has the named papers, with sizes that follow from their millimetres or inches', () => {
    expect(PAPER_NAMES).toEqual(['Letter', 'Legal', 'Tabloid', 'Executive', 'Statement', 'Folio', 'A3', 'A4', 'A5', 'A6', 'B4', 'B5', 'B6']);
    const mm = (n: number): number => Math.round((n / 25.4) * 1440);
    const expectMm: Record<string, [number, number]> = { A5: [148, 210], A6: [105, 148], B4: [250, 353], B5: [176, 250], B6: [125, 176] };
    for (const [name, [w, h]] of Object.entries(expectMm)) {
      const p = findPaper(name)!;
      expect([p.width, p.height], name).toEqual([mm(w), mm(h)]);
    }
    expect(findPaper('Letter')).toMatchObject({ width: 12240, height: 15840 });
    expect(findPaper('Legal')).toMatchObject({ width: 12240, height: 20160 });
    expect(findPaper('A4')).toMatchObject({ width: 11906, height: 16838 });
    expect(findPaper('A3')).toMatchObject({ width: 16838, height: 23811 });
    expect(findPaper('Tabloid')).toMatchObject({ width: 15840, height: 24480 });
  });
  it('is found without regard to case and detected in either orientation within one twip', () => {
    expect(findPaper(' a4 ')?.name).toBe('A4');
    expect(detectPaper(11906, 16838)?.name).toBe('A4');
    expect(detectPaper(16838, 11906)?.name).toBe('A4');
    expect(detectPaper(11905, 16839)?.name).toBe('A4');
    expect(detectPaper(11900, 16838)).toBeUndefined();
  });
  it('every named paper has defaults that a page can be laid out with', () => {
    for (const p of PAPERS) {
      const m = p.margins;
      expect(() => checkPage({ width: p.width, height: p.height, ...m }, p.name), p.name).not.toThrow();
      expect(checkPage({ width: p.width, height: p.height, ...m }, p.name), p.name).toEqual([]);
    }
  });
  it('the A4 defaults are what the A4 templates have', () => {
    expect(findPaper('A4')!.margins).toEqual({ top: 1440, bottom: 1440, left: 1417, right: 1417, header: 709, footer: 709 });
    expect(findPaper('Letter')!.margins).toEqual({ top: 1440, bottom: 1440, left: 1440, right: 1440, header: 720, footer: 720 });
  });
  it('a custom size gets 1 in from 8 in of short side up, and scales down to a floor of 0.5 in', () => {
    expect(defaultMarginsFor(14400, 28800).left).toBe(1440);
    expect(defaultMarginsFor(11520, 17280).left).toBe(1440);
    expect(defaultMarginsFor(8640, 12960).left).toBe(1080);
    expect(defaultMarginsFor(5760, 8640).left).toBe(720);
    expect(defaultMarginsFor(2880, 4320).left).toBe(720);
  });
});

describe('page sizes', () => {
  it('reads a name, an object and a shorthand', () => {
    expect(resolvePageSize('a5')).toMatchObject({ width: 8391, height: 11906, paper: { name: 'A5' } });
    expect(resolvePageSize({ width: '148mm', height: '210mm' })).toMatchObject({ width: 8391, height: 11906, paper: { name: 'A5' } });
    expect(resolvePageSize('148x210mm')).toMatchObject({ width: 8391, height: 11906 });
    expect(resolvePageSize('6x9in')).toEqual({ width: 8640, height: 12960, paper: undefined, named: false });
    expect(resolvePageSize('210mm x 297mm').paper?.name).toBe('A4');
    expect(resolvePageSize({ width: 9000, height: 13000 })).toEqual({ width: 9000, height: 13000, paper: undefined, named: false });
  });
  it('refuses an unknown name, a half size and extra fields, and lists what is valid', () => {
    expect(() => resolvePageSize('A9000')).toThrow(/not a known paper.*Letter, Legal/s);
    expect(() => resolvePageSize({ width: '10in' })).toThrow(/width and a height/);
    expect(() => resolvePageSize({ width: '10in', height: '5in', depth: 1 })).toThrow(/unknown field depth/);
    expect(() => resolvePageSize(42)).toThrow(/paper name or a size/);
  });
});

describe('margins', () => {
  it('reads one length for all four, or any of the seven fields', () => {
    expect(resolveMargins('25mm')).toEqual({ top: 1417, bottom: 1417, left: 1417, right: 1417 });
    expect(resolveMargins({ top: '1in', gutter: '5mm', header: 360 })).toEqual({ top: 1440, gutter: 283, header: 360 });
    expect(resolveMargins(undefined)).toEqual({});
  });
  it('refuses an unknown field', () => {
    expect(() => resolveMargins({ topp: 1 })).toThrow(/unknown field: topp/);
    expect(() => resolveMargins([1])).toThrow(/length or an object/);
  });
});

describe('what a page may be', () => {
  const ok = { width: 12240, height: 15840, top: 1440, bottom: 1440, left: 1440, right: 1440 };
  it('allows from 1 in to 22 in on a side and refuses the rest', () => {
    expect(() => checkPage({ ...ok, width: 1440, height: 1440, top: 200, bottom: 200, left: 180, right: 180 })).not.toThrow();
    expect(() => checkPage({ ...ok, width: 31680, height: 31680 })).not.toThrow();
    expect(() => checkPage({ ...ok, width: 31681 })).toThrow(/1 in to 22 in/);
    expect(() => checkPage({ ...ok, height: 144000 })).toThrow(/1 in to 22 in/);
    expect(() => checkPage({ ...ok, height: 1439 })).toThrow(/1 in to 22 in/);
    expect(checkPage({ ...ok, width: 31680 })).toEqual([]);
  });
  it('an existing page above 22 in is not refused when the request sets no size', () => {
    expect(() => checkPage({ ...ok, width: 43200 }, 'page setup', { size: false, margins: [] })).not.toThrow();
    expect(() => checkPage({ ...ok, width: 43200 }, 'page setup', { size: true, margins: [] })).toThrow(/1 in to 22 in/);
  });
  it('refuses a margin that is negative or not whole, and margins that leave too little text', () => {
    expect(() => checkPage({ ...ok, top: -1 })).toThrow(/top must be zero or more/);
    expect(() => checkPage({ ...ok, left: 720.5 })).toThrow(/left must be zero or more/);
    expect(() => checkPage({ ...ok, left: 20000, right: 20000 })).toThrow(/leave 0 in of text width/);
    expect(() => checkPage({ ...ok, top: 8000, bottom: 8000 })).toThrow(/text height/);
    expect(() => checkPage({ ...ok, left: 5000, right: 5000, gutter: 2000 })).toThrow(/and gutter/);
    expect(() => checkPage({ ...ok, top: 7000, bottom: 7000, gutter: 2000, gutterAtTop: true })).toThrow(/and gutter/);
  });
  it('the minimum text width is 0.75 in and the minimum text height is 0.5 in, to the twip', () => {
    const side = (12240 - MIN_TEXT_WIDTH) / 2;
    expect(() => checkPage({ ...ok, left: side, right: side })).not.toThrow();
    expect(() => checkPage({ ...ok, left: side + 1, right: side })).toThrow(/text width.*at least 0.75 in/);
    const tb = (15840 - MIN_TEXT_HEIGHT) / 2;
    expect(() => checkPage({ ...ok, top: tb, bottom: tb })).not.toThrow();
    expect(() => checkPage({ ...ok, top: tb + 1, bottom: tb })).toThrow(/text height.*at least 0.5 in/);
  });
  it('a negative margin that the request does not set is kept, and one that it sets is refused', () => {
    const neg = { ...ok, top: -720 };
    expect(() => checkPage(neg, 'page setup', { margins: ['left'] })).not.toThrow();
    expect(() => checkPage(neg, 'page setup', { margins: ['top'] })).toThrow(/top must be zero or more/);
  });
  it('a page that is already under the minimum is refused only when the request makes it smaller', () => {
    const narrow = { ...ok, left: 5600, right: 5600 };
    const before = { ...narrow };
    expect(() => checkPage(narrow, 'page setup', { margins: [], before })).not.toThrow();
    expect(() => checkPage({ ...narrow, left: 5700 }, 'page setup', { margins: ['left'], before })).toThrow(/text width/);
    expect(() => checkPage({ ...narrow, left: 5500 }, 'page setup', { margins: ['left'], before })).not.toThrow();
    const short = { ...ok, top: 7700, bottom: 7700 };
    expect(() => checkPage(short, 'page setup', { margins: [], before: short })).not.toThrow();
    expect(() => checkPage({ ...short, top: 7800 }, 'page setup', { margins: ['top'], before: short })).toThrow(/text height/);
  });
  it('warns about a low margin and a header that is larger than its margin, for what the request sets', () => {
    expect(checkPage({ ...ok, left: 200 })).toEqual([expect.stringMatching(/left margin is under 0.25 in/)]);
    expect(checkPage({ ...ok, left: 200 }, 'page setup', { margins: ['top'] })).toEqual([]);
    expect(checkPage({ ...ok, header: 2000 })).toEqual([expect.stringMatching(/header distance is larger than the top margin/)]);
    expect(checkPage({ ...ok, footer: 2000 })).toEqual([expect.stringMatching(/footer distance/)]);
  });
  it('names the section in a message', () => {
    expect(() => checkPage({ ...ok, top: -1 }, 'section 2')).toThrow(/section 2:/);
  });
});

describe('default margins keep the minimum text area at every size', () => {
  const sizes: Array<[number, number]> = [];
  for (let w = 1440; w <= 2880; w += 144) for (const h of [1440, 1584, 2160, 15840]) sizes.push([w, h], [h, w]);
  it('every size from 1 in to 2 in on a side gets margins that leave the minimum', () => {
    for (const [w, h] of sizes) {
      const m = defaultMarginsFor(w, h);
      const g = { width: w, height: h, ...m };
      expect(w - m.left - m.right, `${w}x${h} width`).toBeGreaterThanOrEqual(MIN_TEXT_WIDTH);
      expect(h - m.top - m.bottom, `${w}x${h} height`).toBeGreaterThanOrEqual(MIN_TEXT_HEIGHT);
      expect(() => checkPage(g), `${w}x${h}`).not.toThrow();
      expect(m.header).toBeLessThanOrEqual(m.top);
    }
  });
  it('a 1 in wide page gets 0.125 in side margins and 0.75 in of text', () => {
    const m = defaultMarginsFor(1440, 15840);
    expect(m.left).toBe(180);
    expect(1440 - m.left - m.right).toBe(1080);
  });
  it('where there is room the margin is the usual one, and it is cut only as far as the minimum needs', () => {
    expect(defaultMarginsFor(12240 + 1, 15840).left).toBeGreaterThan(0);
    const m = defaultMarginsFor(2160, 15840);
    expect(2160 - m.left - m.right).toBeGreaterThanOrEqual(MIN_TEXT_WIDTH);
    expect(m.left).toBe(Math.min(540, Math.floor((2160 - MIN_TEXT_WIDTH) / 2)));
    expect(defaultMarginsFor(12240, 15840).left).toBe(1440);
  });
});
