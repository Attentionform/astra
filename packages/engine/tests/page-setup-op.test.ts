import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer } from '../src/internal';
import { PAPERS } from '../src/page';

const sectPrs = (parser: DocxParser): string[] => (new AdmZip(parser.rezip()).readAsText('word/document.xml').match(/<w:sectPr[\s\S]*?<\/w:sectPr>/g) ?? []);
const attr = (xml: string, tag: string, name: string): string | undefined => new RegExp(`<${tag}\\b[^>]*\\b${name}="([^"]*)"`).exec(xml)?.[1];
const settings = (parser: DocxParser): string => new AdmZip(parser.rezip()).readAsText('word/settings.xml');

async function run(template: string, ops: object[], opts?: { pageSize?: 'letter' | 'a4' }) {
  const parser = DocxParser.createNewDocument(template, opts);
  const engine = new OperationEngine(await parser.getAST(), parser);
  const res = await engine.executeBatch({ operations: ops as never });
  return { parser, res, engine };
}
const valid = async (parser: DocxParser): Promise<void> => {
  const v = await validateDocxBuffer(parser.rezip());
  expect(v.ok, JSON.stringify(v.errors)).toBe(true);
};

describe('every named paper, in both orientations', () => {
  for (const paper of PAPERS) {
    for (const orientation of ['portrait', 'landscape'] as const) {
      it(`${paper.name} ${orientation}`, async () => {
        const { parser, res } = await run('blank', [{ action: 'pageSetup', paperSize: paper.name, orientation }]);
        expect(res.success, JSON.stringify(res.errors)).toBe(true);
        const sect = sectPrs(parser)[0];
        const [w, h] = orientation === 'portrait' ? [paper.width, paper.height] : [paper.height, paper.width];
        expect(attr(sect, 'w:pgSz', 'w:w')).toBe(String(w));
        expect(attr(sect, 'w:pgSz', 'w:h')).toBe(String(h));
        expect(attr(sect, 'w:pgSz', 'w:orient')).toBe(orientation);
        expect(attr(sect, 'w:pgSz', 'w:code')).toBe(paper.code === undefined ? undefined : String(paper.code));
        const ast = await parser.getAST();
        expect(ast.sections[0].pageSetup?.paperSize).toBe(paper.name);
        await valid(parser);
      });
    }
  }
});

describe('custom sizes', () => {
  const cases: Array<[string, unknown, number, number]> = [
    ['an object in millimetres', { width: '150mm', height: '200mm' }, 8504, 11339],
    ['an object in inches', { width: '6in', height: '9in' }, 8640, 12960],
    ['an object in centimetres', { width: '15cm', height: '22cm' }, 8504, 12472],
    ['an object in points', { width: '432pt', height: '648pt' }, 8640, 12960],
    ['an object in twips', { width: 9000, height: 13000 }, 9000, 13000],
    ['a shorthand in inches', '6x9in', 8640, 12960],
    ['a shorthand in millimetres', '160x240mm', 9071, 13606],
    ['the largest page, 22 in', '22x22in', 31680, 31680],
  ];
  for (const [label, paperSize, w, h] of cases) {
    it(`${label} is written as given and has no printer code`, async () => {
      const { parser, res } = await run('blank', [{ action: 'pageSetup', paperSize, margins: '25mm' }]);
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      const sect = sectPrs(parser)[0];
      expect(attr(sect, 'w:pgSz', 'w:w')).toBe(String(w));
      expect(attr(sect, 'w:pgSz', 'w:h')).toBe(String(h));
      expect(attr(sect, 'w:pgSz', 'w:code')).toBeUndefined();
      const ast = await parser.getAST();
      expect(ast.sections[0].pageSetup).toMatchObject({ pageWidthTwips: w, pageHeightTwips: h });
      await valid(parser);
    });
  }
  it('a custom size that is a named paper within a twip reads back as that paper, and any other as Custom', async () => {
    const a = await run('blank', [{ action: 'pageSetup', paperSize: { width: '210mm', height: '297mm' } }]);
    expect((await a.parser.getAST()).sections[0].pageSetup?.paperSize).toBe('A4');
    const exact = await run('blank', [{ action: 'pageSetup', paperSize: { width: '148mm', height: '210mm' } }]);
    expect(attr(sectPrs(exact.parser)[0], 'w:pgSz', 'w:code')).toBe('11');
    const b = await run('blank', [{ action: 'pageSetup', paperSize: { width: '150mm', height: '200mm' } }]);
    expect((await b.parser.getAST()).sections[0].pageSetup?.paperSize).toBe('Custom');
  });
  it('a width and a height are the page as given, and orientation turns it', async () => {
    const wide = await run('blank', [{ action: 'pageSetup', paperSize: { width: '300mm', height: '100mm' } }]);
    expect(attr(sectPrs(wide.parser)[0], 'w:pgSz', 'w:orient')).toBe('landscape');
    const turned = await run('blank', [{ action: 'pageSetup', paperSize: { width: '300mm', height: '100mm' }, orientation: 'portrait' }]);
    expect(attr(sectPrs(turned.parser)[0], 'w:pgSz', 'w:w')).toBe('5669');
    expect(attr(sectPrs(turned.parser)[0], 'w:pgSz', 'w:h')).toBe('17008');
  });
  it('a paper name keeps the orientation that the section is in', async () => {
    const { parser } = await run('blank', [{ action: 'pageSetup', orientation: 'landscape' }, { action: 'pageSetup', paperSize: 'A5' }]);
    expect(attr(sectPrs(parser)[0], 'w:pgSz', 'w:w')).toBe('11906');
    expect(attr(sectPrs(parser)[0], 'w:pgSz', 'w:h')).toBe('8391');
  });
});

describe('margins', () => {
  it('reads every unit, per field, and writes whole twips with every pgMar attribute present', async () => {
    const { parser, res } = await run('blank', [{ action: 'pageSetup', margins: { top: '20mm', bottom: '0.75in', left: '2cm', right: '54pt', header: 400, footer: '8mm', gutter: '5mm' } }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const sect = sectPrs(parser)[0];
    expect(['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'].map((n) => attr(sect, 'w:pgMar', `w:${n}`))).toEqual(['1134', '1080', '1134', '1080', '400', '454', '283']);
    await valid(parser);
  });
  it('one length is all four page margins and leaves header and footer distances alone', async () => {
    const { parser } = await run('blank', [{ action: 'pageSetup', margins: '1in' }]);
    const sect = sectPrs(parser)[0];
    expect(['top', 'bottom', 'left', 'right', 'header', 'footer'].map((n) => attr(sect, 'w:pgMar', `w:${n}`))).toEqual(['1440', '1440', '1440', '1440', '720', '720']);
  });
  it('the fields can be given at the top level or in params, and params win', async () => {
    const { parser } = await run('blank', [{ action: 'pageSetup', margins: '1in', params: { margins: { left: '2in' } } }]);
    expect(attr(sectPrs(parser)[0], 'w:pgMar', 'w:left')).toBe('2880');
  });
  it('a fraction never reaches the file: 720.5 twips rounds, and the file validates', async () => {
    const { parser, res } = await run('blank', [{ action: 'pageSetup', margins: { top: 720.5 } }]);
    expect(res.success).toBe(true);
    expect(attr(sectPrs(parser)[0], 'w:pgMar', 'w:top')).toBe('721');
    await valid(parser);
  });
  it('is read back with header, footer and gutter', async () => {
    const { parser } = await run('blank', [{ action: 'pageSetup', margins: { header: 300, footer: 310, gutter: 200 } }]);
    expect((await parser.getAST()).sections[0].pageSetup).toMatchObject({ headerDistance: 300, footerDistance: 310, gutter: 200 });
  });
  it('a file whose pgMar lacks attributes ends up with all of them', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const zip = new AdmZip(parser.rezip());
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:pgMar[^>]*\/>/, '<w:pgMar w:top="1440"/>')));
    const reopened = new DocxParser(zip.toBuffer());
    const engine = new OperationEngine(await reopened.getAST(), reopened);
    const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', margins: { left: '1in' } }] as never });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const sect = sectPrs(reopened)[0];
    for (const n of ['top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter']) expect(attr(sect, 'w:pgMar', `w:${n}`), n).toMatch(/^\d+$/);
    await valid(reopened);
  });
});

describe('gutter and mirrored margins', () => {
  it('mirrorMargins and a left gutter go into settings.xml in schema order and the file validates', async () => {
    const { parser, res } = await run('blank', [{ action: 'pageSetup', margins: { gutter: '10mm' }, mirrorMargins: true }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(settings(parser)).toMatch(/<w:mirrorMargins\/>/);
    expect(attr(sectPrs(parser)[0], 'w:pgMar', 'w:gutter')).toBe('567');
    await valid(parser);
  });
  it('a top gutter goes into settings.xml and the file validates', async () => {
    const { parser, res } = await run('blank', [{ action: 'pageSetup', margins: { gutter: '10mm' }, gutterPosition: 'top' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(settings(parser)).toMatch(/<w:gutterAtTop\/>/);
    await valid(parser);
  });
  it('a top gutter with mirrored margins is refused, in one request and across two, and nothing changes', async () => {
    const both = await run('blank', [{ action: 'pageSetup', mirrorMargins: true, gutterPosition: 'top' }]);
    expect(both.res.success).toBe(false);
    expect(JSON.stringify(both.res.errors)).toMatch(/top gutter cannot be used with mirrorMargins/);
    expect(settings(both.parser)).not.toMatch(/mirrorMargins|gutterAtTop/);
    const mirrored = await run('blank', [{ action: 'pageSetup', mirrorMargins: true }]);
    const engine = new OperationEngine(await mirrored.parser.getAST(), mirrored.parser);
    const top = await engine.executeBatch({ operations: [{ action: 'pageSetup', gutterPosition: 'top' }] as never });
    expect(top.success).toBe(false);
    expect(settings(mirrored.parser)).not.toMatch(/gutterAtTop/);
    const together = await engine.executeBatch({ operations: [{ action: 'pageSetup', mirrorMargins: false, gutterPosition: 'top' }] as never });
    expect(together.success, JSON.stringify(together.errors)).toBe(true);
    expect(settings(mirrored.parser)).not.toMatch(/mirrorMargins/);
    expect(settings(mirrored.parser)).toMatch(/gutterAtTop/);
  });
  it('a top gutter is refused with book fold printing or two pages per sheet in the file', async () => {
    for (const tag of ['w:bookFoldPrinting', 'w:bookFoldRevPrinting', 'w:printTwoOnOne']) {
      const parser = DocxParser.createNewDocument('blank');
      const zip = new AdmZip(parser.rezip());
      zip.updateFile('word/settings.xml', Buffer.from(zip.readAsText('word/settings.xml').replace('</w:settings>', `<${tag}/></w:settings>`)));
      const reopened = new DocxParser(zip.toBuffer());
      const engine = new OperationEngine(await reopened.getAST(), reopened);
      const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', gutterPosition: 'top' }] as never });
      expect(res.success, tag).toBe(false);
      expect(JSON.stringify(res.errors), tag).toMatch(/top gutter cannot be used with/);
    }
  });
  it('a file that already combines them can still be edited for something else', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const zip = new AdmZip(parser.rezip());
    zip.updateFile('word/settings.xml', Buffer.from(zip.readAsText('word/settings.xml').replace('</w:settings>', '<w:mirrorMargins/><w:gutterAtTop/></w:settings>')));
    const reopened = new DocxParser(zip.toBuffer());
    const engine = new OperationEngine(await reopened.getAST(), reopened);
    const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', orientation: 'landscape' }] as never });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
  });
  it('switching them off removes the elements, and the same request twice changes nothing more', async () => {
    const on = await run('blank', [{ action: 'pageSetup', mirrorMargins: true }]);
    const engine = new OperationEngine(await on.parser.getAST(), on.parser);
    const twice = await engine.executeBatch({ operations: [{ action: 'pageSetup', mirrorMargins: true }] as never });
    expect(twice.success).toBe(true);
    expect(settings(on.parser).match(/<w:mirrorMargins/g)).toHaveLength(1);
    const off = await engine.executeBatch({ operations: [{ action: 'pageSetup', mirrorMargins: false, gutterPosition: 'left' }] as never });
    expect(off.success).toBe(true);
    expect(settings(on.parser)).not.toMatch(/mirrorMargins|gutterAtTop/);
    await valid(on.parser);
  });
  it('the gutter changes the outcome: the same margins pass without it and are refused with it', async () => {
    // 5 x 5 in: 1.5 in side margins leave 2 in of width; a 1.2 in left gutter leaves 0.8 in (passes), a 1.4 in gutter leaves 0.6 in (refused, under 0.75 in).
    const base = { action: 'pageSetup', paperSize: '5x5in' };
    const noGutter = await run('blank', [{ ...base, margins: { top: '1in', bottom: '1in', left: '1.5in', right: '1.5in', gutter: 0 } }]);
    expect(noGutter.res.success, JSON.stringify(noGutter.res.errors)).toBe(true);
    const fits = await run('blank', [{ ...base, margins: { top: '1in', bottom: '1in', left: '1.5in', right: '1.5in', gutter: '1.2in' } }]);
    expect(fits.res.success, JSON.stringify(fits.res.errors)).toBe(true);
    const tooMuch = await run('blank', [{ ...base, margins: { top: '1in', bottom: '1in', left: '1.5in', right: '1.5in', gutter: '1.4in' } }]);
    expect(tooMuch.res.success).toBe(false);
    expect(JSON.stringify(tooMuch.res.errors)).toMatch(/text width.*and gutter/s);
    // The same gutter on the top edge comes off the height instead: a 5 in page with 1.2 in top and bottom margins leaves 2.6 in, and a 2.2 in gutter leaves 0.4 in.
    const top = await run('blank', [{ ...base, gutterPosition: 'top', margins: { top: '1.2in', bottom: '1.2in', left: '1.5in', right: '1.5in', gutter: '2.2in' } }]);
    expect(top.res.success).toBe(false);
    expect(JSON.stringify(top.res.errors)).toMatch(/text height.*and gutter/s);
    const leftSame = await run('blank', [{ ...base, margins: { top: '1.2in', bottom: '1.2in', left: '1.5in', right: '1.5in', gutter: '2.2in' } }]);
    expect(JSON.stringify(leftSame.res.errors ?? [])).toMatch(/text width/);
  });
  it('moving the gutter edge is checked against every section, even when one section is named', async () => {
    const parser = DocxParser.createNewDocument('report');
    const engine = new OperationEngine(await parser.getAST(), parser);
    // Section 2 gets a 4.5 in left gutter on a Letter page: width 8.5 - 2 - 4.5 = 2 in, height untouched. A top gutter then takes 4.5 in off section 2's height, which stays 6 in.
    const big = await engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 2, margins: { gutter: '9in' } }] as never });
    expect(big.success).toBe(false);
    const ok = await engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 2, margins: { gutter: '5in' } }] as never });
    expect(ok.success, JSON.stringify(ok.errors)).toBe(true);
    const before = sectPrs(parser).join('') + settings(parser);
    // Moving the gutter to the top for section 1 only: section 2 has a 5 in gutter and 9 in of height after margins, so it fits; shrink section 2 first to make it not fit.
    const squeeze = await engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 2, margins: { top: '4in', bottom: '4in' } }] as never });
    expect(squeeze.success, JSON.stringify(squeeze.errors)).toBe(true);
    const mid = sectPrs(parser).join('') + settings(parser);
    const move = await engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 1, gutterPosition: 'top' }] as never });
    expect(move.success).toBe(false);
    expect(JSON.stringify(move.errors)).toMatch(/section 2/);
    expect(sectPrs(parser).join('') + settings(parser)).toBe(mid);
    expect(before).not.toBe(mid);
  });
});

describe('what is refused, and what is only warned about', () => {
  const refuse = async (op: object, pattern: RegExp) => {
    const { parser, res } = await run('blank', [op]);
    const before = sectPrs(DocxParser.createNewDocument('blank'))[0];
    expect(res.success).toBe(false);
    expect(JSON.stringify(res.errors)).toMatch(pattern);
    expect(sectPrs(parser)[0]).toBe(before);
  };
  it('an unknown paper, a half size, a bad unit and a bad number', async () => {
    await refuse({ action: 'pageSetup', paperSize: 'A9000' }, /not a known paper/);
    await refuse({ action: 'pageSetup', paperSize: { width: '10in' } }, /width and a height/);
    await refuse({ action: 'pageSetup', margins: '12 parsecs' }, /unknown unit/);
    await refuse({ action: 'pageSetup', margins: { top: 'abc' } }, /not a length/);
  });
  it('a size outside 1 in to 22 in', async () => {
    await refuse({ action: 'pageSetup', paperSize: '22.1x10in' }, /1 in to 22 in/);
    await refuse({ action: 'pageSetup', paperSize: '100x60in' }, /1 in to 22 in/);
    await refuse({ action: 'pageSetup', paperSize: '0.9x10in' }, /1 in to 22 in/);
  });
  it('a negative margin and margins that leave no text', async () => {
    await refuse({ action: 'pageSetup', margins: { top: '-5mm' } }, /top must be zero or more/);
    await refuse({ action: 'pageSetup', margins: { left: 20000, right: 20000 } }, /text width/);
    await refuse({ action: 'pageSetup', margins: { left: 5640, right: 5640 } }, /at least 0.75 in/);
    await refuse({ action: 'pageSetup', margins: { top: 9000, bottom: 9000 } }, /text height/);
  });
  it('an unknown margin field and a bad orientation', async () => {
    await refuse({ action: 'pageSetup', margins: { topp: 1 } }, /Unrecognized key|unknown field/i);
    await refuse({ action: 'pageSetup', orientation: 'sideways' }, /orientation|Invalid/i);
  });
  it('one bad section in an all-sections change stops the whole change', async () => {
    const parser = DocxParser.createNewDocument('report');
    const before = sectPrs(parser).join('');
    const engine = new OperationEngine(await parser.getAST(), parser);
    const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', margins: { left: '3in', right: '3in' }, paperSize: 'A6' }] as never });
    expect(res.success).toBe(false);
    expect(sectPrs(parser).join('')).toBe(before);
  });
  it('a small margin and a large header distance are warned about', async () => {
    const { res } = await run('blank', [{ action: 'pageSetup', margins: { left: '3mm', header: '40mm' } }]);
    expect(res.success).toBe(true);
    expect(res.warnings.join('|')).toMatch(/left margin is under 0.25 in/);
    expect(res.warnings.join('|')).toMatch(/header distance is larger than the top margin/);
  });
});

describe('sections', () => {
  it('sectionIndex changes one section and leaves the others', async () => {
    const parser = DocxParser.createNewDocument('report');
    const before = sectPrs(parser);
    expect(before.length).toBeGreaterThan(1);
    const engine = new OperationEngine(await parser.getAST(), parser);
    const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 2, paperSize: 'A5', margins: '15mm' }] as never });
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const after = sectPrs(parser);
    expect(after[0]).toBe(before[0]);
    expect(attr(after[1], 'w:pgSz', 'w:w')).toBe('8391');
    for (let i = 2; i < after.length; i++) expect(after[i]).toBe(before[i]);
    await valid(parser);
  });
  it('without sectionIndex every section changes', async () => {
    const parser = DocxParser.createNewDocument('report');
    const engine = new OperationEngine(await parser.getAST(), parser);
    const res = await engine.executeBatch({ operations: [{ action: 'pageSetup', paperSize: 'Legal' }] as never });
    expect(res.success).toBe(true);
    for (const s of sectPrs(parser)) expect(attr(s, 'w:pgSz', 'w:h')).toBe('20160');
  });
});

describe('any request that succeeds produces a file the validator accepts', () => {
  it('300 random pages and margins in mixed units, up to 30 in so that some are refused for size', async () => {
    let seed = 20261121;
    const rnd = (): number => ((seed = (seed * 1664525 + 1013904223) >>> 0) / 4294967296);
    const units = ['mm', 'cm', 'in', 'pt', ''];
    const len = (min: number, max: number): string => { const u = units[Math.floor(rnd() * units.length)]; const base = min + rnd() * (max - min); return u === '' ? String(Math.round(base * 1440)) : `${(u === 'mm' ? base * 25.4 : u === 'cm' ? base * 2.54 : u === 'pt' ? base * 72 : base).toFixed(2)}${u}`; };
    const parser = DocxParser.createNewDocument('blank');
    const engine = new OperationEngine(await parser.getAST(), parser);
    let accepted = 0;
    for (let i = 0; i < 300; i++) {
      const op = { action: 'pageSetup', paperSize: rnd() < 0.3 ? PAPERS[Math.floor(rnd() * PAPERS.length)].name : { width: len(0.5, 30), height: len(0.5, 30) }, margins: { top: len(0, 6), bottom: len(0, 6), left: len(0, 6), right: len(0, 6), gutter: len(0, 2) } };
      const res = await engine.executeBatch({ operations: [op] as never });
      if (res.success) accepted += 1;
      else expect(JSON.stringify(res.errors)).toMatch(/E_VALIDATION|1 in to 22 in|text width|text height|must be zero or more/);
      if (i % 50 === 49) await valid(parser);
    }
    expect(accepted).toBeGreaterThan(30);
    await valid(parser);
  });
});
