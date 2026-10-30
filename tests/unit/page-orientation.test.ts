/**
 * pageSetup turns the page: Word and LibreOffice lay a section out from
 * `w:pgSz w:w` and `w:h`, not from the `w:orient` flag, so changing the
 * orientation has to change the two numbers as well. Margins are never
 * rotated; the values given are the values written.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { execFileSync } from 'child_process';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, convertDocxToPdf, type DocumentAST } from '@attentionform/astra-engine';
import { withLoSlot } from '../helpers/lo-slot';

interface Fixture {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}

async function fixtureFor(parser: DocxParser): Promise<Fixture> {
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}

const blank = (): Promise<Fixture> => fixtureFor(DocxParser.createNewDocument('blank'));

/** A one-paragraph document whose `w:sectPr` holds only what the caller writes. */
function withSectPr(sectPrInner: string): Promise<Fixture> {
  const zip = new AdmZip();
  zip.addFile(
    'word/document.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Hello</w:t></w:r></w:p><w:sectPr>${sectPrInner}</w:sectPr></w:body></w:document>`,
      'utf8',
    ),
  );
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'),
  );
  return fixtureFor(new DocxParser(zip.toBuffer()));
}

/** The blank document split into two sections by a next-page break. */
async function twoSections(): Promise<Fixture> {
  const f = await blank();
  const anchor = f.ast.sections[0].nodes[0].id;
  const out = await f.engine.executeBatch({
    operations: [
      { action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'nextPage' },
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Second section text' },
    ] as never,
  });
  expect(out.success).toBe(true);
  expect(f.ast.sections.length).toBe(2);
  return f;
}

interface Page {
  w: string | null;
  h: string | null;
  orient: string | null;
  code: string | null;
}

const pageOf = (f: Fixture, section = 1): Page => {
  const pgSz = f.parser.getSectPrElements()[section - 1].getElementsByTagName('w:pgSz')[0];
  const attr = (name: string): string | null => (pgSz.hasAttribute(name) ? pgSz.getAttribute(name) : null);
  return { w: attr('w:w'), h: attr('w:h'), orient: attr('w:orient'), code: attr('w:code') };
};

const marginsOf = (f: Fixture, section = 1): Record<string, string | null> => {
  const pgMar = f.parser.getSectPrElements()[section - 1].getElementsByTagName('w:pgMar')[0];
  return Object.fromEntries(['top', 'bottom', 'left', 'right'].map((k) => [k, pgMar?.getAttribute(`w:${k}`) ?? null]));
};

const setup = async (f: Fixture, params: Record<string, unknown>) => {
  const out = await f.engine.executeBatch({ operations: [{ action: 'pageSetup', ...params }] as never });
  expect(out.success, JSON.stringify(out.errors)).toBe(true);
};

const LETTER = { w: '12240', h: '15840' };
const LETTER_TURNED = { w: '15840', h: '12240' };
const A4 = { w: '11906', h: '16838' };
const A4_TURNED = { w: '16838', h: '11906' };

describe('orientation turns the page (the matrix from the findings report)', () => {
  it('baseline is Letter portrait', async () => {
    const f = await blank();
    expect(pageOf(f)).toMatchObject(LETTER);
  });

  it('orientation: landscape alone turns the page', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toMatchObject({ ...LETTER_TURNED, orient: 'landscape' });
  });

  it('orientation: landscape with paperSize: A4 is A4 turned', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape', paperSize: 'A4' });
    expect(pageOf(f)).toMatchObject({ ...A4_TURNED, orient: 'landscape' });
  });

  it('A4 first, then orientation: landscape turns A4', async () => {
    const f = await blank();
    await setup(f, { paperSize: 'A4' });
    expect(pageOf(f)).toMatchObject(A4);
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toMatchObject({ ...A4_TURNED, orient: 'landscape' });
  });

  it('landscape A4, then orientation: portrait turns it back', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape', paperSize: 'A4' });
    await setup(f, { orientation: 'portrait' });
    expect(pageOf(f)).toMatchObject({ ...A4, orient: 'portrait' });
  });

  it('landscape Letter, then paperSize: A4 stays landscape', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape' });
    await setup(f, { paperSize: 'A4' });
    expect(pageOf(f)).toMatchObject({ ...A4_TURNED, orient: 'landscape' });
  });

  it('margins given with orientation are written as given, never rotated', async () => {
    const f = await blank();
    const before = marginsOf(f);
    await setup(f, { orientation: 'landscape', marginTop: 1000, marginLeft: 2000 });
    expect(pageOf(f)).toMatchObject({ ...LETTER_TURNED, orient: 'landscape' });
    expect(marginsOf(f)).toEqual({ ...before, top: '1000', left: '2000' });
  });

  it('orientation alone leaves the margins untouched', async () => {
    const f = await blank();
    const before = marginsOf(f);
    await setup(f, { orientation: 'landscape' });
    expect(marginsOf(f)).toEqual(before);
  });
});

describe('every paper size, both orientations', () => {
  const SIZES: Array<[string, number, number]> = [
    ['A3', 16838, 23811],
    ['A4', 11906, 16838],
    ['Letter', 12240, 15840],
    ['Legal', 12240, 20160],
  ];

  for (const [name, short, long] of SIZES) {
    it(`${name}: portrait is ${short} x ${long}, landscape is ${long} x ${short}`, async () => {
      const f = await blank();
      await setup(f, { paperSize: name, orientation: 'portrait' });
      expect(pageOf(f)).toMatchObject({ w: String(short), h: String(long), orient: 'portrait' });
      await setup(f, { orientation: 'landscape' });
      expect(pageOf(f)).toMatchObject({ w: String(long), h: String(short), orient: 'landscape' });
      await setup(f, { orientation: 'portrait' });
      expect(pageOf(f)).toMatchObject({ w: String(short), h: String(long), orient: 'portrait' });
    });
  }

  it('asking for the orientation the page already has changes nothing', async () => {
    const f = await blank();
    await setup(f, { paperSize: 'Legal', orientation: 'landscape' });
    const turned = pageOf(f);
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toEqual(turned);
  });
});

describe('sections and odd files', () => {
  it('sectionIndex 2 turns only the second section', async () => {
    const f = await twoSections();
    await setup(f, { sectionIndex: 2, orientation: 'landscape' });
    expect(pageOf(f, 1)).toMatchObject({ ...LETTER });
    expect(pageOf(f, 1).orient).not.toBe('landscape');
    expect(pageOf(f, 2)).toMatchObject({ ...LETTER_TURNED, orient: 'landscape' });
  });

  it('without sectionIndex every section turns', async () => {
    const f = await twoSections();
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f, 1)).toMatchObject(LETTER_TURNED);
    expect(pageOf(f, 2)).toMatchObject(LETTER_TURNED);
  });

  it('a section with no w:pgSz is laid out as Letter, then turned', async () => {
    const f = await withSectPr('<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    expect(f.parser.getSectPrElements()[0].getElementsByTagName('w:pgSz')).toHaveLength(0);
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toMatchObject({ ...LETTER_TURNED, orient: 'landscape' });
    expect(marginsOf(f)).toEqual({ top: '1440', bottom: '1440', left: '1440', right: '1440' });
  });

  it('a pgSz with a flag but no numbers gets the numbers', async () => {
    const f = await withSectPr('<w:pgSz w:orient="portrait"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toMatchObject({ ...LETTER_TURNED, orient: 'landscape' });
  });

  it('a page that is already turned by its numbers (no flag) is kept when only the paper changes', async () => {
    const f = await withSectPr('<w:pgSz w:w="15840" w:h="12240"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    await setup(f, { paperSize: 'A4' });
    expect(pageOf(f)).toMatchObject(A4_TURNED);
  });

  it('a square page is left alone', async () => {
    const f = await withSectPr('<w:pgSz w:w="12000" w:h="12000"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toMatchObject({ w: '12000', h: '12000', orient: 'landscape' });
  });

  it('a new paper size drops the old printer paper code', async () => {
    const f = await withSectPr('<w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    await setup(f, { paperSize: 'Letter' });
    expect(pageOf(f)).toMatchObject({ ...LETTER, code: null });
    const reopened = await fixtureFor(new DocxParser(f.parser.rezip()));
    expect(reopened.ast.sections[0].pageSetup?.paperSize).toBe('Letter');
  });

  it('orientation alone keeps the printer paper code of the same paper', async () => {
    const f = await withSectPr('<w:pgSz w:w="11906" w:h="16838" w:code="9"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>');
    await setup(f, { orientation: 'landscape' });
    expect(pageOf(f)).toMatchObject({ ...A4_TURNED, code: '9' });
  });
});

describe('the model and the saved file agree', () => {
  it('the model reports the orientation, and a save and reopen reads the same page', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape', paperSize: 'A4' });
    expect(f.ast.sections[0].pageSetup).toMatchObject({ orientation: 'landscape', paperSize: 'A4' });

    const reopened = await fixtureFor(new DocxParser(f.parser.rezip()));
    expect(reopened.ast.sections[0].pageSetup).toMatchObject({ orientation: 'landscape', paperSize: 'A4' });
    expect(pageOf(reopened)).toMatchObject({ ...A4_TURNED, orient: 'landscape' });
  });

  it('orientation alone: the saved page is wider than tall, which is what the reader derives landscape from', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape' });
    const reopened = await fixtureFor(new DocxParser(f.parser.rezip()));
    const { w, h } = pageOf(reopened);
    expect(Number(w)).toBeGreaterThan(Number(h));
    expect(reopened.ast.sections[0].pageSetup?.orientation).toBe('landscape');
  });

  it('a failed batch leaves the page as it was', async () => {
    const f = await blank();
    const before = pageOf(f);
    const out = await f.engine.executeBatch({
      operations: [
        { action: 'pageSetup', orientation: 'landscape' },
        { action: 'pageSetup', paperSize: 'Tabloid' },
      ] as never,
    });
    expect(out.success).toBe(false);
    expect(pageOf(f)).toEqual(before);
  });
});

// The page size in the PDF is what a reader sees, so the live check asks the
// renderer, not the XML. Opt-in like the other LibreOffice probes.
const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1';

describe.skipIf(!RUN_LIVE)('LibreOffice lays the page out as turned', () => {
  const pdfPageSize = async (f: Fixture): Promise<{ width: number; height: number }> => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-orient-'));
    try {
      const pdf = path.join(dir, 'page.pdf');
      await withLoSlot(() => convertDocxToPdf(f.parser.rezip(), pdf));
      const match = /Page size:\s+([\d.]+) x ([\d.]+) pts/.exec(execFileSync('pdfinfo', [pdf], { encoding: 'utf8' }));
      expect(match, 'pdfinfo printed a page size').not.toBeNull();
      return { width: Number(match![1]), height: Number(match![2]) };
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  };

  it('orientation: landscape alone gives a wide page; portrait gives it back', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape' });
    const wide = await pdfPageSize(f);
    expect(wide.width).toBeGreaterThan(wide.height);
    expect(Math.round(wide.width)).toBe(792);
    expect(Math.round(wide.height)).toBe(612);

    await setup(f, { orientation: 'portrait' });
    const tall = await pdfPageSize(f);
    expect(Math.round(tall.width)).toBe(612);
    expect(Math.round(tall.height)).toBe(792);
  }, 300_000);

  it('A4 landscape is 842 x 595 points', async () => {
    const f = await blank();
    await setup(f, { orientation: 'landscape', paperSize: 'A4' });
    const size = await pdfPageSize(f);
    expect(Math.round(size.width)).toBe(842);
    expect(Math.round(size.height)).toBe(595);
  }, 300_000);
});
