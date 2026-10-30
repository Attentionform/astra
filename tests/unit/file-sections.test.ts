/**
 * Opening a file gives the sections the file defines: one per section break
 * paragraph plus the one the body ends, each holding the blocks the file puts
 * in it, with the page setup its own w:sectPr gives. Checked against an
 * independent reading of the XML, over hand-built shapes and over files made by
 * python-docx and LibreOffice.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, createBlankDocxBuffer, type DocumentAST } from '@attentionform/astra-engine';
import { SECTION_SHAPES, blockSignature, countFileSections, fileBlocksBySection } from '../helpers/section-docs';

const FIXTURES = path.resolve(__dirname, '../fixtures/sections');

interface Source {
  name: string;
  build: () => Buffer;
}

const sources: Source[] = [
  ...Object.entries(SECTION_SHAPES).map(([key, shape]) => ({ name: `shape ${key}`, build: shape.build })),
  { name: 'python-docx', build: () => fs.readFileSync(path.join(FIXTURES, 'python-docx-three-sections.docx')) },
  { name: 'libreoffice', build: () => fs.readFileSync(path.join(FIXTURES, 'libreoffice-two-page-styles.docx')) },
  { name: 'blank document', build: () => createBlankDocxBuffer('blank') },
];

const open = async (buf: Buffer): Promise<{ parser: DocxParser; ast: DocumentAST }> => {
  const parser = new DocxParser(buf);
  return { parser, ast: await parser.getAST() };
};

const blockTexts = (ast: DocumentAST): string[][] => ast.sections.map((s) => s.nodes.map((n) => blockSignature(n.xml as Element)));

describe('the model has the sections the file has', () => {
  for (const source of sources) {
    it(`${source.name}: same sections, same blocks in each, same order`, async () => {
      const buf = source.build();
      const { ast } = await open(buf);
      expect(ast.sections).toHaveLength(countFileSections(buf));
      expect(blockTexts(ast)).toEqual(fileBlocksBySection(buf));
    });

    it(`${source.name}: section ids are sec_root, sec_2, sec_3 ... and no block is in two sections`, async () => {
      const { ast } = await open(source.build());
      expect(ast.sections.map((s) => s.id)).toEqual(ast.sections.map((_, i) => (i === 0 ? 'sec_root' : `sec_${i + 1}`)));
      const ids = ast.sections.flatMap((s) => s.nodes.map((n) => n.id));
      expect(new Set(ids).size).toBe(ids.length);
    });

    it(`${source.name}: every section but the last ends with its break paragraph`, async () => {
      const { ast } = await open(source.build());
      ast.sections.slice(0, -1).forEach((section) => {
        const last = section.nodes[section.nodes.length - 1].xml as Element;
        expect(last.tagName).toBe('w:p');
        const pPr = Array.from(last.childNodes as unknown as Element[]).find((c) => c.tagName === 'w:pPr')!;
        expect(Array.from(pPr.childNodes as unknown as Element[]).some((c) => c.tagName === 'w:sectPr')).toBe(true);
      });
    });
  }

  it('a file with no section break is one section called sec_root', async () => {
    const { ast } = await open(SECTION_SHAPES.single.build());
    expect(ast.sections.map((s) => s.id)).toEqual(['sec_root']);
  });

  it('a break paragraph at the very end leaves a last section with no blocks', async () => {
    const { ast } = await open(SECTION_SHAPES['trailing-empty-section'].build());
    expect(ast.sections.map((s) => s.nodes.length)).toEqual([2, 0]);
  });

  it('a break paragraph at the very start is a section of its own', async () => {
    const { ast } = await open(SECTION_SHAPES['break-paragraph-first'].build());
    expect(ast.sections.map((s) => s.nodes.length)).toEqual([1, 2]);
  });

  it('a break inside a content control still ends its section', async () => {
    const { ast } = await open(SECTION_SHAPES['break-in-content-control'].build());
    expect(blockTexts(ast)).toEqual([['w:p::One A', 'w:p::One B'], ['w:p::Two A']]);
  });

  it('a tracked section change is not a section, and a sectPr in a table cell is not one either', async () => {
    expect((await open(SECTION_SHAPES['tracked-section-change'].build())).ast.sections).toHaveLength(2);
    const cell = await open(SECTION_SHAPES['sectpr-in-table-cell'].build());
    expect(cell.ast.sections).toHaveLength(1);
    expect(cell.parser.getSectPrElements()).toHaveLength(1);
  });
});

describe('each section has the page setup of its own w:sectPr', () => {
  const setups = async (buf: Buffer) =>
    (await open(buf)).ast.sections.map((s) => `${s.pageSetup?.paperSize} ${s.pageSetup?.orientation}`);

  it('portrait, landscape, portrait', async () => {
    expect(await setups(SECTION_SHAPES['three-sections'].build())).toEqual(['Letter portrait', 'Letter landscape', 'Letter portrait']);
  });
  it('Letter then A4', async () => {
    expect(await setups(SECTION_SHAPES['two-sections'].build())).toEqual(['Letter portrait', 'A4 portrait']);
  });
  it('python-docx: the landscape section is the second', async () => {
    expect(await setups(fs.readFileSync(path.join(FIXTURES, 'python-docx-three-sections.docx')))).toEqual([
      'Letter portrait',
      'Letter landscape',
      'Letter portrait',
    ]);
  });
  it('LibreOffice: Letter, A4 landscape, Letter', async () => {
    expect(await setups(fs.readFileSync(path.join(FIXTURES, 'libreoffice-two-page-styles.docx')))).toEqual([
      'Letter portrait',
      'A4 landscape',
      'Letter portrait',
    ]);
  });
  it('margins come from the section they belong to', async () => {
    const { ast } = await open(SECTION_SHAPES['landscape-section'].build());
    expect(ast.sections[0].pageSetup).toMatchObject({ marginTop: 1440, marginLeft: 1440 });
    expect(ast.sections[1].pageSetup).toMatchObject({ marginTop: 1080, marginLeft: 1080 });
  });
  it('a stray sectPr in a table cell does not decide the document page', async () => {
    const { ast } = await open(SECTION_SHAPES['sectpr-in-table-cell'].build());
    expect(ast.sections[0].pageSetup?.paperSize).toBe('Letter');
  });
});

describe('the section accessors list sections, not every sectPr in the XML', () => {
  it('one w:sectPr per section, in order, the body one last', async () => {
    for (const source of sources) {
      const buf = source.build();
      const { parser } = await open(buf);
      const sectPrs = parser.getSectPrElements();
      expect(sectPrs).toHaveLength(countFileSections(buf));
      expect(sectPrs[sectPrs.length - 1]).toBe(parser.getBodySectPr());
    }
  });

  it('the properties kept inside a tracked section change are not listed', async () => {
    const { parser } = await open(SECTION_SHAPES['tracked-section-change'].build());
    const sectPrs = parser.getSectPrElements();
    expect(sectPrs).toHaveLength(2);
    expect(sectPrs.every((s) => (s.parentNode as Element).tagName !== 'w:sectPrChange')).toBe(true);
  });
});

describe('header and footer types come from the section that uses the part', () => {
  it('a part only the second section references as "first" is typed first', async () => {
    const { ast } = await open(SECTION_SHAPES['first-page-header'].build());
    expect(ast.headers['header1.xml'].type).toBe('default');
    expect(ast.headers['header2.xml'].type).toBe('first');
  });

  it('the default header is the document header text, not the first-page one', async () => {
    const { ast } = await open(SECTION_SHAPES['first-page-header'].build());
    expect(ast.metadata.headerText).toBe('Default header');
  });

  it('each section can have its own default header', async () => {
    const { ast } = await open(SECTION_SHAPES['headers-per-section'].build());
    expect(ast.headers['header1.xml'].type).toBe('default');
    expect(ast.headers['header2.xml'].type).toBe('default');
  });

  it('the parity flags still describe the first section only', async () => {
    const { ast, parser } = await open(SECTION_SHAPES['first-page-header'].build());
    // Section 2 has w:titlePg; section 1 does not.
    expect(ast.sectPrFlags).toEqual({ titlePg: false, evenAndOddHeaders: false });
    expect(parser.getStructuralSummary().sectPrFlags).toEqual({ titlePg: false, evenAndOddHeaders: false });
  });
});

describe('saving and reopening keeps the sections', () => {
  for (const source of sources) {
    it(`${source.name}: a save with no edits reads back the same sections`, async () => {
      const first = await open(source.build());
      const second = await open(first.parser.rezip());
      expect(second.ast.sections.map((s) => s.id)).toEqual(first.ast.sections.map((s) => s.id));
      expect(blockTexts(second.ast)).toEqual(blockTexts(first.ast));
      expect(second.ast.sections.map((s) => s.pageSetup)).toEqual(first.ast.sections.map((s) => s.pageSetup));
      expect(second.ast.sections.flatMap((s) => s.nodes.map((n) => n.id))).toEqual(first.ast.sections.flatMap((s) => s.nodes.map((n) => n.id)));
    });
  }
});
