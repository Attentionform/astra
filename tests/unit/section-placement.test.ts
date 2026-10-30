/**
 * New content lands in the section the file puts it in, on a file that was
 * opened with several sections. After every operation the model's sections must
 * hold the same blocks, in the same order, as the sections of the saved file,
 * and nothing may crash because the document has more than one section.
 *
 * Covers the placements that do not depend on the break paragraph itself:
 * content with no anchor goes to the end of the document (the last section),
 * and content next to a block goes in that block's section.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine, type DocumentAST } from '@attentionform/astra-engine';
import { SECTION_SHAPES, blockSignature, countFileSections, fileBlocksBySection } from '../helpers/section-docs';

const FIXTURES = path.resolve(__dirname, '../fixtures/sections');
/** A 1x1 PNG. */
const PNG =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

interface Source {
  name: string;
  build: () => Buffer;
}
const sources: Source[] = [
  { name: 'two sections', build: SECTION_SHAPES['two-sections'].build },
  { name: 'three sections', build: SECTION_SHAPES['three-sections'].build },
  { name: 'table before a break', build: SECTION_SHAPES['table-before-break'].build },
  { name: 'break in a content control', build: SECTION_SHAPES['break-in-content-control'].build },
  { name: 'python-docx', build: () => fs.readFileSync(path.join(FIXTURES, 'python-docx-three-sections.docx')) },
  { name: 'libreoffice', build: () => fs.readFileSync(path.join(FIXTURES, 'libreoffice-two-page-styles.docx')) },
];

interface Session {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}
const open = async (buf: Buffer): Promise<Session> => {
  const parser = new DocxParser(buf);
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
};

const modelBlocks = (ast: DocumentAST): string[][] => ast.sections.map((s) => s.nodes.map((n) => blockSignature(n.xml as Element)));

/** The model has the sections, and the blocks in them, that the file would have if saved now. */
const expectModelMatchesFile = (session: Session): void => {
  const saved = session.parser.rezip();
  expect(session.ast.sections).toHaveLength(countFileSections(saved));
  expect(modelBlocks(session.ast)).toEqual(fileBlocksBySection(saved));
};

const apply = async (session: Session, op: Record<string, unknown>) => {
  const out = await session.engine.executeBatch({ operations: [op] as never });
  expect(out.success, JSON.stringify(out.errors)).toBe(true);
  return out;
};

const lastSection = (ast: DocumentAST) => ast.sections[ast.sections.length - 1];
/** A node that is not a break paragraph: the first block of the second section. */
const secondSectionFirst = (ast: DocumentAST) => ast.sections[1].nodes[0];
/** The last block of the last section that has any. */
const lastBlock = (ast: DocumentAST) => [...ast.sections].reverse().find((s) => s.nodes.length > 0)!.nodes.slice(-1)[0];

/** The break paragraph that ends the first section: the last block it holds. */
const firstBreak = (ast: DocumentAST) => ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
const counts = (ast: DocumentAST): number[] => ast.sections.map((section) => section.nodes.length);

/**
 * Content put "after" a break paragraph is in the next section in the file, so
 * the first section keeps its blocks and the second gains the new one.
 */
const expectStartedSecondSection = (session: Session, before: number[]): void => {
  expectModelMatchesFile(session);
  const after = counts(session.ast);
  expect(after[0]).toBe(before[0]);
  expect(after[1]).toBeGreaterThan(before[1]);
};

const chart = (extra: Record<string, unknown> = {}) => ({
  action: 'chart',
  chartType: 'column',
  title: 'C',
  categories: ['a', 'b'],
  series: [{ name: 's', values: [1, 2] }],
  ...extra,
});

for (const source of sources) {
  describe(`placement on a reopened file: ${source.name}`, () => {
    it('insert paragraph with no anchor goes to the end of the document', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'insert', type: 'paragraph', text: 'NO ANCHOR' });
      expectModelMatchesFile(s);
      expect(lastSection(s.ast).nodes.slice(-1)[0].xml.textContent).toBe('NO ANCHOR');
    });

    it('insert paragraph after a block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'insert', type: 'paragraph', anchor: secondSectionFirst(s.ast).id, position: 'after', text: 'AFTER S2' });
      expectModelMatchesFile(s);
      expect(s.ast.sections[1].nodes.some((n) => n.xml.textContent === 'AFTER S2')).toBe(true);
    });

    it('insert paragraph before the first block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'insert', type: 'paragraph', anchor: secondSectionFirst(s.ast).id, position: 'before', text: 'BEFORE S2' });
      expectModelMatchesFile(s);
      expect(s.ast.sections[1].nodes[0].xml.textContent).toBe('BEFORE S2');
    });

    it('insert table with no anchor goes to the end of the document', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'insert', type: 'table', headers: ['A'], rows: [['1']] });
      expectModelMatchesFile(s);
      expect(lastSection(s.ast).nodes.some((n) => n.type === 'table' && n.xml.textContent?.includes('A1'))).toBe(true);
    });

    it('callout with no anchor goes to the end of the document (it used to crash)', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'callout', text: 'Note', title: 'T' });
      expectModelMatchesFile(s);
    });

    it('callout after a block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'callout', anchor: secondSectionFirst(s.ast).id, position: 'after', text: 'Note', title: 'T' });
      expectModelMatchesFile(s);
    });

    it('a table of contents with no anchor goes to the end of the document (it used to crash)', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'field', operation: 'insertTOC' });
      expectModelMatchesFile(s);
    });

    it('a table of contents after a block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'field', operation: 'insertTOC', anchor: secondSectionFirst(s.ast).id, position: 'after' });
      expectModelMatchesFile(s);
    });

    it('chart with no anchor goes to the end of the document', async () => {
      const s = await open(source.build());
      await apply(s, chart());
      expectModelMatchesFile(s);
      expect(lastSection(s.ast).nodes.slice(-1)[0].type).toBe('chart');
    });

    it('chart after a block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, chart({ anchor: secondSectionFirst(s.ast).id }));
      expectModelMatchesFile(s);
    });

    it('image after a block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'image', operation: 'insert', anchor: secondSectionFirst(s.ast).id, fileData: PNG, width: 914400, height: 914400 });
      expectModelMatchesFile(s);
    });

    it('image with no anchor is refused and leaves the sections as they were', async () => {
      const s = await open(source.build());
      const before = modelBlocks(s.ast);
      const out = await s.engine.executeBatch({ operations: [{ action: 'image', operation: 'insert', fileData: PNG, width: 914400, height: 914400 }] as never });
      expect(out.success).toBe(false);
      expect(JSON.stringify(out.errors)).toContain('requires an anchor');
      expect(modelBlocks(s.ast)).toEqual(before);
      expectModelMatchesFile(s);
    });

    describe('after the break paragraph that ends the first section', () => {
      it('a paragraph starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, { action: 'insert', type: 'paragraph', anchor: firstBreak(s.ast).id, position: 'after', text: 'AFTER BREAK' });
        expectStartedSecondSection(s, before);
        expect(s.ast.sections[1].nodes[0].xml.textContent).toBe('AFTER BREAK');
      });

      it('two paragraphs inserted one after the other keep their order', async () => {
        const s = await open(source.build());
        const anchor = firstBreak(s.ast).id;
        await apply(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'FIRST' });
        await apply(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'SECOND' });
        expectModelMatchesFile(s);
        expect(s.ast.sections[1].nodes.slice(0, 2).map((n) => n.xml.textContent)).toEqual(['FIRST', 'SECOND']);
      });

      it('a callout starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, { action: 'callout', anchor: firstBreak(s.ast).id, position: 'after', text: 'Note', title: 'T' });
        expectStartedSecondSection(s, before);
      });

      it('a table of contents starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, { action: 'field', operation: 'insertTOC', anchor: firstBreak(s.ast).id, position: 'after' });
        expectStartedSecondSection(s, before);
      });

      it('a page-number field starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, { action: 'field', operation: 'insertPageNumber', target: 'body', anchor: firstBreak(s.ast).id, position: 'after' });
        expectStartedSecondSection(s, before);
      });

      it('a chart starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, chart({ anchor: firstBreak(s.ast).id }));
        expectStartedSecondSection(s, before);
        expect(s.ast.sections[1].nodes[0].type).toBe('chart');
      });

      it('an image starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, { action: 'image', operation: 'insert', anchor: firstBreak(s.ast).id, fileData: PNG, width: 914400, height: 914400 });
        expectStartedSecondSection(s, before);
        expect(s.ast.sections[1].nodes[0].type).toBe('image');
      });

      it('an equation starts the second section', async () => {
        const s = await open(source.build());
        const before = counts(s.ast);
        await apply(s, { action: 'equation', operation: 'insert', anchor: firstBreak(s.ast).id, position: 'after', latex: 'x^2' });
        expectStartedSecondSection(s, before);
        expect(s.ast.sections[1].nodes[0].type).toBe('equation');
      });

      it('a moved block starts the second section', async () => {
        const s = await open(source.build());
        const mover = lastSection(s.ast).nodes[0];
        await apply(s, { action: 'move', target: mover.id, anchor: firstBreak(s.ast).id, position: 'after' });
        expectModelMatchesFile(s);
        expect(s.ast.sections[1].nodes[0].id).toBe(mover.id);
      });

      it('a new section break goes in the second section and splits it', async () => {
        const s = await open(source.build());
        const sectionCount = s.ast.sections.length;
        await apply(s, { action: 'insert', type: 'sectionBreak', anchor: firstBreak(s.ast).id, position: 'after' });
        expectModelMatchesFile(s);
        expect(s.ast.sections).toHaveLength(sectionCount + 1);
      });
    });

    it('a block added to the document behind the model is registered where it stands', async () => {
      const s = await open(source.build());
      const xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>LATE</w:t></w:r></w:p>';
      const id = s.parser.insertRootNode(xml, secondSectionFirst(s.ast).id, 'after', 'late_node');
      await s.parser.updateNodeInAST(s.ast, id, xml);
      expectModelMatchesFile(s);
      expect(s.ast.sections[1].nodes.some((n) => n.id === 'late_node')).toBe(true);
    });

    it('a block added behind the model right after a break paragraph is registered in the next section', async () => {
      const s = await open(source.build());
      const xml = '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>LATE</w:t></w:r></w:p>';
      const id = s.parser.insertRootNode(xml, firstBreak(s.ast).id, 'after', 'late_node');
      await s.parser.updateNodeInAST(s.ast, id, xml);
      expectModelMatchesFile(s);
      expect(s.ast.sections[1].nodes[0].id).toBe('late_node');
    });

    it('an equation next to a block of the second section stays in the second section', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'equation', operation: 'insert', anchor: secondSectionFirst(s.ast).id, position: 'after', latex: 'x^2' });
      expectModelMatchesFile(s);
    });

    it('a block moved before the first block of the second section is in the second section', async () => {
      const s = await open(source.build());
      const mover = lastSection(s.ast).nodes.slice(-1)[0];
      // The first block of the second section is not the mover (it would be moved relative to itself).
      if (secondSectionFirst(s.ast).id === mover.id) return;
      await apply(s, { action: 'move', target: mover.id, anchor: secondSectionFirst(s.ast).id, position: 'before' });
      expectModelMatchesFile(s);
      expect(s.ast.sections[1].nodes[0].id).toBe(mover.id);
    });

    it('insert after the last block of the document is the end of the document', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'insert', type: 'paragraph', anchor: lastBlock(s.ast).id, position: 'after', text: 'END' });
      expectModelMatchesFile(s);
      expect(lastSection(s.ast).nodes.slice(-1)[0].xml.textContent).toBe('END');
    });

    it('a second insert with no anchor goes after the first', async () => {
      const s = await open(source.build());
      await apply(s, { action: 'insert', type: 'paragraph', text: 'ONE' });
      await apply(s, { action: 'insert', type: 'paragraph', text: 'TWO' });
      expectModelMatchesFile(s);
      expect(lastSection(s.ast).nodes.slice(-2).map((n) => n.xml.textContent)).toEqual(['ONE', 'TWO']);
    });
  });
}

describe('a file whose last section holds no blocks', () => {
  it('insert with no anchor lands in the empty last section, after the break paragraph', async () => {
    const s = await open(SECTION_SHAPES['trailing-empty-section'].build());
    expect(s.ast.sections.map((x) => x.nodes.length)).toEqual([2, 0]);
    await apply(s, { action: 'insert', type: 'paragraph', text: 'IN LAST' });
    expectModelMatchesFile(s);
    expect(s.ast.sections.map((x) => x.nodes.length)).toEqual([2, 1]);
  });

  it('a callout and a table of contents with no anchor do not crash', async () => {
    const s = await open(SECTION_SHAPES['trailing-empty-section'].build());
    await apply(s, { action: 'callout', text: 'Note', title: 'T' });
    await apply(s, { action: 'field', operation: 'insertTOC' });
    expectModelMatchesFile(s);
  });
});

describe('moving several blocks keeps the order they were named in', () => {
  it('after the break paragraph of the first section', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const [a, b] = s.ast.sections[1].nodes;
    await apply(s, { action: 'move', target: [b.id, a.id], anchor: firstBreak(s.ast).id, position: 'after' });
    expectModelMatchesFile(s);
    expect(s.ast.sections[1].nodes.map((n) => n.id)).toEqual([b.id, a.id]);
  });

  it('into the middle of the first section, leaving the second empty', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const [a, b] = s.ast.sections[1].nodes;
    const [firstOfOne] = s.ast.sections[0].nodes;
    await apply(s, { action: 'move', target: [b.id, a.id], anchor: firstOfOne.id, position: 'after' });
    expectModelMatchesFile(s);
    expect(s.ast.sections[0].nodes.slice(1, 3).map((n) => n.id)).toEqual([b.id, a.id]);
    expect(s.ast.sections[1].nodes).toHaveLength(0);
  });

  it('before a block, in a one-section document', async () => {
    const s = await open(SECTION_SHAPES.single.build());
    const [alpha, beta] = s.ast.sections[0].nodes;
    await apply(s, { action: 'move', target: beta.id, anchor: alpha.id, position: 'before' });
    expectModelMatchesFile(s);
    expect(s.ast.sections[0].nodes.map((n) => n.id)).toEqual([beta.id, alpha.id]);
  });
});

describe('one-section documents place content exactly as before', () => {
  it('no anchor still goes to the end of the only section', async () => {
    const s = await open(SECTION_SHAPES.single.build());
    await apply(s, { action: 'insert', type: 'paragraph', text: 'END' });
    await apply(s, { action: 'callout', text: 'Note', title: 'T' });
    await apply(s, { action: 'field', operation: 'insertTOC' });
    await apply(s, chart());
    expectModelMatchesFile(s);
    expect(s.ast.sections).toHaveLength(1);
  });
});
