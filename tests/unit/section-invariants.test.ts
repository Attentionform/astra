/**
 * The sections the model reports are always the sections the file would have.
 *
 * Every shape of document (hand-built, and made by python-docx and LibreOffice)
 * meets every kind of operation, once on its own and once chained in a single
 * session. After each step that succeeds:
 *
 *   I1  the model has as many sections as the saved file
 *   I2  each section holds the same blocks, in the same order, as the file puts in it
 *   I3  every block id is unique and resolves in the parser
 *   I4  save and reopen gives the same sections and blocks
 *   I5  no header or footer part is left that no section refers to
 *       (except after a section break is deleted: its section's references go
 *       with it, as in Word, and the parts stay in the package unreferenced;
 *       a test below pins that, and that the file stays valid)
 *
 * A step may be refused (an operation that does not apply to that shape), but it
 * may never leave the model and the file disagreeing, and a refused batch must
 * leave both exactly as they were.
 */
import AdmZip from 'adm-zip';
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine, assertValidDocx, type DocumentAST, type Node } from '@attentionform/astra-engine';
import { endsSection } from '../../packages/engine/src/parser/sections';
import { SECTION_SHAPES, blockSignature, countFileSections, fileBlocksBySection, orphanParts } from '../helpers/section-docs';

const FIXTURES = path.resolve(__dirname, '../fixtures/sections');
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

const shapes: Array<{ name: string; build: () => Buffer }> = [
  ...Object.entries(SECTION_SHAPES).map(([key, shape]) => ({ name: `shape ${key}`, build: shape.build })),
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

const flat = (ast: DocumentAST): Node[] => ast.sections.flatMap((section) => section.nodes);
const signatures = (ast: DocumentAST): string[][] => ast.sections.map((section) => section.nodes.map((n) => blockSignature(n.xml as Element)));
const isBreak = (n: Node): boolean => endsSection(n.xml as Element);
const lastFilled = (ast: DocumentAST) => [...ast.sections].reverse().find((section) => section.nodes.length > 0);
/** An ordinary paragraph (not a break) to point operations at. */
const plainParagraph = (ast: DocumentAST): Node | undefined =>
  flat(ast).find((n) => (n.xml as Element).tagName === 'w:p' && !isBreak(n) && n.type === 'paragraph');

async function expectInvariants(s: Session, label: string): Promise<void> {
  const saved = s.parser.rezip();
  // I1
  expect(s.ast.sections, `${label}: I1 section count`).toHaveLength(countFileSections(saved));
  // I2
  expect(signatures(s.ast), `${label}: I2 blocks per section`).toEqual(fileBlocksBySection(saved));
  // I3
  const ids = flat(s.ast).map((n) => n.id);
  expect(new Set(ids).size, `${label}: I3 unique ids`).toBe(ids.length);
  for (const node of flat(s.ast)) expect(s.parser.getNode(node.id), `${label}: I3 ${node.id} resolves`).toBe(node.xml);
  // I4
  const reopened = await new DocxParser(saved).getAST();
  expect(reopened.sections, `${label}: I4 reopened section count`).toHaveLength(s.ast.sections.length);
  expect(signatures(reopened), `${label}: I4 reopened blocks`).toEqual(signatures(s.ast));
  // I5
  expect(orphanParts(saved), `${label}: I5 unreferenced header/footer parts`).toEqual([]);
}

type Step = { name: string; op: (s: Session) => Record<string, unknown> | null };

const steps: Step[] = [
  { name: 'insert paragraph, no anchor', op: () => ({ action: 'insert', type: 'paragraph', text: 'P-END' }) },
  {
    name: 'insert paragraph after the first section break',
    op: (s) => (s.ast.sections.length > 1 && isBreak(s.ast.sections[0].nodes.slice(-1)[0]) ? { action: 'insert', type: 'paragraph', anchor: s.ast.sections[0].nodes.slice(-1)[0].id, position: 'after', text: 'P-AFTER-BREAK' } : null),
  },
  {
    name: 'insert paragraph before the first block of the last section',
    op: (s) => ({ action: 'insert', type: 'paragraph', anchor: lastFilled(s.ast)!.nodes[0].id, position: 'before', text: 'P-BEFORE-LAST' }),
  },
  { name: 'insert table, no anchor', op: () => ({ action: 'insert', type: 'table', headers: ['H'], rows: [['v']] }) },
  { name: 'callout, no anchor', op: () => ({ action: 'callout', text: 'Note', title: 'T' }) },
  { name: 'table of contents, no anchor', op: () => ({ action: 'field', operation: 'insertTOC' }) },
  { name: 'page number field, no anchor', op: () => ({ action: 'field', operation: 'insertPageNumber', target: 'body' }) },
  {
    name: 'chart after the first block of the last section',
    op: (s) => ({ action: 'chart', chartType: 'column', title: 'C', categories: ['a', 'b'], series: [{ name: 's', values: [1, 2] }], anchor: lastFilled(s.ast)!.nodes[0].id }),
  },
  {
    name: 'image after the first block of the last section',
    op: (s) => ({ action: 'image', operation: 'insert', anchor: lastFilled(s.ast)!.nodes[0].id, fileData: PNG, width: 914400, height: 914400 }),
  },
  {
    name: 'equation after the first block of the last section',
    op: (s) => ({ action: 'equation', operation: 'insert', anchor: lastFilled(s.ast)!.nodes[0].id, position: 'after', latex: 'x^2' }),
  },
  {
    name: 'footnote on a paragraph',
    op: (s) => {
      const p = plainParagraph(s.ast);
      return p ? { action: 'footnote', anchor: p.id, text: 'Note text' } : null;
    },
  },
  {
    name: 'make a paragraph a bullet item',
    op: (s) => {
      const p = plainParagraph(s.ast);
      return p ? { action: 'list', operation: 'create', target: p.id, listType: 'bullet' } : null;
    },
  },
  {
    name: 'new section break inside the last section',
    op: (s) => ({ action: 'insert', type: 'sectionBreak', anchor: lastFilled(s.ast)!.nodes[0].id, position: 'after' }),
  },
  {
    name: 'move the first block of the last section after the first section break',
    op: (s) => {
      const first = s.ast.sections[0].nodes.slice(-1)[0];
      const mover = lastFilled(s.ast)!.nodes[0];
      return s.ast.sections.length > 1 && isBreak(first) && mover !== first ? { action: 'move', target: mover.id, anchor: first.id, position: 'after' } : null;
    },
  },
  { name: 'page setup on the last section', op: (s) => ({ action: 'pageSetup', sectionIndex: s.ast.sections.length, orientation: 'landscape' }) },
  { name: 'header on the last section', op: (s) => ({ action: 'headerFooter', operation: 'set', targetType: 'header', sectionIndex: s.ast.sections.length, text: 'HDR' }) },
  { name: 'footer on the first section', op: () => ({ action: 'headerFooter', operation: 'set', targetType: 'footer', sectionIndex: 1, text: 'FTR' }) },
  { name: 'page border on the last section', op: (s) => ({ action: 'pageBorder', sectionIndex: s.ast.sections.length }) },
  { name: 'columns on the last section', op: (s) => ({ action: 'section', operation: 'setColumns', sectionIndex: s.ast.sections.length, columns: 2 }) },
  {
    name: 'delete a block that is not a break',
    op: (s) => {
      const victim = lastFilled(s.ast)!.nodes[0];
      return isBreak(victim) ? null : { action: 'delete', target: victim.id };
    },
  },
  {
    name: 'delete the first section break',
    op: (s) => {
      const first = s.ast.sections[0].nodes.slice(-1)[0];
      return s.ast.sections.length > 1 && first && isBreak(first) ? { action: 'delete', target: first.id } : null;
    },
  },
  { name: 'cover page', op: () => ({ action: 'coverPage', title: 'Cover' }) },
];

/** Runs one operation; returns whether it was applied. A refused batch must leave everything as it was. */
async function applyStep(s: Session, step: Step, label: string): Promise<boolean> {
  const op = step.op(s);
  if (!op) return false;
  const before = { sections: signatures(s.ast), xml: s.parser.readDocumentXml() };
  const out = await s.engine.executeBatch({ operations: [op] as never });
  if (!out.success) {
    expect(out.rollbackTriggered, `${label}: refused without rollback`).toBe(true);
    expect(signatures(s.ast), `${label}: refused batch changed the model`).toEqual(before.sections);
    expect(s.parser.readDocumentXml(), `${label}: refused batch changed the document`).toBe(before.xml);
    await expectInvariants(s, `${label} (after refusal)`);
    return false;
  }
  await expectInvariants(s, label);
  return true;
}

describe('each operation on each shape', () => {
  for (const shape of shapes) {
    describe(shape.name, () => {
      for (const step of steps) {
        it(step.name, async () => {
          const s = await open(shape.build());
          await expectInvariants(s, 'as opened');
          await applyStep(s, step, step.name);
        });
      }
    });
  }
});

describe('every operation in turn, in one session', () => {
  for (const shape of shapes) {
    it(shape.name, async () => {
      const s = await open(shape.build());
      await expectInvariants(s, 'as opened');
      let applied = 0;
      for (const step of steps) {
        if (await applyStep(s, step, `${shape.name} > ${step.name}`)) applied++;
      }
      // Most steps apply to most shapes; a run where nearly all were refused would prove nothing.
      expect(applied).toBeGreaterThanOrEqual(12);
      // The python-docx file takes about 12 seconds on a fast machine and about a minute on a CI runner.
    }, 180_000);
  }
});

describe('deleting a section break', () => {
  it('takes its section\'s header and footer with it: no part stays unreferenced, their text is gone from the file, and the file is valid', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    for (const targetType of ['header', 'footer'] as const) {
      const out = await s.engine.executeBatch({ operations: [{ action: 'headerFooter', operation: 'set', targetType, sectionIndex: 1, text: `FIRST ${targetType}` }] as never });
      expect(out.success).toBe(true);
    }
    expect(orphanParts(s.parser.rezip())).toEqual([]);

    const out = await s.engine.executeBatch({ operations: [{ action: 'delete', target: s.ast.sections[0].nodes.slice(-1)[0].id }] as never });
    expect(out.success).toBe(true);

    const saved = s.parser.rezip();
    expect(s.ast.sections).toHaveLength(1);
    expect(orphanParts(saved)).toEqual([]);
    expect(new AdmZip(saved).getEntries().some((e) => /FIRST (header|footer)/.test(e.getData().toString('latin1')))).toBe(false);
    await assertValidDocx(saved);
  });
});
