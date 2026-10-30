/**
 * The model's sections follow the section breaks the document has. Operations
 * that add, delete or move a break paragraph change where the file's sections
 * end; after every batch the model is checked against the document and, if they
 * no longer agree, regrouped at the file's breaks, with a warning saying so.
 *
 * Two layers: `reconcileSections` itself on a real parsed document that is edited
 * by hand, and the engine running real operations on reopened multi-section files.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine, type DocumentAST } from '@attentionform/astra-engine';
import { reconcileSections } from '../../packages/engine/src/parser/sections';
import { SECTION_SHAPES, blockSignature, countFileSections, fileBlocksBySection } from '../helpers/section-docs';

const FIXTURES = path.resolve(__dirname, '../fixtures/sections');

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
const bodyOf = (s: Session): Element => s.parser.getDocumentDom().getElementsByTagName('w:body')[0] as Element;
const modelBlocks = (ast: DocumentAST): string[][] => ast.sections.map((section) => section.nodes.map((n) => blockSignature(n.xml as Element)));
const expectModelMatchesFile = (s: Session): void => {
  const saved = s.parser.rezip();
  expect(s.ast.sections).toHaveLength(countFileSections(saved));
  expect(modelBlocks(s.ast)).toEqual(fileBlocksBySection(saved));
};
const breakOf = (s: Session, sectionIndex: number) => {
  const nodes = s.ast.sections[sectionIndex].nodes;
  return nodes[nodes.length - 1];
};
const ids = (ast: DocumentAST): string[] => ast.sections.map((section) => section.id);

/** What an operation that deletes a block does: out of the document and out of its section. */
const removeBlock = (s: Session, id: string): void => {
  for (const section of s.ast.sections) {
    const index = section.nodes.findIndex((n) => n.id === id);
    if (index !== -1) {
      const [node] = section.nodes.splice(index, 1);
      node.xml.parentNode?.removeChild(node.xml);
      return;
    }
  }
  throw new Error(`no block ${id}`);
};

describe('reconcileSections on an edited document', () => {
  it('does nothing, and returns null, when the sections already match', async () => {
    const s = await open(SECTION_SHAPES['three-sections'].build());
    const sections = [...s.ast.sections];
    const note = reconcileSections(s.ast.sections, bodyOf(s));
    expect(note).toBeNull();
    expect(s.ast.sections).toHaveLength(3);
    s.ast.sections.forEach((section, i) => expect(section).toBe(sections[i]));
  });

  it('merges two sections when the break between them is deleted, keeping the first one\'s id', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const [first, second] = s.ast.sections;
    removeBlock(s, breakOf(s, 0).id);
    const note = reconcileSections(s.ast.sections, bodyOf(s));
    expect(note).toMatch(/rebuilt/);
    expect(note).toMatch(/now has 1 section, it had 2/);
    expect(s.ast.sections).toHaveLength(1);
    expect(s.ast.sections[0].id).toBe('sec_root');
    expect(s.ast.sections[0].nodes).toHaveLength(first.nodes.length + second.nodes.length);
    expectModelMatchesFile(s);
  });

  it('takes the page setup of the section after a deleted break', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    expect(s.ast.sections.map((x) => x.pageSetup?.paperSize)).toEqual(['Letter', 'A4']);
    removeBlock(s, breakOf(s, 0).id);
    reconcileSections(s.ast.sections, bodyOf(s));
    expect(s.ast.sections[0].pageSetup?.paperSize).toBe('A4');
  });

  it('keeps the ids of the sections that remain when the first one loses all its blocks', async () => {
    const s = await open(SECTION_SHAPES['three-sections'].build());
    expect(ids(s.ast)).toEqual(['sec_root', 'sec_2', 'sec_3']);
    removeBlock(s, breakOf(s, 0).id); // section 1 held only its break paragraph
    reconcileSections(s.ast.sections, bodyOf(s));
    expect(ids(s.ast)).toEqual(['sec_2', 'sec_3']);
    expectModelMatchesFile(s);
  });

  it('splits a section when a break is added, giving the new part a new id', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const target = s.ast.sections[1].nodes[0]; // "Two A": becomes the end of a new section
    const pPr = (target.xml as Element).getElementsByTagName('w:pPr')[0] ?? (() => {
      const created = target.xml.ownerDocument!.createElement('w:pPr');
      target.xml.insertBefore(created, target.xml.firstChild);
      return created;
    })();
    pPr.appendChild(target.xml.ownerDocument!.createElement('w:sectPr'));
    const note = reconcileSections(s.ast.sections, bodyOf(s));
    expect(note).toMatch(/now has 3 sections, it had 2/);
    expect(s.ast.sections).toHaveLength(3);
    expect(s.ast.sections[0].id).toBe('sec_root');
    expect(s.ast.sections[1].id).not.toBe('sec_root');
    expect(new Set(ids(s.ast)).size).toBe(3);
    expectModelMatchesFile(s);
  });

  it('keeps every other block, once, in its order', async () => {
    const s = await open(SECTION_SHAPES['three-sections'].build());
    const removed = breakOf(s, 1).id;
    const before = s.ast.sections.flatMap((section) => section.nodes.map((n) => n.id)).filter((id) => id !== removed);
    removeBlock(s, removed);
    reconcileSections(s.ast.sections, bodyOf(s));
    expect(s.ast.sections.flatMap((section) => section.nodes.map((n) => n.id))).toEqual(before);
  });

  it('bumps the version of a section whose blocks changed, and only that one', async () => {
    const s = await open(SECTION_SHAPES['three-sections'].build());
    const versions = s.ast.sections.map((section) => section.version);
    removeBlock(s, breakOf(s, 1).id); // sections 2 and 3 become one
    reconcileSections(s.ast.sections, bodyOf(s));
    expect(ids(s.ast)).toEqual(['sec_root', 'sec_2']);
    expect(s.ast.sections[0].version).toBe(versions[0]);
    expect(s.ast.sections[1].version).toBe(versions[1] + 1);
  });

  it('refuses to regroup when a break paragraph is not a block the model holds, and says so', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    // The document gets a break the model does not know: drop the paragraph from the model only.
    const hidden = breakOf(s, 0);
    s.ast.sections[0].nodes.pop();
    const snapshot = modelBlocks(s.ast);
    const note = reconcileSections(s.ast.sections, bodyOf(s));
    expect(note).toMatch(/could not be checked/);
    expect(modelBlocks(s.ast)).toEqual(snapshot);
    expect(hidden.xml.parentNode).not.toBeNull();
  });

  it('is a no-op without a body or without sections', async () => {
    const s = await open(SECTION_SHAPES.single.build());
    expect(reconcileSections(s.ast.sections, null)).toBeNull();
    expect(reconcileSections([], bodyOf(s))).toBeNull();
  });
});

const sources = [
  { name: 'two sections', build: SECTION_SHAPES['two-sections'].build },
  { name: 'three sections', build: SECTION_SHAPES['three-sections'].build },
  { name: 'table before a break', build: SECTION_SHAPES['table-before-break'].build },
  { name: 'break in a content control', build: SECTION_SHAPES['break-in-content-control'].build },
  { name: 'python-docx', build: () => fs.readFileSync(path.join(FIXTURES, 'python-docx-three-sections.docx')) },
  { name: 'libreoffice', build: () => fs.readFileSync(path.join(FIXTURES, 'libreoffice-two-page-styles.docx')) },
];

const RECONCILE_NOTE = /rebuilt from it|could not be checked against/;

const run = async (s: Session, op: Record<string, unknown>) => {
  const out = await s.engine.executeBatch({ operations: [op] as never });
  expect(out.success, JSON.stringify(out.errors)).toBe(true);
  return out;
};

describe('operations that delete or move a break paragraph', () => {
  for (const source of sources) {
    it(`${source.name}: deleting the first break merges the sections and says so`, async () => {
      const s = await open(source.build());
      const count = s.ast.sections.length;
      const out = await run(s, { action: 'delete', target: breakOf(s, 0).id });
      expect(s.ast.sections).toHaveLength(count - 1);
      expect(out.warnings.some((w) => RECONCILE_NOTE.test(w))).toBe(true);
      expectModelMatchesFile(s);
    });

    it(`${source.name}: moving the first break to the end of the document keeps the model in step`, async () => {
      const s = await open(source.build());
      const last = s.ast.sections.flatMap((section) => section.nodes).slice(-1)[0];
      const moving = breakOf(s, 0);
      if (moving.id === last.id) return;
      await run(s, { action: 'move', target: moving.id, anchor: last.id, position: 'after' });
      expectModelMatchesFile(s);
    });

    it(`${source.name}: deleting a block that is not a break leaves the sections alone, with no warning`, async () => {
      const s = await open(source.build());
      // The last section ends at the body, so its first block is never a break paragraph.
      const victim = s.ast.sections[s.ast.sections.length - 1].nodes[0];
      expect(victim).toBeDefined();
      const count = s.ast.sections.length;
      const out = await run(s, { action: 'delete', target: victim.id });
      expect(out.warnings.some((w) => RECONCILE_NOTE.test(w))).toBe(false);
      expect(s.ast.sections).toHaveLength(count);
      expectModelMatchesFile(s);
    });
  }
});

describe('operations that keep the sections consistent produce no warning', () => {
  for (const source of sources) {
    it(`${source.name}: insert paragraph, new section break, page setup`, async () => {
      const s = await open(source.build());
      const anchor = s.ast.sections[1].nodes[0].id;
      const ops: Array<Record<string, unknown>> = [
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'X' },
        { action: 'insert', type: 'sectionBreak', anchor, position: 'after' },
        { action: 'pageSetup', sectionIndex: 2, orientation: 'landscape' },
      ];
      for (const op of ops) {
        const out = await run(s, op);
        expect(out.warnings.filter((w) => RECONCILE_NOTE.test(w)), JSON.stringify(op)).toEqual([]);
        expectModelMatchesFile(s);
      }
    });
  }

  it('section split and merge on a reopened file', async () => {
    const s = await open(SECTION_SHAPES['three-sections'].build());
    const target = s.ast.sections[1].nodes[1]; // "Two B"
    for (const op of [
      { action: 'section', operation: 'split', target: target.id },
      { action: 'section', operation: 'merge', sectionIndex: 2 },
    ]) {
      const out = await run(s, op);
      expect(out.warnings.filter((w) => RECONCILE_NOTE.test(w)), JSON.stringify(op)).toEqual([]);
      expectModelMatchesFile(s);
    }
  });
});

describe('the single-operation path', () => {
  it('reconciles after execute() too and returns the sentence', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const result = await s.engine.execute({ action: 'delete', target: breakOf(s, 0).id } as never);
    expect(result.warnings.some((w) => RECONCILE_NOTE.test(w))).toBe(true);
    expectModelMatchesFile(s);
  });
});

describe('a document with one section is never touched', () => {
  it('batches leave the section object and its id alone', async () => {
    const s = await open(SECTION_SHAPES.single.build());
    const section = s.ast.sections[0];
    const out = await run(s, { action: 'insert', type: 'paragraph', text: 'More' });
    expect(out.warnings.filter((w) => RECONCILE_NOTE.test(w))).toEqual([]);
    expect(s.ast.sections[0]).toBe(section);
    expect(section.id).toBe('sec_root');
  });
});
