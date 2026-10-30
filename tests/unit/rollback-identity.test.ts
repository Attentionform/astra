/**
 * A failed batch changes nothing the caller can see. The document is rewound to
 * its state before the batch and read again, and the blocks and sections get
 * back the ids and versions they had, so an id handed out by an earlier batch
 * (a client-chosen `md_0001`, a generated `_bk_new_...`, a section made in the
 * session) still resolves and an `expectedVersion` taken before the failure is
 * still valid after it. The same goes for a dry run, which rewinds a batch that
 * worked, and for a single operation that fails.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, OperationEngine, createBlankDocxBuffer, type DocumentAST } from '@attentionform/astra-engine';
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

const flat = (ast: DocumentAST) => ast.sections.flatMap((section) => section.nodes);
/** Everything a caller can have been handed: block ids and versions, section ids and versions. */
const identity = (ast: DocumentAST) => ({
  blocks: flat(ast).map((n) => [n.id, n.version]),
  sections: ast.sections.map((s) => [s.id, s.version, s.nodes.length]),
});

const GHOST = { action: 'delete', target: 'no_such_block', strict: true };

const apply = async (s: Session, ...operations: Array<Record<string, unknown>>) => {
  const out = await s.engine.executeBatch({ operations: operations as never, sequential: true });
  expect(out.success, JSON.stringify(out.errors)).toBe(true);
  return out;
};
/** A batch whose first operations work and whose last one fails, so there is something to rewind. */
const failAfter = async (s: Session, ...operations: Array<Record<string, unknown>>) => {
  const out = await s.engine.executeBatch({ operations: [...operations, GHOST] as never, sequential: true });
  expect(out.success).toBe(false);
  expect(out.rollbackTriggered).toBe(true);
  return out;
};

const sources: Array<{ name: string; build: () => Buffer }> = [
  { name: 'blank document', build: () => createBlankDocxBuffer('blank') },
  { name: 'two sections', build: SECTION_SHAPES['two-sections'].build },
  { name: 'three sections', build: SECTION_SHAPES['three-sections'].build },
  { name: 'table before a break', build: SECTION_SHAPES['table-before-break'].build },
  { name: 'python-docx', build: () => fs.readFileSync(path.join(FIXTURES, 'python-docx-three-sections.docx')) },
  { name: 'libreoffice', build: () => fs.readFileSync(path.join(FIXTURES, 'libreoffice-two-page-styles.docx')) },
];

for (const source of sources) {
  describe(`a failed batch keeps ids and versions: ${source.name}`, () => {
    it('after earlier work in the session', async () => {
      const s = await open(source.build());
      const anchor = flat(s.ast)[0].id;
      await apply(
        s,
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Generated id' },
        { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Chosen id', id: 'md_0001' },
        { action: 'insert', type: 'table', anchor, position: 'after', headers: ['A'], rows: [['1']], id: 'tbl_chosen' },
      );
      const before = identity(s.ast);
      const xmlBefore = s.parser.readDocumentXml();

      await failAfter(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Rolled back' });

      expect(identity(s.ast)).toEqual(before);
      expect(s.parser.readDocumentXml()).toBe(xmlBefore);
    });

    it('so ids returned earlier can still be used as anchors and the parser resolves them', async () => {
      const s = await open(source.build());
      const anchor = flat(s.ast)[0].id;
      await apply(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Chosen id', id: 'md_0001' });
      await apply(s, { action: 'insert', type: 'paragraph', text: 'Generated id' });
      const generated = flat(s.ast).find((n) => n.xml.textContent === 'Generated id')!.id;

      await failAfter(s);
      await failAfter(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Rolled back' });

      for (const id of ['md_0001', generated]) {
        expect(flat(s.ast).some((n) => n.id === id), id).toBe(true);
        expect(s.parser.getNode(id), id).toBeDefined();
        await apply(s, { action: 'insert', type: 'paragraph', anchor: id, position: 'after', text: `after ${id}` });
      }
    });

    it('so the DOM, the parser map and the model agree about every id', async () => {
      const s = await open(source.build());
      const anchor = flat(s.ast)[0].id;
      await apply(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Chosen id', id: 'md_0001' });
      await failAfter(s);
      for (const node of flat(s.ast)) {
        expect(s.parser.getNode(node.id), node.id).toBe(node.xml);
        expect((node.xml as Element).getAttribute('data-astra-id'), node.id).toBe(node.id);
      }
      expect(new Set(flat(s.ast).map((n) => n.id)).size).toBe(flat(s.ast).length);
    });

    it('so an expectedVersion taken before the failure is still valid', async () => {
      const s = await open(source.build());
      const nodes = flat(s.ast);
      const mover = nodes[0];
      const anchor = nodes[1];
      await apply(s, { action: 'move', target: mover.id, anchor: anchor.id, position: 'after' }); // bumps the version
      const version = flat(s.ast).find((n) => n.id === mover.id)!.version;
      expect(version).toBeGreaterThan(1);

      await failAfter(s);

      expect(flat(s.ast).find((n) => n.id === mover.id)!.version).toBe(version);
      await apply(s, { action: 'insert', type: 'paragraph', anchor: mover.id, position: 'after', text: 'Versioned', expectedVersion: version });
    });

    it('twice in a row, and with a dry run in between', async () => {
      const s = await open(source.build());
      const anchor = flat(s.ast)[0].id;
      await apply(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Chosen id', id: 'md_0001' });
      const before = identity(s.ast);

      await failAfter(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'One' });
      await failAfter(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Two' });
      const dry = await s.engine.executeBatch({
        operations: [{ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Preview only' }] as never,
        dryRun: true,
      });
      expect(dry.success).toBe(true);

      expect(identity(s.ast)).toEqual(before);
    });

    it('when a single operation fails', async () => {
      const s = await open(source.build());
      const anchor = flat(s.ast)[0].id;
      await apply(s, { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Chosen id', id: 'md_0001' });
      const before = identity(s.ast);
      await expect(s.engine.execute(GHOST as never)).rejects.toThrow();
      expect(identity(s.ast)).toEqual(before);
    });
  });
}

describe('a failed batch keeps the sections', () => {
  it('reopened file: same section ids, versions and page setup', async () => {
    const s = await open(SECTION_SHAPES['three-sections'].build());
    const before = s.ast.sections.map((section) => ({ id: section.id, version: section.version, pageSetup: section.pageSetup }));
    await failAfter(s, { action: 'insert', type: 'paragraph', text: 'Rolled back' });
    expect(s.ast.sections.map((section) => ({ id: section.id, version: section.version, pageSetup: section.pageSetup }))).toEqual(before);
    expect(s.ast.sections).toHaveLength(countFileSections(s.parser.rezip()));
  });

  it('a section made in the session keeps its generated id and the split', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const anchor = s.ast.sections[1].nodes[0].id;
    await apply(s, { action: 'insert', type: 'sectionBreak', anchor, position: 'after' });
    expect(s.ast.sections).toHaveLength(3);
    const ids = s.ast.sections.map((section) => section.id);
    expect(ids.slice(0, 2)).toEqual(['sec_root', 'sec_2']);
    expect(ids[2]).not.toMatch(/^sec_\d$/); // the session's hashed id, not a position

    await failAfter(s, { action: 'insert', type: 'paragraph', text: 'Rolled back' });

    expect(s.ast.sections.map((section) => section.id)).toEqual(ids);
    expect(s.ast.sections).toHaveLength(countFileSections(s.parser.rezip()));
    expect(s.ast.sections.map((section) => section.nodes.map((n) => blockSignature(n.xml as Element)))).toEqual(
      fileBlocksBySection(s.parser.rezip()),
    );
  });
});

describe('when the ids cannot be kept', () => {
  it('says so, and still leaves a consistent document', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    // The model lost a block the document still has, so the two cannot be matched block for block.
    s.ast.sections[1].nodes.pop();

    const out = await failAfter(s, { action: 'insert', type: 'paragraph', text: 'Rolled back' });

    expect(out.errors.map((e) => e.error).join('\n')).toMatch(/Ids were re-read from the document/);
    expect(flat(s.ast).length).toBe(4);
    expect(s.ast.sections).toHaveLength(countFileSections(s.parser.rezip()));
    for (const node of flat(s.ast)) expect(s.parser.getNode(node.id)).toBe(node.xml);
  });

  it('says nothing when nothing was lost', async () => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    const out = await failAfter(s, { action: 'insert', type: 'paragraph', text: 'Rolled back' });
    expect(out.errors.map((e) => e.error).join('\n')).not.toMatch(/Ids were re-read/);
  });
});
