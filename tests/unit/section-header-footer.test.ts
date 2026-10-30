/**
 * Header, footer, page border and column settings land in the section they
 * name, on files that were opened with several sections, survive a save and
 * reopen, and never leave a header or footer part that no section refers to.
 * When a section has no section properties to hold the change, the batch fails
 * (and rolls back) instead of doing nothing.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, type DocumentAST } from '@attentionform/astra-engine';
import { SECTION_SHAPES, orphanParts, partText, refsIn, sectPrsOf } from '../helpers/section-docs';

const FIXTURES = path.resolve(__dirname, '../fixtures/sections');
const sources = [
  { name: 'two sections', build: SECTION_SHAPES['two-sections'].build },
  { name: 'three sections', build: SECTION_SHAPES['three-sections'].build },
  { name: 'table before a break', build: SECTION_SHAPES['table-before-break'].build },
  { name: 'break in a content control', build: SECTION_SHAPES['break-in-content-control'].build },
  { name: 'trailing empty section', build: SECTION_SHAPES['trailing-empty-section'].build },
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
const run = async (s: Session, op: Record<string, unknown>) => {
  const out = await s.engine.executeBatch({ operations: [op] as never });
  expect(out.success, JSON.stringify(out.errors)).toBe(true);
};

for (const source of sources) {
  describe(`headers and footers on a reopened file: ${source.name}`, () => {
    it('every section gets its own header and footer, and they survive save and reopen', async () => {
      const s = await open(source.build());
      const count = s.ast.sections.length;
      for (let n = 1; n <= count; n++) {
        await run(s, { action: 'headerFooter', operation: 'set', targetType: 'header', sectionIndex: n, text: `HEADER ${n}` });
        await run(s, { action: 'headerFooter', operation: 'set', targetType: 'footer', sectionIndex: n, text: `FOOTER ${n}` });
      }
      const saved = s.parser.rezip();
      const sectPrs = sectPrsOf(saved);
      expect(sectPrs).toHaveLength(count);
      sectPrs.forEach((sp, i) => {
        const headers = refsIn(sp, 'w:headerReference');
        const footers = refsIn(sp, 'w:footerReference');
        // A section that already had a header of its own keeps exactly one default; ours is among the section's references.
        expect(headers.map((id) => partText(saved, id)).some((t) => t.includes(`HEADER ${i + 1}`)), `header ${i + 1}`).toBe(true);
        expect(footers.map((id) => partText(saved, id)).some((t) => t.includes(`FOOTER ${i + 1}`)), `footer ${i + 1}`).toBe(true);
      });
      expect(orphanParts(saved)).toEqual([]);

      // Reopen the saved file: the references are still where they were.
      const reopened = sectPrsOf(new DocxParser(saved).rezip());
      expect(reopened.map((sp) => refsIn(sp, 'w:headerReference').length)).toEqual(sectPrs.map((sp) => refsIn(sp, 'w:headerReference').length));
    });

    it('a header on one section does not touch the others', async () => {
      const s = await open(source.build());
      const count = s.ast.sections.length;
      const before = sectPrsOf(s.parser.rezip()).map((sp) => refsIn(sp, 'w:headerReference').join(','));
      await run(s, { action: 'headerFooter', operation: 'set', targetType: 'header', sectionIndex: count, text: 'LAST ONLY' });
      const after = sectPrsOf(s.parser.rezip()).map((sp) => refsIn(sp, 'w:headerReference').join(','));
      expect(after.slice(0, count - 1)).toEqual(before.slice(0, count - 1));
      expect(after[count - 1]).not.toBe('');
    });

    it('removing a header from a section removes only that reference', async () => {
      const s = await open(source.build());
      const count = s.ast.sections.length;
      await run(s, { action: 'headerFooter', operation: 'set', targetType: 'header', sectionIndex: 1, text: 'ONE' });
      await run(s, { action: 'headerFooter', operation: 'set', targetType: 'header', sectionIndex: count, text: 'LAST' });
      await run(s, { action: 'headerFooter', operation: 'remove', targetType: 'header', sectionIndex: count });
      const sectPrs = sectPrsOf(s.parser.rezip());
      expect(refsIn(sectPrs[count - 1], 'w:headerReference').filter((id) => partText(s.parser.rezip(), id).includes('LAST'))).toEqual([]);
      expect(refsIn(sectPrs[0], 'w:headerReference').some((id) => partText(s.parser.rezip(), id).includes('ONE'))).toBe(true);
    });

    it('page border and columns land in the named section only', async () => {
      const s = await open(source.build());
      const count = s.ast.sections.length;
      await run(s, { action: 'pageBorder', sectionIndex: count });
      await run(s, { action: 'section', operation: 'setColumns', sectionIndex: count, columns: 2 });
      const sectPrs = sectPrsOf(s.parser.rezip());
      sectPrs.forEach((sp, i) => {
        const isTarget = i === count - 1;
        expect(sp.getElementsByTagName('w:pgBorders').length > 0, `border ${i + 1}`).toBe(isTarget);
        expect(Array.from(sp.getElementsByTagName('w:cols')).some((c) => (c as Element).getAttribute('w:num') === '2'), `cols ${i + 1}`).toBe(isTarget);
      });
    });
  });
}

describe('when a section has no section properties to hold the change', () => {
  /** A model with a section the document does not have, as a broken session would. */
  const withPhantomSection = async (): Promise<Session> => {
    const s = await open(SECTION_SHAPES['two-sections'].build());
    s.ast.sections.push({ id: 'sec_phantom', heading: null, nodes: [], pageSetup: null, version: 1, checksum: '' });
    return s;
  };

  it('a header fails the batch and leaves no part behind', async () => {
    const s = await withPhantomSection();
    const out = await s.engine.executeBatch({
      operations: [{ action: 'headerFooter', operation: 'set', targetType: 'header', sectionIndex: 3, text: 'NOWHERE' }] as never,
    });
    expect(out.success).toBe(false);
    expect(out.rollbackTriggered).toBe(true);
    expect(JSON.stringify(out.errors)).toMatch(/no section properties/);
    expect(orphanParts(s.parser.rezip())).toEqual([]);
    expect(new AdmZip(s.parser.rezip()).getEntries().some((e) => /word\/header\d*\.xml/.test(e.entryName))).toBe(false);
  });

  it('a page border fails the batch', async () => {
    const s = await withPhantomSection();
    const out = await s.engine.executeBatch({ operations: [{ action: 'pageBorder', sectionIndex: 3 }] as never });
    expect(out.success).toBe(false);
    expect(JSON.stringify(out.errors)).toMatch(/no section properties/);
  });

  it('columns fail the batch', async () => {
    const s = await withPhantomSection();
    const out = await s.engine.executeBatch({ operations: [{ action: 'section', operation: 'setColumns', sectionIndex: 3, columns: 2 }] as never });
    expect(out.success).toBe(false);
    expect(JSON.stringify(out.errors)).toMatch(/no section properties/);
  });
});
