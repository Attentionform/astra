/**
 * Sections are numbered from 1 everywhere: sectionIndex 1 is the first section
 * in headerFooter, pageSetup, pageBorder, section (setColumns, merge) and in the
 * header/footer routing of insert and field. 0 is refused with the valid range.
 */
import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine, type DocumentAST } from '@attentionform/astra-engine';
import { OperationSchema } from '../../packages/engine/src/operation-schema';

interface Fixture {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}

/** Blank doc split into two sections by a next-page section break. */
async function twoSections(): Promise<Fixture> {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const anchor = ast.sections[0].nodes[0].id;
  const result = await engine.executeBatch({
    operations: [
      { action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'nextPage' },
      { action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Second section text' },
    ] as never,
  });
  expect(result.success).toBe(true);
  expect(ast.sections.length).toBe(2);
  return { parser, ast, engine };
}

const orientations = (f: Fixture): Array<string | null> =>
  f.parser.getSectPrElements().map((sectPr) => sectPr.getElementsByTagName('w:pgSz')[0]?.getAttribute('w:orient') ?? null);

describe('sectionIndex counts from 1', () => {
  it('pageSetup sectionIndex 1 is the first section and 2 the second', async () => {
    const f = await twoSections();
    const first = await f.engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 1, orientation: 'landscape' }] as never });
    expect(first.success).toBe(true);
    expect(orientations(f)[0]).toBe('landscape');
    expect(orientations(f)[1]).not.toBe('landscape');
    const second = await f.engine.executeBatch({ operations: [{ action: 'pageSetup', sectionIndex: 2, orientation: 'landscape' }] as never });
    expect(second.success).toBe(true);
    expect(orientations(f)).toEqual(['landscape', 'landscape']);
  });

  it('0 and a number past the last section are refused with the valid range', async () => {
    for (const bad of [0, 3]) {
      const f = await twoSections();
      const out = await f.engine.executeBatch({ operations: [{ action: 'pageSetup', params: { sectionIndex: bad, orientation: 'landscape' } }] as never });
      expect(out.success).toBe(false);
      expect(JSON.stringify(out.errors)).toMatch(new RegExp(`sectionIndex ${bad} is out of range \\(document has 2 sections\\)`));
      expect(JSON.stringify(out.errors)).toMatch(/pass sectionIndex in \[1, 2\], counting from 1/);
      expect(orientations(f).every((o) => o !== 'landscape')).toBe(true);
    }
  });

  it('the schema refuses 0 and negatives before the engine runs', () => {
    for (const bad of [0, -1]) {
      const parsed = OperationSchema.safeParse({ action: 'pageSetup', sectionIndex: bad, orientation: 'landscape' });
      expect(parsed.success).toBe(false);
      expect(JSON.stringify(parsed.error?.issues)).toMatch(/sections count from 1/);
    }
    expect(OperationSchema.safeParse({ action: 'pageSetup', sectionIndex: 1, orientation: 'landscape' }).success).toBe(true);
  });

  it('headerFooter sectionIndex 2 writes the second section header', async () => {
    const f = await twoSections();
    const set = await f.engine.executeBatch({
      operations: [{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'SECOND', sectionIndex: 2 }] as never,
    });
    expect(set.success).toBe(true);
    const sectPrs = f.parser.getSectPrElements();
    expect(sectPrs[0].getElementsByTagName('w:headerReference').length).toBe(0);
    expect(sectPrs[1].getElementsByTagName('w:headerReference').length).toBeGreaterThan(0);
    const other = await twoSections();
    const bad = await other.engine.executeBatch({
      operations: [{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'X', sectionIndex: 3 }] as never,
    });
    expect(bad.success).toBe(false);
    expect(JSON.stringify(bad.errors)).toMatch(/sectionIndex 3 is out of range \(document has 2 sections\)/);
  });

  it('insert and field routing to a footer take the same numbers', async () => {
    const f = await twoSections();
    const anchor = f.ast.sections[0].nodes[0].id;
    const ok = await f.engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor, position: 'after', text: 'Footer note', target: 'footer', pageType: 'default', sectionIndex: 1 }] as never,
    });
    expect(ok.success).toBe(true);
    const zero = await twoSections();
    const zeroOut = await zero.engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: zero.ast.sections[0].nodes[0].id, position: 'after', text: 'never', params: { target: 'footer', sectionIndex: 0 } }] as never,
    });
    expect(zeroOut.success).toBe(false);
    expect(JSON.stringify(zeroOut.errors)).toMatch(/positive integer \(sections count from 1\)/);
    const past = await twoSections();
    const pastOut = await past.engine.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: past.ast.sections[0].nodes[0].id, position: 'after', text: 'never', target: 'footer', sectionIndex: 3 }] as never,
    });
    expect(pastOut.success).toBe(false);
    expect(JSON.stringify(pastOut.errors)).toMatch(/sectionIndex 3 is out of range \(document has 2 section\(s\)\)/);
  });

  it('section merge sectionIndex 1 joins the first section with the second; the last one cannot merge', async () => {
    const edge = await twoSections();
    const last = await edge.engine.executeBatch({ operations: [{ action: 'section', operation: 'merge', sectionIndex: 2 }] as never });
    expect(last.success).toBe(false);
    expect(JSON.stringify(last.errors)).toMatch(/pass sectionIndex in \[1, 1\]/);
    const f = await twoSections();
    const merged = await f.engine.executeBatch({ operations: [{ action: 'section', operation: 'merge', sectionIndex: 1 }] as never });
    expect(merged.success).toBe(true);
    expect(f.ast.sections.length).toBe(1);
  });

  it('pageBorder refuses a section that does not exist instead of using the first', async () => {
    const bad = await twoSections();
    const out = await bad.engine.executeBatch({ operations: [{ action: 'pageBorder', sectionIndex: 5, style: 'single' }] as never });
    expect(out.success).toBe(false);
    expect(JSON.stringify(out.errors)).toMatch(/sectionIndex 5 is out of range \(document has 2 sections\)/);
    const f = await twoSections();
    const ok = await f.engine.executeBatch({ operations: [{ action: 'pageBorder', sectionIndex: 2, style: 'single' }] as never });
    expect(ok.success).toBe(true);
  });
});
