/**
 * A table is one node, so the paragraph inside a cell has no id of its own.
 * `<tableId>:r<row>c<col>[:p<n>]` (1-based, as `inspect` prints the grid)
 * lets the text and formatting operations reach it. Every markup assertion
 * reads the saved package, never just the live tree.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  cellAddressesMatching,
  findNode,
  parseCellAddress,
  syncCellStandIns,
  type DocumentAST,
  type Node,
} from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

interface Fixture {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
  table: Node;
}

/** Blank doc + a 3-row x 3-col table (header Region/Revenue/Growth, then North and South). */
async function withTable(): Promise<Fixture> {
  const parser = DocxParser.createNewDocument('blank');
  const ast = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const anchor = ast.sections[0].nodes[0].id;
  const inserted = await engine.executeBatch({
    operations: [
      {
        action: 'insert',
        type: 'table',
        anchor,
        position: 'after',
        headers: ['Region', 'Revenue', 'Growth'],
        rows: [
          ['North', '120', '4%'],
          ['South', '98', '2%'],
        ],
      },
    ] as never,
  });
  expect(inserted.success).toBe(true);
  const table = ast.sections[0].nodes.find((n) => n.type === 'table') as Node;
  expect(table).toBeDefined();
  return { parser, ast, engine, table };
}

/**
 * The table as the AST holds it now. A rolled-back batch restores the package and
 * re-reads it, so the table comes back as a new node under a positional id (t_001).
 */
const tableOf = (f: Fixture): Node => f.ast.sections[0].nodes.find((n) => n.type === 'table') as Node;
const idOf = (f: Fixture): string => tableOf(f).id;

const run = (f: Fixture, operations: unknown[], extra: Record<string, unknown> = {}) =>
  f.engine.executeBatch({ operations: operations as never, ...extra });

const documentXml = (f: Fixture): string =>
  new AdmZip(f.parser.rezip()).readAsText('word/document.xml');

const cellText = (table: Node, row: number, col: number): string => {
  const tr = (table.xml as Element).getElementsByTagName('w:tr')[row - 1];
  const tc = Array.from(tr.childNodes as unknown as Element[]).filter((c) => c.tagName === 'w:tc')[col - 1];
  return tc.textContent ?? '';
};

/** Adds a second paragraph "Second" to the cell at row/col (1-based). */
function addSecondParagraph(table: Node, row: number, col: number): void {
  const tr = (table.xml as Element).getElementsByTagName('w:tr')[row - 1];
  const tc = Array.from(tr.childNodes as unknown as Element[]).filter((c) => c.tagName === 'w:tc')[col - 1];
  const doc = tc.ownerDocument;
  const p = doc.createElement('w:p');
  const r = doc.createElement('w:r');
  const t = doc.createElement('w:t');
  t.appendChild(doc.createTextNode('Second'));
  r.appendChild(t);
  p.appendChild(r);
  tc.appendChild(p);
}

describe('cell addresses', () => {
  it('parses table id, 1-based row and column, and an optional paragraph', () => {
    expect(parseCellAddress('t_001:r2c3')).toEqual({ tableId: 't_001', row: 2, col: 3 });
    expect(parseCellAddress('t_001:r2c3:p2')).toEqual({ tableId: 't_001', row: 2, col: 3, para: 2 });
    expect(parseCellAddress('3F2A9C1B')).toBeNull();
    expect(parseCellAddress('t_001:r2')).toBeNull();
  });

  it('replace edits one cell and leaves the others alone', async () => {
    const f = await withTable();
    const result = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c2`, text: '130' }]);
    expect(result.success).toBe(true);
    expect(cellText(tableOf(f), 2, 2)).toBe('130');
    expect(cellText(tableOf(f), 2, 1)).toBe('North');
    expect(cellText(tableOf(f), 3, 2)).toBe('98');
    const xml = documentXml(f);
    expect(xml).toContain('130');
    expect(xml).not.toContain('>120<');
    expect((await validateDocxBuffer(f.parser.rezip())).ok).toBe(true);
  });

  it('replace keeps the table text and checksum in step with the cell', async () => {
    const f = await withTable();
    const before = tableOf(f).checksum;
    await run(f, [{ action: 'replace', target: `${idOf(f)}:r3c1`, text: 'Central' }]);
    expect(tableOf(f).metadata.text).toContain('Central');
    expect(tableOf(f).metadata.text).not.toContain('South');
    expect(tableOf(f).checksum).not.toBe(before);
  });

  it('style formats the cell paragraph', async () => {
    const f = await withTable();
    const result = await run(f, [{ action: 'style', target: `${idOf(f)}:r2c1`, bold: true, color: 'C00000' }]);
    expect(result.success).toBe(true);
    const xml = documentXml(f);
    expect(xml).toMatch(/<w:b\/>|<w:b w:val="(true|1)"\/>/);
    expect(xml).toContain('C00000');
  });

  it('namedStyle applies a paragraph style to a cell', async () => {
    const f = await withTable();
    const result = await run(f, [
      { action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: `${idOf(f)}:r1c1` },
    ]);
    expect(result.success).toBe(true);
    expect(documentXml(f)).toContain('w:pStyle w:val="Heading1"');
  });

  it('clearFormat strips direct formatting from a cell', async () => {
    const f = await withTable();
    const target = `${idOf(f)}:r2c3`;
    await run(f, [{ action: 'style', target, italic: true, color: '0000FF' }]);
    expect(documentXml(f)).toContain('0000FF');
    const result = await run(f, [{ action: 'clearFormat', target }]);
    expect(result.success).toBe(true);
    expect(documentXml(f)).not.toContain('0000FF');
    expect(cellText(tableOf(f), 2, 3)).toBe('4%');
  });

  it('hyperlink wraps the cell text in a link', async () => {
    const f = await withTable();
    const result = await run(f, [
      { action: 'hyperlink', target: `${idOf(f)}:r2c1`, url: 'https://example.com/north' },
    ]);
    expect(result.success).toBe(true);
    expect(documentXml(f)).toContain('w:hyperlink');
    expect(documentXml(f)).toContain('https://example.com/north');
    expect(cellText(tableOf(f), 2, 1)).toContain('North');
  });

  it('comment anchors to the cell paragraph', async () => {
    const f = await withTable();
    const result = await run(f, [
      { action: 'comment', operation: 'insert', target: `${idOf(f)}:r3c2`, text: 'Check this figure' },
    ]);
    expect(result.success).toBe(true);
    const xml = documentXml(f);
    expect(xml).toContain('w:commentReference');
    expect(new AdmZip(f.parser.rezip()).readAsText('word/comments.xml')).toContain('Check this figure');
  });

  describe('merged cells', () => {
    async function merged(): Promise<Fixture> {
      const f = await withTable();
      const result = await run(f, [
        { action: 'table', operation: 'mergeCells', target: idOf(f), fromRow: 2, fromCol: 1, toRow: 2, toCol: 2 },
      ]);
      expect(result.success).toBe(true);
      return f;
    }

    it('edits the cell where the merge starts', async () => {
      const f = await merged();
      const result = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c1`, text: 'North + Revenue' }]);
      expect(result.success).toBe(true);
      expect(cellText(tableOf(f), 2, 1)).toContain('North + Revenue');
    });

    it('refuses a cell the merge covers and points at the restart cell', async () => {
      const f = await merged();
      const result = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c2`, text: 'x' }]);
      expect(result.success).toBe(false);
      const message = JSON.stringify(result.errors);
      expect(message).toMatch(/covered by a merged cell/);
      expect(message).toMatch(/where the merge starts/);
    });

    it('names where a two-by-two merge starts, whichever covered cell is addressed', async () => {
      const f = await withTable();
      const id = idOf(f);
      const merge = await run(f, [
        { action: 'table', operation: 'mergeCells', target: id, fromRow: 2, fromCol: 1, toRow: 3, toCol: 2 },
      ]);
      expect(merge.success).toBe(true);
      for (const covered of ['r2c2', 'r3c1', 'r3c2']) {
        const current = idOf(f);
        const result = await run(f, [{ action: 'replace', target: `${current}:${covered}`, text: 'x' }]);
        expect(result.success).toBe(false);
        expect(JSON.stringify(result.errors), covered).toContain(`Use ${current}:r2c1, where the merge starts`);
      }
      const ok = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c1`, text: 'Block' }]);
      expect(ok.success).toBe(true);
    });

    it('names the restart cell when a gridSpan cell is addressed by a column it spans', async () => {
      const f = await withTable();
      // Word writes a horizontal merge as one cell with gridSpan; build that shape by hand.
      const tr = (tableOf(f).xml as Element).getElementsByTagName('w:tr')[1];
      const cells = Array.from(tr.childNodes as unknown as Element[]).filter((c) => c.tagName === 'w:tc');
      const doc = tr.ownerDocument;
      const tcPr = doc.createElement('w:tcPr');
      const span = doc.createElement('w:gridSpan');
      span.setAttribute('w:val', '2');
      tcPr.appendChild(span);
      cells[0].insertBefore(tcPr, cells[0].firstChild);
      tr.removeChild(cells[1]);
      const id = idOf(f);
      const result = await run(f, [{ action: 'replace', target: `${id}:r2c2`, text: 'x' }]);
      expect(result.success).toBe(false);
      const message = JSON.stringify(result.errors);
      expect(message).toMatch(/merged cell that starts at column 1/);
      expect(message).toContain(`Use ${id}:r2c1.`);
      const ok = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c1`, text: 'Spanned' }]);
      expect(ok.success).toBe(true);
      expect(cellAddressesMatching(tableOf(f), (t) => t === 'Spanned')).toEqual([`${idOf(f)}:r2c1`]);
      expect(cellAddressesMatching(tableOf(f), (t) => t === '4%')).toEqual([`${idOf(f)}:r2c3`]);
    });
  });

  it('carries the comment, hyperlink and bookmark flags up to the table', async () => {
    const f = await withTable();
    expect(tableOf(f).metadata.hasComment).toBeFalsy();
    await run(f, [
      { action: 'comment', operation: 'insert', target: `${idOf(f)}:r2c1`, text: 'Note' },
      { action: 'hyperlink', target: `${idOf(f)}:r2c2`, url: 'https://example.com' },
      { action: 'bookmark', target: `${idOf(f)}:r3c1`, bookmarkName: 'south_row' },
    ]);
    const meta = tableOf(f).metadata;
    expect(meta.hasComment).toBe(true);
    expect(meta.hasHyperlink).toBe(true);
    expect(meta.hasBookmark).toBe(true);
    expect(documentXml(f)).toContain('south_row');
  });

  it('hands executors a stand-in whose xml cannot be swapped out from under the table', async () => {
    const f = await withTable();
    const standIn = findNode(f.ast, `${idOf(f)}:r2c1`) as Node;
    expect(standIn.xml.tagName).toBe('w:p');
    expect(() => {
      (standIn as { xml: unknown }).xml = {};
    }).toThrow(TypeError);
    await syncCellStandIns(f.ast);
  });

  describe('refusals', () => {
    it('asks which paragraph when a cell holds several, and :pN picks one', async () => {
      const f = await withTable();
      addSecondParagraph(tableOf(f), 2, 1);
      const id = idOf(f);
      const ambiguous = await run(f, [{ action: 'replace', target: `${id}:r2c1`, text: 'x' }]);
      expect(ambiguous.success).toBe(false);
      expect(JSON.stringify(ambiguous.errors)).toContain(`${id}:r2c1:p1`);
      const picked = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c1:p2`, text: 'Changed' }]);
      expect(picked.success).toBe(true);
      expect(cellText(tableOf(f), 2, 1)).toBe('NorthChanged');
    });

    it('rejects a paragraph number past the end', async () => {
      const f = await withTable();
      const result = await run(f, [{ action: 'replace', target: `${idOf(f)}:r2c1:p3`, text: 'x' }]);
      expect(result.success).toBe(false);
      expect(JSON.stringify(result.errors)).toMatch(/has 1 paragraph/);
    });

    it('rejects a row or column outside the table and says the numbering is 1-based', async () => {
      const f = await withTable();
      const row = await run(f, [{ action: 'replace', target: `${idOf(f)}:r4c1`, text: 'x' }]);
      expect(row.success).toBe(false);
      expect(JSON.stringify(row.errors)).toMatch(/row 4 is outside table .*\(it has 3 rows\)/);
      expect(JSON.stringify(row.errors)).toMatch(/1-based/);
      const col = await run(f, [{ action: 'replace', target: `${idOf(f)}:r1c4`, text: 'x' }]);
      expect(col.success).toBe(false);
      expect(JSON.stringify(col.errors)).toMatch(/column 4 is outside row 1/);
      const zero = await run(f, [{ action: 'replace', target: `${idOf(f)}:r0c1`, text: 'x' }]);
      expect(zero.success).toBe(false);
    });

    it('refuses an action that cannot act on a cell and names what can', async () => {
      const f = await withTable();
      const result = await run(f, [{ action: 'delete', target: `${idOf(f)}:r2c1` }]);
      expect(result.success).toBe(false);
      const text = JSON.stringify(result.errors);
      expect(text).toContain("'delete' cannot act on a cell address");
      expect(text).toContain('replace, style, comment, clearFormat, namedStyle, hyperlink');
      expect(f.ast.sections[0].nodes.some((n) => n.type === 'table')).toBe(true);
    });

    it('points the table action at the table id and the same row and column numbers', async () => {
      const f = await withTable();
      const result = await run(f, [
        { action: 'table', operation: 'updateCell', target: `${idOf(f)}:r3c2`, text: 'x' },
      ]);
      expect(result.success).toBe(false);
      expect(JSON.stringify(result.errors)).toContain(`rowIndex 3 and colIndex 2`);
    });

    it('leaves an unknown table alone: the target is simply not found', async () => {
      const f = await withTable();
      const result = await run(f, [{ action: 'replace', target: 't_999:r1c1', text: 'x' }]);
      expect(result.success).toBe(false);
      expect(JSON.stringify(result.errors)).toMatch(/not found/);
    });
  });

  describe('versions, dry runs and rollback', () => {
    it('guards with the table version and bumps it on every cell edit', async () => {
      const f = await withTable();
      const v = tableOf(f).version;
      const stale = await run(f, [
        { action: 'replace', target: `${idOf(f)}:r2c1`, text: 'x', expectedVersion: v + 5 },
      ]);
      expect(stale.success).toBe(false);
      expect(JSON.stringify(stale.errors)).toMatch(/Version mismatch/);
      expect(tableOf(f).version).toBe(v);
      const fresh = await run(f, [
        { action: 'replace', target: `${idOf(f)}:r2c1`, text: 'Fresh', expectedVersion: v },
      ]);
      expect(fresh.success).toBe(true);
      expect(tableOf(f).version).toBeGreaterThan(v);
    });

    it('a dry run reports the table as touched and changes nothing', async () => {
      const f = await withTable();
      const before = documentXml(f);
      const id = idOf(f);
      const result = await run(f, [{ action: 'replace', target: `${id}:r2c1`, text: 'Preview' }], {
        dryRun: true,
      });
      expect(result.success).toBe(true);
      expect(result.preview?.some((p) => p.nodeId === id)).toBe(true);
      expect(cellText(tableOf(f), 2, 1)).toBe('North');
      expect(documentXml(f)).toBe(before);
    });

    it('rolls the cell edit back when a later operation in the batch fails', async () => {
      const f = await withTable();
      const version = tableOf(f).version;
      const result = await run(
        f,
        [
          { action: 'replace', target: `${idOf(f)}:r2c1`, text: 'Doomed' },
          { action: 'replace', target: 'no_such_node', text: 'x' },
        ],
        { rollbackOnError: true },
      );
      expect(result.success).toBe(false);
      expect(result.rollbackTriggered).toBe(true);
      expect(cellText(tableOf(f), 2, 1)).toBe('North');
      expect(tableOf(f).version).toBe(version);
    });
  });

  describe('search', () => {
    it('lists the cells that hold the match, in reading order', async () => {
      const f = await withTable();
      expect(cellAddressesMatching(tableOf(f), (t) => t.toLowerCase().includes('north'))).toEqual([`${idOf(f)}:r2c1`]);
      expect(cellAddressesMatching(tableOf(f), (t) => /%/.test(t))).toEqual([`${idOf(f)}:r2c3`, `${idOf(f)}:r3c3`]);
    });

    it('names the paragraph when a cell holds several, and skips covered cells', async () => {
      const f = await withTable();
      addSecondParagraph(tableOf(f), 2, 1);
      expect(cellAddressesMatching(tableOf(f), (t) => t === 'Second')).toEqual([`${idOf(f)}:r2c1:p2`]);
      await run(f, [
        { action: 'table', operation: 'mergeCells', target: idOf(f), fromRow: 3, fromCol: 2, toRow: 3, toCol: 3 },
      ]);
      // Row 3 now merges columns 2 and 3: the restart cell (r3c2) still answers, r3c3 is covered.
      const all = cellAddressesMatching(tableOf(f), () => true, 20);
      expect(all).toContain(`${idOf(f)}:r3c2`);
      expect(all).not.toContain(`${idOf(f)}:r3c3`);
    });

    it('caps the list', async () => {
      const f = await withTable();
      expect(cellAddressesMatching(tableOf(f), () => true, 2)).toHaveLength(2);
    });

    it('an address from search is accepted as an edit target', async () => {
      const f = await withTable();
      const [address] = cellAddressesMatching(tableOf(f), (t) => t === 'South');
      const result = await run(f, [{ action: 'replace', target: address, text: 'Southeast' }]);
      expect(result.success).toBe(true);
      expect(cellText(tableOf(f), 3, 1)).toBe('Southeast');
    });
  });
});
