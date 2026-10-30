/**
 * Table ergonomics wave: cell alignment + widths pinning + cell padding,
 * identity recursion for table-cell paragraphs, and the shading contrast
 * guardrail.
 *
 * Every markup assertion reads SAVED bytes (rezip -> word/document.xml),
 * never just the live DOM.
 */
import { describe, it, expect, vi } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executeTable,
  OperationEngine,
  validateOperationStrict,
  OperationSchema,
  isValidParaId,
  hexToRelativeLuminance,
  contrastRatio,
  TABLE_CONTRAST_MIN_RATIO,
} from '@attentionform/astra-engine';
import { handleExecuteBatch } from '../../packages/mcp-server/src/lib/mcp-tool-adapter';
import { createMcpSessionStore } from '../../packages/mcp-server/src/lib/session-store';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'text/xml');
}

/** Blank doc + one 3-row x 2-col table inserted live (headers A/B, rows 1/2, 3/4). */
async function freshTable(): Promise<{ parser: DocxParser; ast: any; tableId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['A', 'B'], rows: [['1', '2'], ['3', '4']],
  } as any, ast, parser);
  const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tblNode).toBeDefined();
  return { parser, ast, tableId: tblNode.id as string };
}

/** Blank doc + one 3-row x 3-col table (for widths pinning). */
async function freshTable3col(): Promise<{ parser: DocxParser; ast: any; tableId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['A', 'B', 'C'], rows: [['1', '2', '3'], ['4', '5', '6']],
  } as any, ast, parser);
  const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tblNode).toBeDefined();
  return { parser, ast, tableId: tblNode.id as string };
}

/** SAVED bytes: rezip -> word/document.xml (not the live DOM). */
function savedDocumentXml(parser: DocxParser, ast: any): string {
  const buf = parser.rezip(ast);
  return new AdmZip(buf).readAsText('word/document.xml');
}

function onlyTable(doc: Document): Element {
  const tbls = doc.getElementsByTagName('w:tbl');
  expect(tbls.length).toBe(1);
  return tbls[0] as Element;
}

function directTrs(tbl: Element): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < tbl.childNodes.length; i++) {
    const c = tbl.childNodes[i] as Element;
    if (c.nodeType === 1 && c.tagName === 'w:tr') out.push(c);
  }
  return out;
}

function directTcs(tr: Element): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < tr.childNodes.length; i++) {
    const c = tr.childNodes[i] as Element;
    if (c.nodeType === 1 && c.tagName === 'w:tc') out.push(c);
  }
  return out;
}

function directChildren(parent: Element, tagName: string): Element[] {
  const out: Element[] = [];
  for (let i = 0; i < parent.childNodes.length; i++) {
    const c = parent.childNodes[i] as Element;
    if (c.nodeType === 1 && c.tagName === tagName) out.push(c);
  }
  return out;
}

function cellJcVal(tc: Element): (string | null)[] {
  return directChildren(tc, 'w:p').map((p) => {
    const pPr = directChildren(p, 'w:pPr')[0];
    const jc = pPr?.getElementsByTagName('w:jc')[0] as Element | undefined;
    return jc ? jc.getAttribute('w:val') : null;
  });
}

function cellPadVals(tc: Element): Record<string, string | null> {
  const tcPr = directChildren(tc, 'w:tcPr')[0];
  const mar = tcPr?.getElementsByTagName('w:tcMar')[0] as Element | undefined;
  const out: Record<string, string | null> = {};
  for (const side of ['top', 'left', 'bottom', 'right']) {
    const el = mar ? directChildren(mar, `w:${side}`)[0] as Element | undefined : undefined;
    out[side] = el ? el.getAttribute('w:w') : null;
  }
  return out;
}

/** Merged fixture: row 0 = [gridSpan=2 SPAN_CELL, REGULAR_CELL]; row 1 = [restart, continue, tail]. */
async function freshTableWithMerges(): Promise<{ parser: DocxParser; ast: any; tableId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert',
    type: 'table',
    anchor,
    position: 'after',
    headers: ['SPAN_CELL', 'B_COL', 'REGULAR_CELL'],
    rows: [['RESTART_CELL', 'COVER_CELL', 'TAIL_CELL']],
  } as any, ast, parser);
  const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tblNode).toBeDefined();
  const tblEl = tblNode.xml as Element;
  const rows = directTrs(tblEl);

  const row0Cells = directTcs(rows[0]);
  let tcPr0 = directChildren(row0Cells[0], 'w:tcPr')[0];
  if (!tcPr0) {
    tcPr0 = row0Cells[0].ownerDocument.createElementNS(W_NS, 'w:tcPr');
    row0Cells[0].insertBefore(tcPr0, row0Cells[0].firstChild);
  }
  const gs = row0Cells[0].ownerDocument.createElementNS(W_NS, 'w:gridSpan');
  gs.setAttribute('w:val', '2');
  tcPr0.appendChild(gs);
  rows[0].removeChild(row0Cells[1]);

  const row1Cells = directTcs(rows[1]);
  let tcPr1_0 = directChildren(row1Cells[0], 'w:tcPr')[0];
  if (!tcPr1_0) {
    tcPr1_0 = row1Cells[0].ownerDocument.createElementNS(W_NS, 'w:tcPr');
    row1Cells[0].insertBefore(tcPr1_0, row1Cells[0].firstChild);
  }
  const hmRestart = row1Cells[0].ownerDocument.createElementNS(W_NS, 'w:hMerge');
  hmRestart.setAttribute('w:val', 'restart');
  tcPr1_0.appendChild(hmRestart);

  let tcPr1_1 = directChildren(row1Cells[1], 'w:tcPr')[0];
  if (!tcPr1_1) {
    tcPr1_1 = row1Cells[1].ownerDocument.createElementNS(W_NS, 'w:tcPr');
    row1Cells[1].insertBefore(tcPr1_1, row1Cells[1].firstChild);
  }
  const hmContinue = row1Cells[1].ownerDocument.createElementNS(W_NS, 'w:hMerge');
  hmContinue.setAttribute('w:val', 'continue');
  tcPr1_1.appendChild(hmContinue);

  return { parser, ast, tableId: tblNode.id as string };
}

/** Paint every run of one live cell with an explicit text color (for contrast tests). */
function paintCellText(ast: any, tableId: string, rowIndex: number, colIndex: number, color: string): void {
  const tblNode = ast.sections[0].nodes.find((n: any) => n.id === tableId);
  const tblEl = tblNode.xml as Element;
  const tc = directTcs(directTrs(tblEl)[rowIndex])[colIndex];
  const doc = tc.ownerDocument;
  for (const p of directChildren(tc, 'w:p')) {
    for (const r of directChildren(p, 'w:r')) {
      let rPr = directChildren(r, 'w:rPr')[0];
      if (!rPr) {
        rPr = doc.createElementNS(W_NS, 'w:rPr');
        r.insertBefore(rPr, r.firstChild);
      }
      let colorEl = rPr.getElementsByTagName('w:color')[0] as Element | undefined;
      if (!colorEl) {
        colorEl = doc.createElementNS(W_NS, 'w:color');
        rPr.appendChild(colorEl);
      }
      colorEl.setAttribute('w:val', color);
    }
  }
}

describe('cell alignment (saved bytes)', () => {
  it('per-cell center lands on exactly one cell paragraph set', async () => {
    const { parser, ast, tableId } = await freshTable();
    const res: any = await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 2, colIndex: 1, align: 'center' },
    } as any, ast);
    expect(res.warnings).toEqual([]);

    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(3);
    for (let r = 0; r < rows.length; r++) {
      const cells = directTcs(rows[r]);
      for (let c = 0; c < cells.length; c++) {
        const vals = cellJcVal(cells[c]);
        if (r === 1 && c === 0) {
          expect(vals.length).toBeGreaterThan(0);
          for (const v of vals) expect(v).toBe('center');
        } else {
          for (const v of vals) expect(v).toBeNull();
        }
      }
    }
  });

  it('per-column right aligns the whole logical column; row/col aliases work flat', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId, operation: 'setCellAlignment',
      col: 2, alignment: 'right',
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const rows = directTrs(tbl);
    for (const tr of rows) {
      const cells = directTcs(tr);
      for (const v of cellJcVal(cells[0])) expect(v).toBeNull();
      for (const v of cellJcVal(cells[1])) expect(v).toBe('right');
    }
  });

  it('merged table: logical cols resolve to span + regular cells; cover cell rejects loudly', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 1, colIndex: 3, align: 'right' },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 1, colIndex: 1, align: 'left' },
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const rows = directTrs(tbl);
    const row0Cells = directTcs(rows[0]);
    expect(row0Cells.length).toBe(2);
    for (const v of cellJcVal(row0Cells[0])) expect(v).toBe('left');
    for (const v of cellJcVal(row0Cells[1])) expect(v).toBe('right');

    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 2, colIndex: 2, align: 'center' },
    } as any, ast)).rejects.toThrow(
      'E_VALIDATION: setCellAlignment column index 2 targets a merged cover cell (hMerge continue) and is not directly editable; target the merge restart instead.'
    );
  });

  it('bad / missing align and non-integer coords are loud E_VALIDATION', async () => {
    const { ast, tableId } = await freshTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 1, colIndex: 1, align: 'justify' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*align/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 1, colIndex: 1 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*align/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 1.5, colIndex: 1, align: 'left' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*integer/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', rowIndex: 100, colIndex: 1, align: 'left' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*out of bounds/);
  });
});

describe('column widths pinning (saved bytes)', () => {
  it('explicit widths honored exactly + fixed layout pinned', async () => {
    const { parser, ast, tableId } = await freshTable3col();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setColumnWidths', widths: [1500, 2500, 3500] },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const grid = directChildren(tbl, 'w:tblGrid')[0];
    const gridCols = directChildren(grid, 'w:gridCol').map((g) => g.getAttribute('w:w'));
    expect(gridCols).toEqual(['1500', '2500', '3500']);

    const rows = directTrs(tbl);
    expect(rows.length).toBe(3);
    const expected = ['1500', '2500', '3500'];
    for (const tr of rows) {
      const cells = directTcs(tr);
      expect(cells.length).toBe(3);
      cells.forEach((tc, i) => {
        const tcW = tc.getElementsByTagName('w:tcW')[0] as Element;
        expect(tcW.getAttribute('w:w')).toBe(expected[i]);
        expect(tcW.getAttribute('w:type')).toBe('dxa');
      });
    }
    const tblPr = directChildren(tbl, 'w:tblPr')[0];
    const layout = tblPr.getElementsByTagName('w:tblLayout')[0] as Element;
    expect(layout.getAttribute('w:type')).toBe('fixed');
  });

  it('odd widths are preserved exactly, never re-averaged', async () => {
    const { parser, ast, tableId } = await freshTable3col();
    await executeTable({
      action: 'table', target: tableId, operation: 'setColumnWidths',
      columnWidths: [1111, 2222, 3333],
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const gridCols = directChildren(directChildren(tbl, 'w:tblGrid')[0], 'w:gridCol')
      .map((g) => g.getAttribute('w:w'));
    expect(gridCols).toEqual(['1111', '2222', '3333']);
    expect(gridCols.map(Number).reduce((a, b) => a + b, 0)).toBe(6666);
  });

  it('merged table grid rebuilt exactly from the widths array', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setColumnWidths', widths: [1000, 2000, 3000] },
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const gridCols = directChildren(directChildren(tbl, 'w:tblGrid')[0], 'w:gridCol')
      .map((g) => g.getAttribute('w:w'));
    expect(gridCols).toEqual(['1000', '2000', '3000']);
    const layout = tbl.getElementsByTagName('w:tblLayout')[0] as Element;
    expect(layout.getAttribute('w:type')).toBe('fixed');
  });

  it('empty / non-array / non-positive-integer widths are loud E_VALIDATION', async () => {
    const { ast, tableId } = await freshTable3col();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setColumnWidths', widths: [] },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*widths/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setColumnWidths', widths: 'wide' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*widths/);
    for (const bad of [[2000, -5, 3000], [2000, 0, 3000], [2000, 12.5, 3000], [2000, 'x' as any, 3000]]) {
      await expect(executeTable({
        action: 'table', target: tableId,
        params: { op: 'setColumnWidths', widths: bad },
      } as any, ast)).rejects.toThrow(/E_VALIDATION.*widths\[1\]/);
    }
  });
});

describe('cell padding (saved bytes)', () => {
  it('uniform padding lands on exactly one cell with w:type dxa', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', rowIndex: 3, colIndex: 2, padding: 120 },
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const rows = directTrs(tbl);
    for (let r = 0; r < rows.length; r++) {
      const cells = directTcs(rows[r]);
      for (let c = 0; c < cells.length; c++) {
        const pads = cellPadVals(cells[c]);
        if (r === 2 && c === 1) {
          expect(pads).toEqual({ top: '120', left: '120', bottom: '120', right: '120' });
          const mar = cells[c].getElementsByTagName('w:tcMar')[0] as Element;
          for (const side of ['top', 'left', 'bottom', 'right']) {
            expect(directChildren(mar, `w:${side}`)[0].getAttribute('w:type')).toBe('dxa');
          }
        } else {
          expect(cells[c].getElementsByTagName('w:tcMar').length).toBe(0);
        }
      }
    }
  });

  it('per-side object applies table-wide with side granularity', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', padding: { top: 100, left: 150, bottom: 200, right: 250 } },
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    for (const tr of directTrs(tbl)) {
      for (const tc of directTcs(tr)) {
        expect(cellPadVals(tc)).toEqual({ top: '100', left: '150', bottom: '200', right: '250' });
      }
    }
  });

  it('merged span cell padded via logical col; bad padding is loud E_VALIDATION', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', rowIndex: 1, colIndex: 2, padding: 80 },
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const row0Cells = directTcs(directTrs(tbl)[0]);
    expect(cellPadVals(row0Cells[0])).toEqual({ top: '80', left: '80', bottom: '80', right: '80' });
    expect(row0Cells[1].getElementsByTagName('w:tcMar').length).toBe(0);

    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', rowIndex: 2, colIndex: 2, padding: 80 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*merged cover cell/);
    for (const bad of [-10, 12.5]) {
      await expect(executeTable({
        action: 'table', target: tableId,
        params: { op: 'setCellPadding', rowIndex: 1, colIndex: 1, padding: bad },
      } as any, ast)).rejects.toThrow(/E_VALIDATION.*padding/);
    }
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', rowIndex: 1, colIndex: 1 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*padding/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', rowIndex: 1, colIndex: 1, padding: { top: -5 } },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*padding\.top/);
  });
});

describe('identity recursion for table-cell paragraphs', () => {
  it('reload assigns a stable id to every table-cell paragraph, resolvable via getNode/readNode', async () => {
    const { parser, ast } = await freshTable();
    const reloaded = new DocxParser(parser.rezip(ast));
    const cellIds = reloaded.getTableCellParagraphIds();
    // 3 rows x 2 cols, one paragraph per cell.
    expect(cellIds.length).toBe(6);
    for (const id of cellIds) {
      expect(isValidParaId(id)).toBe(true);
      expect(reloaded.isTableCellScopedId(id)).toBe(true);
      const el = reloaded.getNode(id);
      expect(el?.tagName).toBe('w:p');
      expect((el?.parentNode as Element)?.tagName).toBe('w:tc');
      expect(reloaded.readNode(id)).toContain('<w:p');
    }
    // Root-level consumers keep legacy shape: summary + AST skip cell ids.
    const summaryIds = new Set(reloaded.getStructuralSummary().nodes.map((n) => n.id));
    for (const id of cellIds) expect(summaryIds.has(id)).toBe(false);
    const reAst: any = await reloaded.getAST();
    const astIds = new Set(reAst.sections.flatMap((s: any) => s.nodes.map((n: any) => n.id)));
    for (const id of cellIds) expect(astIds.has(id)).toBe(false);
  });

  it('no collisions with the body map; ids survive edits and reloads', async () => {
    const { parser, ast } = await freshTable();
    const buf = parser.rezip(ast);
    const p1 = new DocxParser(buf);
    const ids1 = p1.getTableCellParagraphIds();
    expect(new Set(ids1).size).toBe(ids1.length);

    // Internal map holds exactly roots + cell ids (no dup overwrote a root).
    const internalSize = (p1 as any).nodeMap.size as number;
    const summaryCount = p1.getStructuralSummary().nodes.length;
    expect(internalSize).toBe(summaryCount + ids1.length);

    // Stable across a second reload (persisted w14:paraIds round-trip).
    const p2 = new DocxParser(p1.rezip(await p1.getAST()));
    expect(new Set(p2.getTableCellParagraphIds())).toEqual(new Set(ids1));

    // Stable across a cell edit + save + reload (traceability).
    const ast1: any = await p1.getAST();
    const tableId1 = ast1.sections[0].nodes.find((n: any) => n.type === 'table').id;
    await executeTable({
      action: 'table', target: tableId1,
      params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'edited' },
    } as any, ast1);
    const p3 = new DocxParser(p1.rezip(ast1));
    expect(new Set(p3.getTableCellParagraphIds())).toEqual(new Set(ids1));
  });
});

describe('contrast guardrail', () => {
  it('luminance + ratio math follows WCAG (black 0, white 1, navy dark)', () => {
    expect(TABLE_CONTRAST_MIN_RATIO).toBe(4.5);
    expect(hexToRelativeLuminance('000000')).toBe(0);
    expect(hexToRelativeLuminance('FFFFFF')).toBeCloseTo(1, 5);
    expect(contrastRatio('000000', 'FFFFFF')).toBeCloseTo(21, 0);
    // Navy (000080) vs black fails AA; vs white passes comfortably.
    expect(contrastRatio('000000', '000080')).toBeLessThan(4.5);
    expect(contrastRatio('FFFFFF', '000080')).toBeGreaterThan(4.5);
  });

  it('navy fill + black text WARNS in the result, but still applies', async () => {
    const { parser, ast, tableId } = await freshTable();
    const res: any = await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '000080' },
    } as any, ast);
    expect(res.warnings.length).toBe(1);
    expect(res.warnings[0]).toMatch(/contrast.*000080.*4\.5/);
    // Advisory only: the fill landed in saved bytes.
    expect(savedDocumentXml(parser, ast)).toContain('w:fill="000080"');
  });

  it('navy fill + white text passes silent', async () => {
    const { parser, ast, tableId } = await freshTable();
    paintCellText(ast, tableId, 0, 0, 'FFFFFF');
    const spy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const res: any = await executeTable({
        action: 'table', target: tableId,
        params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '000080' },
      } as any, ast);
      expect(res.warnings).toEqual([]);
      expect(spy).not.toHaveBeenCalled();
    } finally {
      spy.mockRestore();
    }
    expect(savedDocumentXml(parser, ast)).toContain('w:fill="000080"');
  });
});

describe('schema gate for the new surface', () => {
  it('strict schema accepts the new sub-ops with flat args', () => {
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setCellAlignment',
      rowIndex: 1, colIndex: 2, align: 'center',
    } as never).ok).toBe(true);
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setCellPadding',
      rowIndex: 1, colIndex: 2, padding: 100,
    } as never).ok).toBe(true);
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setCellPadding',
      padding: { top: 100, left: 150 },
    } as never).ok).toBe(true);
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setColumnWidths',
      columnWidths: [2000, 3000],
    } as never).ok).toBe(true);
  });

  it('strict schema rejects bad align, negative padding, unknown sub-op', () => {
    expect(OperationSchema.safeParse({
      action: 'table', target: 't_001', operation: 'setCellAlignment',
      rowIndex: 1, colIndex: 2, align: 'justify',
    }).success).toBe(false);
    expect(OperationSchema.safeParse({
      action: 'table', target: 't_001', operation: 'setCellPadding',
      rowIndex: 1, colIndex: 2, padding: -5,
    }).success).toBe(false);
    expect(() => validateOperationStrict({
      action: 'table', target: 't_001', operation: 'frobCells',
    })).toThrow(/E_VALIDATION/);
  });
});

describe('combined ergo doc passes the real OOXML validator', () => {
  it('alignment + pinned widths + padding + shading round-trip valid', async () => {
    const { parser, ast, tableId } = await freshTable3col();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', colIndex: 3, align: 'right' },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setColumnWidths', widths: [2000, 3000, 4000] },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', padding: { top: 50, left: 100, bottom: 50, right: 100 } },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: 'D9E2F3' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect(saved).toContain('w:val="right"');
    expect(saved).toContain('w:type="fixed"');
    expect(saved).toContain('<w:tcMar>');
    const res = await validateDocxBuffer(parser.rezip(ast));
    expect(res.ok).toBe(true);
  }, 120000);
});

function buildNestedTable(doc: Document, cellText: string): Element {
  const mk = (tag: string): Element => doc.createElementNS(W_NS, tag);
  const tbl = mk('w:tbl');
  const tblPr = mk('w:tblPr');
  tbl.appendChild(tblPr);
  const grid = mk('w:tblGrid');
  const gc = mk('w:gridCol');
  gc.setAttribute('w:w', '2000');
  grid.appendChild(gc);
  tbl.appendChild(grid);
  const tr = mk('w:tr');
  const tc = mk('w:tc');
  const p = mk('w:p');
  const r = mk('w:r');
  const t = mk('w:t');
  t.setAttribute('xml:space', 'preserve');
  t.textContent = cellText;
  r.appendChild(t);
  p.appendChild(r);
  tc.appendChild(p);
  tr.appendChild(tc);
  tbl.appendChild(tr);
  return tbl;
}

async function freshTwoBodyTables(): Promise<{ parser: DocxParser; ast: any; ids: string[] }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['A', 'B'], rows: [['1', '2']],
  } as any, ast, parser);
  const firstId = (ast.sections[0].nodes.find((n: any) => n.type === 'table').id as string);
  await executeInsert({
    action: 'insert', type: 'table', anchor: firstId, position: 'after',
    headers: ['C', 'D'], rows: [['3', '4']],
  } as any, ast, parser);
  const ids = ast.sections[0].nodes.filter((n: any) => n.type === 'table').map((n: any) => n.id as string);
  return { parser, ast, ids };
}

describe('nested tables keep legacy root shape', () => {
  it('nested tbl is cell-scoped: summary/AST counts identical, body t_ ids unchanged, nested paras resolvable', async () => {
    const control = await freshTwoBodyTables();
    const controlSummary = new DocxParser(control.parser.rezip(control.ast)).getStructuralSummary();
    const controlSummaryTables = controlSummary.nodes.filter((n) => (n as any).type === 'table');
    expect(controlSummaryTables.map((n) => n.id)).toEqual(['t_001', 't_002']);

    const live = await freshTwoBodyTables();
    const firstTblEl = (live.ast.sections[0].nodes.find((n: any) => n.id === live.ids[0]).xml as Element);
    const hostTc = directTcs(directTrs(firstTblEl)[0])[0];
    hostTc.appendChild(buildNestedTable(hostTc.ownerDocument, 'NESTED'));
    const reloaded = new DocxParser(live.parser.rezip(live.ast));

    const summary = reloaded.getStructuralSummary();
    const summaryTables = summary.nodes.filter((n) => (n as any).type === 'table');
    expect(summaryTables.map((n) => n.id)).toEqual(['t_001', 't_002']);
    expect(summary.nodes.length).toBe(controlSummary.nodes.length);

    const rast: any = await reloaded.getAST();
    const astTables = rast.sections[0].nodes.filter((n: any) => n.type === 'table');
    expect(astTables.map((n: any) => n.id)).toEqual(['t_001', 't_002']);
    const cast: any = await new DocxParser(control.parser.rezip(control.ast)).getAST();
    expect(rast.sections[0].nodes.length).toBe(cast.sections[0].nodes.length);

    expect(reloaded.getTrackedNodeIds().length).toBe(0);

    const cellIds = reloaded.getTableCellParagraphIds();
    const nestedPara = cellIds
      .map((id) => ({ id, el: reloaded.getNode(id) as Element | undefined }))
      .find(({ el }) => el?.textContent === 'NESTED');
    expect(nestedPara).toBeDefined();
    expect(reloaded.isTableCellScopedId(nestedPara!.id)).toBe(true);
    expect(reloaded.readNode(nestedPara!.id)).toContain('NESTED');

    const summaryIds = new Set(summary.nodes.map((n) => n.id));
    const astIds = new Set(rast.sections[0].nodes.map((n: any) => n.id));
    expect(summaryIds.has(nestedPara!.id)).toBe(false);
    expect(astIds.has(nestedPara!.id)).toBe(false);
    for (const id of cellIds) {
      expect(summaryIds.has(id)).toBe(false);
      expect(astIds.has(id)).toBe(false);
    }

    const internal = (reloaded as any).nodeMap as Map<string, Element>;
    let nestedTblId: string | undefined;
    for (const [id, el] of internal.entries()) {
      if (el.tagName === 'w:tbl' && (el.parentNode as Element)?.tagName === 'w:tc') {
        nestedTblId = id;
        break;
      }
    }
    expect(nestedTblId).toBeDefined();
    expect(reloaded.isTableCellScopedId(nestedTblId!)).toBe(true);
    expect(summaryIds.has(nestedTblId!)).toBe(false);
    expect(astIds.has(nestedTblId!)).toBe(false);
    expect(nestedTblId!.startsWith('t_cell_')).toBe(true);
  });
});

describe('warnings reach engine batch and MCP handler', () => {
  it('contrast warning arrives via OperationEngine batch + single path, not only direct', async () => {
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { parser, ast, tableId } = await freshTable();
      const engine = new OperationEngine(ast, parser);
      const res = await engine.executeBatch({
        operations: [{
          action: 'table', target: tableId,
          params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '000080' },
        } as any],
        sequential: true,
      });
      expect(res.success).toBe(true);
      expect(res.warnings.length).toBe(1);
      expect(res.warnings[0]).toMatch(/contrast.*000080/);

      const second = await freshTable();
      const engine2 = new OperationEngine(second.ast, second.parser);
      const single = await engine2.execute({
        action: 'table', target: second.tableId,
        params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '000080' },
      } as any);
      expect(single.warnings.length).toBe(1);
      expect(single.warnings[0]).toMatch(/contrast/);
    } finally {
      warnSpy.mockRestore();
    }
  });

  it('contrast warning arrives via MCP execute_batch warnings array', async () => {
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { parser, ast } = await freshTable();
      const buf = parser.rezip(ast);
      const store = createMcpSessionStore();
      const handle = await store.openBuffer(buf, 'warn-mcp.docx');
      const liveTableId = handle.ast.sections[0].nodes.find((n) => (n as any).type === 'table')!.id;
      const out = await handleExecuteBatch(store, {
        sessionId: handle.sessionId,
        operations: [{
          action: 'table', target: liveTableId,
          params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '000080' },
        }],
        sequential: true,
      });
      expect(out.ok).toBe(true);
      if (!out.ok) throw new Error('mcp batch failed');
      expect(out.warnings.length).toBe(1);
      expect(out.warnings[0]).toMatch(/contrast.*000080/);
    } finally {
      warnSpy.mockRestore();
    }
  });
});

describe('column-wide unknown columns fail loudly', () => {
  it('setCellAlignment/setCellPadding/setVerticalAlign reject out-of-range colIndex with valid range', async () => {
    const { ast, tableId } = await freshTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', colIndex: 10, align: 'center' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*column index 10.*\[1, 2\]/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellPadding', colIndex: 10, padding: 100 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*column index 10.*\[1, 2\]/);
    await expect(executeTable({
      action: 'table', target: tableId, operation: 'setVerticalAlign', colIndex: 10,
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*column index 10.*\[1, 2\]/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', colIndex: -1, align: 'center' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*column index -1/);
  });

  it('in-range covers still skip quietly for bulk column application', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellAlignment', colIndex: 2, align: 'center' },
    } as any, ast);
    const tbl = onlyTable(parseXml(savedDocumentXml(parser, ast)));
    const rows = directTrs(tbl);
    const row0Vals = cellJcVal(directTcs(rows[0])[0]);
    for (const v of row0Vals) expect(v).toBe('center');

    const second = await freshTableWithMerges();
    await executeTable({
      action: 'table', target: second.tableId,
      params: { op: 'setCellPadding', colIndex: 2, padding: 77 },
    } as any, second.ast);
    const tbl2 = onlyTable(parseXml(savedDocumentXml(second.parser, second.ast)));
    expect(cellPadVals(directTcs(directTrs(tbl2)[0])[0]).top).toBe('77');
  });
});
