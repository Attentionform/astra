/**
 * Phase E3 — table cell controls: setCellShading + setRowHeight.
 *
 * Every markup assertion reads SAVED bytes (rezip -> word/document.xml),
 * never just the live DOM (mergeCells-dropped lesson). Reload tests
 * round-trip save -> new DocxParser -> read back.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executeTable,
  validateOperationStrict,
  OperationSchema,
  resolveTableCell,
  findLogicalTableCell,
} from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';
import { useReproducibleOutput } from '../helpers/reproducible';

// The "leaves the document unchanged" tests compare package bytes before and after; without pinned
// output they carry the clock and differ whenever a 2-second zip-time tick falls between the two saves.
useReproducibleOutput();

const serializer = new XMLSerializer();

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

function firstElementChild(el: Element): Element | null {
  for (let i = 0; i < el.childNodes.length; i++) {
    const c = el.childNodes[i] as Element;
    if (c.nodeType === 1) return c;
  }
  return null;
}

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

  // Row 0: [gridSpan=2, regular_cell]
  const row0Cells = directTcs(rows[0]);
  let tcPr0 = directChildren(row0Cells[0], 'w:tcPr')[0];
  if (!tcPr0) {
    tcPr0 = row0Cells[0].ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:tcPr');
    row0Cells[0].insertBefore(tcPr0, row0Cells[0].firstChild);
  }
  const gs = row0Cells[0].ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:gridSpan');
  gs.setAttribute('w:val', '2');
  tcPr0.appendChild(gs);
  rows[0].removeChild(row0Cells[1]);

  // Row 1: [restart, continue, tail]
  const row1Cells = directTcs(rows[1]);
  let tcPr1_0 = directChildren(row1Cells[0], 'w:tcPr')[0];
  if (!tcPr1_0) {
    tcPr1_0 = row1Cells[0].ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:tcPr');
    row1Cells[0].insertBefore(tcPr1_0, row1Cells[0].firstChild);
  }
  const hmRestart = row1Cells[0].ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:hMerge');
  hmRestart.setAttribute('w:val', 'restart');
  tcPr1_0.appendChild(hmRestart);

  let tcPr1_1 = directChildren(row1Cells[1], 'w:tcPr')[0];
  if (!tcPr1_1) {
    tcPr1_1 = row1Cells[1].ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:tcPr');
    row1Cells[1].insertBefore(tcPr1_1, row1Cells[1].firstChild);
  }
  const hmContinue = row1Cells[1].ownerDocument.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:hMerge');
  hmContinue.setAttribute('w:val', 'continue');
  tcPr1_1.appendChild(hmContinue);

  return { parser, ast, tableId: tblNode.id as string };
}

describe('E3 setCellShading (saved bytes)', () => {
  it('writes w:shd on exactly one cell — SAVED xml, only that cell', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 2, colIndex: 1, fill: 'FF0000' },
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    expect(saved).toContain('w:fill="FF0000"');
    expect((saved.match(/w:fill="FF0000"/g) || []).length).toBe(1);

    const tbl = onlyTable(parseXml(saved));
    const shds = tbl.getElementsByTagName('w:shd');
    expect(shds.length).toBe(1);
    const shd = shds[0] as Element;
    expect(shd.getAttribute('w:val')).toBe('clear');
    expect(shd.getAttribute('w:color')).toBe('auto');
    expect(shd.getAttribute('w:fill')).toBe('FF0000');

    // Only cell (1,0): every other cell in the 3x2 grid has no w:shd.
    const rows = directTrs(tbl);
    expect(rows.length).toBe(3);
    for (let r = 0; r < rows.length; r++) {
      const cells = directTcs(rows[r]);
      expect(cells.length).toBe(2);
      for (let c = 0; c < cells.length; c++) {
        const n = cells[c].getElementsByTagName('w:shd').length;
        expect(n, `cell (${r},${c}) shd count`).toBe(r === 1 && c === 0 ? 1 : 0);
      }
    }
  });

  it('flat shape matches params shape (mergeCells flat-parity precedent)', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId, operation: 'setCellShading',
      rowIndex: 3, colIndex: 2, fill: '00FF00',
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect((saved.match(/w:fill="00FF00"/g) || []).length).toBe(1);
    const tbl = onlyTable(parseXml(saved));
    expect(tbl.getElementsByTagName('w:shd').length).toBe(1);
  });

  it('invalid hex is loud E_VALIDATION (incl short, non-hex, empty)', async () => {
    const { ast, tableId } = await freshTable();
    for (const fill of ['FFF', 'FFFFF', 'FF0000FF', 'GGGGGG', 'red', '', 'FF00ZZ', '#FFF', '#GGGGGG']) {
      await expect(executeTable({
        action: 'table', target: tableId,
        params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill },
      } as any, ast)).rejects.toThrow(/E_VALIDATION.*fill/);
    }
    // Missing fill is loud too (no silent default).
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, colIndex: 1 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*fill/);
  });

  it('leading-# hex is accepted and normalized (setShading precedent)', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, colIndex: 1, fill: '#FF0000' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect((saved.match(/w:fill="FF0000"/g) || []).length).toBe(1);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows[0] && directTcs(rows[0])[0].getElementsByTagName('w:shd').length).toBe(1);
    expect(directTcs(rows[1])[0].getElementsByTagName('w:shd').length).toBe(0);
  });

  it('lowercase hex normalizes to uppercase in saved bytes', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, colIndex: 2, fill: 'ab12cd' },
    } as any, ast);
    expect(savedDocumentXml(parser, ast)).toContain('w:fill="AB12CD"');
  });
});

describe('E3 setRowHeight (saved bytes + schema position)', () => {
  it('exact + default-atLeast land in saved xml with w:val + w:hRule', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 1, height: 600, rule: 'exact' },
    } as any, ast);
    // Flat shape, default rule (atLeast).
    await executeTable({
      action: 'table', target: tableId, operation: 'setRowHeight',
      rowIndex: 2, height: 400,
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    const doc = parseXml(saved);
    const tbl = onlyTable(doc);
    const rows = directTrs(tbl);
    const h0 = rows[0].getElementsByTagName('w:trHeight')[0] as Element;
    const h1 = rows[1].getElementsByTagName('w:trHeight')[0] as Element;
    expect(h0.getAttribute('w:val')).toBe('600');
    expect(h0.getAttribute('w:hRule')).toBe('exact');
    expect(h1.getAttribute('w:val')).toBe('400');
    expect(h1.getAttribute('w:hRule')).toBe('atLeast');
    expect(rows[2].getElementsByTagName('w:trHeight').length).toBe(0);

    // Schema position: w:trPr is the first ELEMENT child of w:tr.
    for (const r of [0, 1]) {
      expect(firstElementChild(rows[r])?.tagName).toBe('w:trPr');
    }
    // Header row trPr keeps schema order: cantSplit < trHeight < tblHeader.
    const kids: string[] = [];
    const trPr = rows[0].getElementsByTagName('w:trPr')[0] as Element;
    for (let i = 0; i < trPr.childNodes.length; i++) {
      const c = trPr.childNodes[i] as Element;
      if (c.nodeType === 1) kids.push(c.tagName);
    }
    expect(kids).toEqual(['w:cantSplit', 'w:trHeight', 'w:tblHeader']);
  });

  it('auto rule round-trips in saved xml', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 3, height: 200, rule: 'auto' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const h = directTrs(tbl)[2].getElementsByTagName('w:trHeight')[0] as Element;
    expect(h.getAttribute('w:val')).toBe('200');
    expect(h.getAttribute('w:hRule')).toBe('auto');
  });

  it('bad rule / bad height are loud E_VALIDATION', async () => {
    const { ast, tableId } = await freshTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 1, height: 400, rule: 'huge' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*rule/);
    for (const height of [0, -5, 12.5, 'tall', undefined, NaN]) {
      await expect(executeTable({
        action: 'table', target: tableId,
        params: { op: 'setRowHeight', rowIndex: 1, height, rule: 'exact' },
      } as any, ast)).rejects.toThrow(/E_VALIDATION.*height/);
    }
    // Missing height is loud (no silent default).
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 1 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*height/);
  });

  it('combined shading + exact height doc passes the real OOXML validator', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 2, colIndex: 2, fill: 'D9E2F3' },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 1, height: 600, rule: 'exact' },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 3, height: 300, rule: 'atLeast' },
    } as any, ast);
    const res = await validateDocxBuffer(parser.rezip(ast));
    expect(res.ok).toBe(true);
  }, 120000);
});

describe('E3 bounds + integer guards (loud E_VALIDATION)', () => {
  it('out-of-bounds row/col are loud for both ops', async () => {
    const { ast, tableId } = await freshTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 100, colIndex: 1, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*out of bounds/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, colIndex: 100, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*out of bounds/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 100, height: 400 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*out of bounds/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: -1, height: 400 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*out of bounds/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: -1, colIndex: 1, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*out of bounds/);
  });

  it('non-integer indexes are loud (mergeCells integer-guard precedent)', async () => {
    const { ast, tableId } = await freshTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1.5, colIndex: 1, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*integer/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 0.5, height: 400 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*integer/);
  });
});

describe('E3 row/col aliases + required coordinates (no silent 0,0)', () => {
  it('row/col aliases shade exactly the aliased cell (nested params)', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', row: 3, col: 2, fill: 'FF0000' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect((saved.match(/w:fill="FF0000"/g) || []).length).toBe(1);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    for (let r = 0; r < rows.length; r++) {
      const cells = directTcs(rows[r]);
      for (let c = 0; c < cells.length; c++) {
        expect(cells[c].getElementsByTagName('w:shd').length, `cell (${r},${c})`).toBe(r === 2 && c === 1 ? 1 : 0);
      }
    }
  });

  it('row/col aliases work flat and mixed with rowIndex/colIndex', async () => {
    const { parser, ast, tableId } = await freshTable();
    // Flat alias shape for setRowHeight.
    await executeTable({
      action: 'table', target: tableId, operation: 'setRowHeight',
      row: 2, height: 450,
    } as any, ast);
    // Mixed: row alias + colIndex canonical.
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', row: 1, colIndex: 2, fill: '00FF00' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const h = directTrs(tbl)[1].getElementsByTagName('w:trHeight')[0] as Element;
    expect(h.getAttribute('w:val')).toBe('450');
    expect((saved.match(/w:fill="00FF00"/g) || []).length).toBe(1);
  });

  it('missing coordinates are loud E_VALIDATION (never silent 0,0)', async () => {
    const { ast, tableId } = await freshTable();
    // setCellShading: no coords anywhere.
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*(rowIndex|row)/);
    // Missing col only.
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 1, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*(colIndex|col)/);
    // Missing row only.
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', colIndex: 1, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*(rowIndex|row)/);
    // setRowHeight: no row anywhere.
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', height: 400 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*rowIndex/);
    // Flat shapes without coords are loud too.
    await expect(executeTable({
      action: 'table', target: tableId, operation: 'setCellShading', fill: 'FF0000',
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
    await expect(executeTable({
      action: 'table', target: tableId, operation: 'setRowHeight', height: 400,
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*rowIndex/);
  });

  it('alias integer guards are loud (non-integer row/col)', async () => {
    const { ast, tableId } = await freshTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', row: 1.5, col: 1, fill: 'FF0000' },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*integer/);
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', row: 0.5, height: 400 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*integer/);
  });
});

describe('E3 refreshTableNode survival (save -> reload -> read back)', () => {
  it('shading + row height survive a later op, save, and reload', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellShading', rowIndex: 2, colIndex: 1, fill: 'FF0000' },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setRowHeight', rowIndex: 2, height: 500, rule: 'exact' },
    } as any, ast);
    // Later ops rebuild via refreshTableNode — markup must survive.
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'touched' },
    } as any, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'addRow', afterIndex: 3 },
    } as any, ast);

    const buf = parser.rezip(ast);
    const saved = new AdmZip(buf).readAsText('word/document.xml');
    expect((saved.match(/w:fill="FF0000"/g) || []).length).toBe(1);
    expect(saved).toContain('w:val="500"');
    expect(saved).toContain('w:hRule="exact"');

    const reloaded: any = await new DocxParser(buf).getAST();
    const all = reloaded.sections.flatMap((s: any) => s.nodes);
    const reTbl = all.find((n: any) => n.type === 'table');
    expect(reTbl).toBeDefined();
    const xml = serializer.serializeToString(reTbl.xml);
    expect((xml.match(/w:fill="FF0000"/g) || []).length).toBe(1);
    expect(xml).toContain('w:val="500"');
    expect(xml).toContain('w:hRule="exact"');
    // Reloaded grid grew by the added row (3 -> 4 tr elements in reloaded XML).
    // (getAST table metadata carries no tableDimensions on reload —
    // pre-existing parser posture, untouched here — so count rows in XML.)
    expect(reTbl.xml.getElementsByTagName('w:tr').length).toBe(4);
  });
});

describe('E3 schema wiring (tableOp union, not a new action)', () => {
  it('strict schema accepts both sub-ops with their flat args', () => {
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setCellShading',
      rowIndex: 1, colIndex: 2, fill: 'FF0000',
    } as never).ok).toBe(true);
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setRowHeight',
      rowIndex: 1, height: 400, rule: 'exact',
    } as never).ok).toBe(true);
    expect(validateOperationStrict({
      action: 'table', target: 't_001',
      params: { op: 'setRowHeight', rowIndex: 1, height: 400 },
    } as never).ok).toBe(true);
  });

  it('strict schema accepts row/col aliases flat (no unrecognized_keys)', () => {
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setCellShading',
      row: 1, col: 2, fill: 'FF0000',
    } as never).ok).toBe(true);
    expect(validateOperationStrict({
      action: 'table', target: 't_001', operation: 'setRowHeight',
      row: 1, height: 400, rule: 'exact',
    } as never).ok).toBe(true);
  });

  it('unknown table sub-op still rejected at the gate; flat bad rule rejected', () => {
    expect(() =>
      validateOperationStrict({ action: 'table', target: 't_001', operation: 'frobCells' })
    ).toThrow(/E_VALIDATION/);
    expect(OperationSchema.safeParse({
      action: 'table', target: 't_001', operation: 'setRowHeight',
      rowIndex: 1, height: 400, rule: 'huge',
    }).success).toBe(false);
  });
});

describe('Phase 2 logical cell resolution (setCellShading with gridSpan and hMerge)', () => {
  it('setCellShading(row: 1, colIndex: 3, fill: "FF0000") shades regular_cell in merged header row', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table',
      target: tableId,
      row: 1,
      colIndex: 3,
      fill: 'FF0000',
      operation: 'setCellShading',
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    const doc = parseXml(saved);
    const tbl = doc.getElementsByTagName('w:tbl')[0];
    const rows = directTrs(tbl);
    const row0Cells = directTcs(rows[0]);
    // Row 0 has 2 physical cells: [gridSpan=2, regular_cell]
    expect(row0Cells.length).toBe(2);

    // Physical cell 0 (gridSpan=2) must not have the FF0000 shading
    const cell0Shd = row0Cells[0].getElementsByTagName('w:shd');
    expect(cell0Shd.length).toBe(0);

    // Physical cell 1 (regular_cell) must have FF0000 shading
    const cell1Shd = row0Cells[1].getElementsByTagName('w:shd');
    expect(cell1Shd.length).toBe(1);
    expect(cell1Shd[0].getAttribute('w:fill')).toBe('FF0000');
    expect(row0Cells[1].textContent).toContain('REGULAR_CELL');
  });

  it('setCellShading(row: 1, colIndex: 1, fill: "00FF00") shades gridSpan=2 cell in merged header row', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table',
      target: tableId,
      row: 1,
      colIndex: 1,
      fill: '00FF00',
      operation: 'setCellShading',
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    const doc = parseXml(saved);
    const tbl = doc.getElementsByTagName('w:tbl')[0];
    const rows = directTrs(tbl);
    const row0Cells = directTcs(rows[0]);
    expect(row0Cells.length).toBe(2);

    // Physical cell 0 (gridSpan=2) must have 00FF00 shading
    const cell0Shd = row0Cells[0].getElementsByTagName('w:shd');
    expect(cell0Shd.length).toBe(1);
    expect(cell0Shd[0].getAttribute('w:fill')).toBe('00FF00');
    expect(row0Cells[0].textContent).toContain('SPAN_CELL');

    // Physical cell 1 (regular_cell) must not have 00FF00 shading
    const cell1Shd = row0Cells[1].getElementsByTagName('w:shd');
    expect(cell1Shd.length).toBe(0);
  });

  it('setCellShading on hMerge continue cover cell throws E_VALIDATION with restart remedy', async () => {
    const { ast, tableId } = await freshTableWithMerges();
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        row: 2,
        colIndex: 2,
        fill: 'FF0000',
        operation: 'setCellShading',
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: setCellShading column index 2 targets a merged cover cell (hMerge continue) and is not directly editable; target the merge restart instead.'
    );
  });

  it('setCellShading out-of-bounds error message reports logical column count', async () => {
    const { ast, tableId } = await freshTableWithMerges();
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        row: 1,
        colIndex: 4,
        fill: 'FF0000',
        operation: 'setCellShading',
      } as any, ast)
    ).rejects.toThrow(/out of bounds \(row has 3 logical columns\)/);

    // Also assert exact remedy message shape
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        row: 1,
        colIndex: 4,
        fill: 'FF0000',
        operation: 'setCellShading',
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: setCellShading column index 4 out of bounds (row has 3 logical columns). Remedy: pass colIndex in [1, 3], counting from 1.'
    );
  });

  it('setVerticalAlign resolves logical columns and respects cover policy', async () => {
    // 1. targetRow: 0, colIndex: 3 resolves to regular_cell (assert against SAVED BYTES)
    const { parser: p1, ast: ast1, tableId: id1 } = await freshTableWithMerges();
    await executeTable({
      action: 'table',
      target: id1,
      rowIndex: 1,
      colIndex: 3,
      operation: 'setVerticalAlign',
      params: { vAlign: 'bottom' },
    } as any, ast1);
    const saved1 = savedDocumentXml(p1, ast1);
    const doc1 = parseXml(saved1);
    const tbl1 = onlyTable(doc1);
    const rows1 = directTrs(tbl1);
    const row0Cells1 = directTcs(rows1[0]);
    expect(row0Cells1[1].getElementsByTagName('w:vAlign')[0]?.getAttribute('w:val')).toBe('bottom');
    expect(row0Cells1[0].getElementsByTagName('w:vAlign').length).toBe(0);

    // 2. targetRow: 1, colIndex: 2 (cover cell) throws Cover Policy A error
    await expect(
      executeTable({
        action: 'table',
        target: id1,
        rowIndex: 2,
        colIndex: 2,
        operation: 'setVerticalAlign',
        params: { vAlign: 'top' },
      } as any, ast1)
    ).rejects.toThrow(
      'E_VALIDATION: setVerticalAlign column index 2 targets a merged cover cell (hMerge continue) and is not directly editable; target the merge restart instead.'
    );

    // 3. Column-wide across rows (only colIndex: 2 specified): skips cover cell in row 1, applies to row 0 span cell
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table',
      target: tableId,
      colIndex: 2,
      operation: 'setVerticalAlign',
      params: { vAlign: 'center' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    const row0Cells = directTcs(rows[0]);
    const row1Cells = directTcs(rows[1]);

    // Assert row 0, cell 0 (the span covering col 0-1) has w:vAlign="center"
    expect(row0Cells[0].getElementsByTagName('w:vAlign')[0]?.getAttribute('w:val')).toBe('center');
    // Assert row 0, cell 1 (the tail regular cell) has NO w:vAlign
    expect(row0Cells[1].getElementsByTagName('w:vAlign').length).toBe(0);
    // Assert row 1, cell 0 (the restart cell) has NO w:vAlign
    expect(row1Cells[0].getElementsByTagName('w:vAlign').length).toBe(0);
    // Assert row 1, cell 1 (the continue cover) has NO w:vAlign
    expect(row1Cells[1].getElementsByTagName('w:vAlign').length).toBe(0);
  });

  it('setCellShading interior-span shading shades gridSpan=2 cell in saved bytes', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table',
      target: tableId,
      row: 1,
      colIndex: 2,
      fill: '112233',
      operation: 'setCellShading',
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    const doc = parseXml(saved);
    const tbl = onlyTable(doc);
    const rows = directTrs(tbl);
    const row0Cells = directTcs(rows[0]);
    expect(row0Cells.length).toBe(2);

    // Assert that cells[0] receives w:fill="112233" in savedDocumentXml
    const cell0Shd = row0Cells[0].getElementsByTagName('w:shd');
    expect(cell0Shd.length).toBe(1);
    expect(cell0Shd[0].getAttribute('w:fill')).toBe('112233');

    // Assert that cells[1] (regular cell) has NO w:shd
    const cell1Shd = row0Cells[1].getElementsByTagName('w:shd');
    expect(cell1Shd.length).toBe(0);
  });

  it('setCellShading positive shading on hMerge="restart" cell in saved bytes', async () => {
    const { parser, ast, tableId } = await freshTableWithMerges();
    await executeTable({
      action: 'table',
      target: tableId,
      row: 2,
      colIndex: 1,
      fill: '445566',
      operation: 'setCellShading',
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    const doc = parseXml(saved);
    const tbl = onlyTable(doc);
    const rows = directTrs(tbl);
    const row1Cells = directTcs(rows[1]);

    // Assert that row 1's restart cell receives w:fill="445566" in savedDocumentXml
    const cell0Shd = row1Cells[0].getElementsByTagName('w:shd');
    expect(cell0Shd.length).toBe(1);
    expect(cell0Shd[0].getAttribute('w:fill')).toBe('445566');

    // Assert that sibling cells in row 1 (cell 1 cover, cell 2 regular) have NO w:shd
    expect(row1Cells[1].getElementsByTagName('w:shd').length).toBe(0);
    expect(row1Cells[2].getElementsByTagName('w:shd').length).toBe(0);

    // Table-wide isolation: exactly 1 w:shd in the whole table
    const allShds = tbl.getElementsByTagName('w:shd');
    expect(allShds.length).toBe(1);
  });

  it('nested table isolation: resolves strictly against outer row direct w:tc children without counting or mutating nested cells', async () => {
    const { parser, ast, tableId } = await freshTable();
    const tblNode = ast.sections[0].nodes.find((n: any) => n.id === tableId);
    const tblEl = tblNode.xml as Element;
    const outerRows = directTrs(tblEl);
    const outerRow0Cells = directTcs(outerRows[0]);
    const hostCell = outerRow0Cells[0]; // outer cell (0, 0)
    const doc = hostCell.ownerDocument;

    // Construct and append a nested <w:tbl> with multiple <w:tc> inside outer cell (0, 0)
    const nestedTbl = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:tbl');
    for (let r = 0; r < 2; r++) {
      const nestedTr = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:tr');
      for (let c = 0; c < 2; c++) {
        const nestedTc = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:tc');
        const p = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:p');
        const run = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:r');
        const t = doc.createElementNS('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'w:t');
        t.textContent = `nested_${r}_${c}`;
        run.appendChild(t);
        p.appendChild(run);
        nestedTc.appendChild(p);
        nestedTr.appendChild(nestedTc);
      }
      nestedTbl.appendChild(nestedTr);
    }
    hostCell.appendChild(nestedTbl);

    // Verify outer host cell contains nested <w:tc> descendants
    expect(hostCell.getElementsByTagName('w:tc').length).toBe(4);

    // Assert findLogicalTableCell and resolveTableCell strictly resolve against outer row's direct <w:tc>
    const logCol0 = findLogicalTableCell(outerRows[0], 0);
    expect(logCol0.tc).toBe(hostCell);
    expect(logCol0.logicalWidth).toBe(2);

    const logCol1 = findLogicalTableCell(outerRows[0], 1);
    expect(logCol1.tc).toBe(outerRow0Cells[1]);
    expect(logCol1.logicalWidth).toBe(2);

    const logCol2 = findLogicalTableCell(outerRows[0], 2);
    expect(logCol2.tc).toBeUndefined();
    expect(logCol2.logicalWidth).toBe(2);

    expect(resolveTableCell(outerRows[0], 0, 'test')).toBe(hostCell);
    expect(resolveTableCell(outerRows[0], 1, 'test')).toBe(outerRow0Cells[1]);
    expect(() => resolveTableCell(outerRows[0], 2, 'test')).toThrow(
      'E_VALIDATION: test column index 3 out of bounds (row has 2 logical columns). Remedy: pass colIndex in [1, 2], counting from 1.'
    );

    // Run setCellShading on outer row 0, colIndex 1
    await executeTable({
      action: 'table',
      target: tableId,
      row: 1,
      colIndex: 2,
      fill: 'AABBCC',
      operation: 'setCellShading',
    } as any, ast);

    const saved = savedDocumentXml(parser, ast);
    const savedDoc = parseXml(saved);
    const savedOuterTbl = savedDoc.getElementsByTagName('w:tbl')[0];
    const savedOuterRows = directTrs(savedOuterTbl);
    const savedOuterRow0Cells = directTcs(savedOuterRows[0]);

    // Outer row has 2 direct w:tc cells
    expect(savedOuterRow0Cells.length).toBe(2);

    // Outer cell 1 received the shading
    const outerCell1Shd = savedOuterRow0Cells[1].getElementsByTagName('w:shd');
    expect(outerCell1Shd.length).toBe(1);
    expect(outerCell1Shd[0].getAttribute('w:fill')).toBe('AABBCC');

    // Outer cell 0's direct w:tc has NO w:shd
    const outerCell0Pr = directChildren(savedOuterRow0Cells[0], 'w:tcPr')[0];
    expect(outerCell0Pr ? directChildren(outerCell0Pr, 'w:shd').length : 0).toBe(0);

    // Crucially: NONE of the nested cells received w:shd (not mutated)
    const savedNestedTbl = savedOuterRow0Cells[0].getElementsByTagName('w:tbl')[0];
    expect(savedNestedTbl).toBeDefined();
    const nestedShds = savedNestedTbl.getElementsByTagName('w:shd');
    expect(nestedShds.length).toBe(0);
  });
});

describe('Phase 2 strict schema validation: width and height positive integer enforcement', () => {
  it('validateOperationStrict rejects flat negative, zero, and float width', () => {
    for (const badWidth of [-100, -1, 0, 100.5, 0.1]) {
      expect(() =>
        validateOperationStrict({
          action: 'table',
          target: 't_001',
          operation: 'setColumnWidth',
          colIndex: 1,
          width: badWidth,
        })
      ).toThrow(/E_VALIDATION/);

      expect(
        OperationSchema.safeParse({
          action: 'table',
          target: 't_001',
          operation: 'setColumnWidth',
          colIndex: 1,
          width: badWidth,
        }).success
      ).toBe(false);
    }
  });

  it('validateOperationStrict rejects flat negative, zero, and float height', () => {
    for (const badHeight of [-200, -1, 0, 400.5, 0.5]) {
      expect(() =>
        validateOperationStrict({
          action: 'table',
          target: 't_001',
          operation: 'setRowHeight',
          rowIndex: 1,
          height: badHeight,
        })
      ).toThrow(/E_VALIDATION/);

      expect(
        OperationSchema.safeParse({
          action: 'table',
          target: 't_001',
          operation: 'setRowHeight',
          rowIndex: 1,
          height: badHeight,
        }).success
      ).toBe(false);
    }
  });

  it('validateOperationStrict accepts positive integer width and height', () => {
    expect(
      validateOperationStrict({
        action: 'table',
        target: 't_001',
        operation: 'setColumnWidth',
        colIndex: 1,
        width: 2000,
      }).ok
    ).toBe(true);

    expect(
      validateOperationStrict({
        action: 'table',
        target: 't_001',
        operation: 'setRowHeight',
        rowIndex: 1,
        height: 400,
      }).ok
    ).toBe(true);
  });
});

describe('table operation error harmonization (deleteRow and deleteColumn E_VALIDATION)', () => {
  it('deleteRow emits E_VALIDATION with accurate row count and remedy on out-of-bounds index', async () => {
    const { ast, tableId } = await freshTable();
    // Index >= rows.length (table has 3 rows: indices 0, 1, 2)
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteRow', index: 4 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteRow row index 4 out of bounds (table has 3 rows). Remedy: pass index in [1, 3], counting from 1.'
    );

    // Flat form with index 99
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        operation: 'deleteRow',
        index: 100,
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteRow row index 100 out of bounds (table has 3 rows). Remedy: pass index in [1, 3], counting from 1.'
    );
  });

  it('deleteRow emits E_VALIDATION on negative index', async () => {
    const { ast, tableId } = await freshTable();
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteRow', index: -1 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteRow row index -1 out of bounds (table has 3 rows). Remedy: pass index in [1, 3], counting from 1.'
    );
  });

  it('deleteRow emits E_VALIDATION on non-integer index', async () => {
    const { ast, tableId } = await freshTable();
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteRow', index: 1.5 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteRow row index 1.5 out of bounds (table has 3 rows). Remedy: pass index in [1, 3], counting from 1.'
    );
  });

  it('deleteColumn emits E_VALIDATION with accurate logical column count and remedy on simple table', async () => {
    const { ast, tableId } = await freshTable(); // 2 logical columns (0, 1)
    // Index >= logical columns
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteColumn', index: 3 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteColumn column index 3 out of bounds (row has 2 logical columns). Remedy: pass index in [1, 2], counting from 1.'
    );

    // Flat form with colIndex alias
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        operation: 'deleteColumn',
        colIndex: 6,
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteColumn column index 6 out of bounds (row has 2 logical columns). Remedy: pass index in [1, 2], counting from 1.'
    );

    // Negative index
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteColumn', index: -1 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteColumn column index -1 out of bounds (row has 2 logical columns). Remedy: pass index in [1, 2], counting from 1.'
    );
  });

  it('deleteColumn emits E_VALIDATION with accurate logical column count and remedy on merged table', async () => {
    const { ast, tableId } = await freshTableWithMerges(); // 3 logical columns (0, 1, 2)
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteColumn', index: 4 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteColumn column index 4 out of bounds (row has 3 logical columns). Remedy: pass index in [1, 3], counting from 1.'
    );

    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        operation: 'deleteColumn',
        index: 11,
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteColumn column index 11 out of bounds (row has 3 logical columns). Remedy: pass index in [1, 3], counting from 1.'
    );
  });

  it('deleteColumn emits E_VALIDATION on non-integer index', async () => {
    const { ast, tableId } = await freshTable();
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        params: { op: 'deleteColumn', index: 1.5 },
      } as any, ast)
    ).rejects.toThrow(
      'E_VALIDATION: deleteColumn column index 1.5 out of bounds. Remedy: pass integer index.'
    );
  });
});

describe('Header text color + keep-together (saved bytes)', () => {
  it('two-step header recipe: dark headerFill then white row-0 text', async () => {
    const { parser, ast, tableId } = await freshTable();
    const shade = await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setShading', fill: 'F7F9FC', headerFill: '1F3A5F' },
    } as any, ast);
    // Dark header against default dark text warns (advisory, still applies).
    expect(shade.warnings.length).toBeGreaterThan(0);
    const colored = await executeTable({
      action: 'table', target: tableId, operation: 'setCellTextColor',
      rowIndex: 1, color: 'FFFFFF',
    } as any, ast);
    expect(colored.warnings).toEqual([]);

    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(3);
    // Header row runs carry white; data rows keep default (no w:color).
    for (const tc of directTcs(rows[0])) {
      const colors = tc.getElementsByTagName('w:color');
      expect(colors.length).toBeGreaterThan(0);
      for (let i = 0; i < colors.length; i++) {
        expect((colors[i] as Element).getAttribute('w:val')).toBe('FFFFFF');
      }
    }
    for (let r = 1; r < rows.length; r++) {
      for (const tc of directTcs(rows[r])) {
        expect(tc.getElementsByTagName('w:color').length).toBe(0);
      }
    }
  });

  it('setCellTextColor targets one cell; bad colors fail loud', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellTextColor', rowIndex: 3, colIndex: 2, color: '#FF0000' },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    expect((saved.match(/w:val="FF0000"/g) || []).length).toBe(1);

    for (const color of ['FFF', 'GGGGGG', 'red', '', '#FFF']) {
      await expect(executeTable({
        action: 'table', target: tableId,
        params: { op: 'setCellTextColor', rowIndex: 1, colIndex: 1, color },
      } as any, ast)).rejects.toThrow(/E_VALIDATION.*color/);
    }
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCellTextColor', rowIndex: 1, colIndex: 1 },
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*color/);
  });

  it('setKeepTogether chains non-last rows with keepNext; last row untouched', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table', target: tableId, operation: 'setKeepTogether',
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(3);
    // Every row cantSplit (no row straddles a page break).
    for (const tr of rows) {
      expect(tr.getElementsByTagName('w:cantSplit').length).toBe(1);
    }
    // Non-last rows glue forward; the last row flows normally.
    for (let r = 0; r < rows.length - 1; r++) {
      for (const tc of directTcs(rows[r])) {
        for (const p of Array.from(tc.getElementsByTagName('w:p'))) {
          const pPr = (p as Element).getElementsByTagName('w:pPr')[0];
          expect(pPr, `row ${r} cell keeps next`).toBeDefined();
          expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
        }
      }
    }
    for (const tc of directTcs(rows[rows.length - 1])) {
      for (const p of Array.from(tc.getElementsByTagName('w:p'))) {
        const pPr = (p as Element).getElementsByTagName('w:pPr')[0];
        expect(pPr ? pPr.getElementsByTagName('w:keepNext').length : 0).toBe(0);
      }
    }
    const validation = await validateDocxBuffer(parser.rezip(ast));
    expect(validation.ok).toBe(true);
    expect(validation.errors).toEqual([]);
  });
});

describe('deleteRow and deleteColumn index handling', () => {
  it('flat rowIndex: 3 removes row 3/4 and preserves header A/B', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table',
      target: tableId,
      operation: 'deleteRow',
      rowIndex: 3,
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(2);
    const row0Cells = directTcs(rows[0]).map((tc) => tc.textContent);
    const row1Cells = directTcs(rows[1]).map((tc) => tc.textContent);
    expect(row0Cells).toEqual(['A', 'B']);
    expect(row1Cells).toEqual(['1', '2']);
  });

  it('flat row: 2 removes row 1/2', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table',
      target: tableId,
      operation: 'deleteRow',
      row: 2,
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(2);
    const row0Cells = directTcs(rows[0]).map((tc) => tc.textContent);
    const row1Cells = directTcs(rows[1]).map((tc) => tc.textContent);
    expect(row0Cells).toEqual(['A', 'B']);
    expect(row1Cells).toEqual(['3', '4']);
  });

  it('nested params: { rowIndex: 2 } removes row 1/2', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table',
      target: tableId,
      params: { op: 'deleteRow', rowIndex: 2 },
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(2);
    const row0Cells = directTcs(rows[0]).map((tc) => tc.textContent);
    const row1Cells = directTcs(rows[1]).map((tc) => tc.textContent);
    expect(row0Cells).toEqual(['A', 'B']);
    expect(row1Cells).toEqual(['3', '4']);
  });

  it('index: 2 wins over rowIndex: 3 to remove row 1/2', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table',
      target: tableId,
      operation: 'deleteRow',
      index: 2,
      rowIndex: 3,
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(2);
    const row0Cells = directTcs(rows[0]).map((tc) => tc.textContent);
    const row1Cells = directTcs(rows[1]).map((tc) => tc.textContent);
    expect(row0Cells).toEqual(['A', 'B']);
    expect(row1Cells).toEqual(['3', '4']);
  });

  it('no index given on deleteRow throws E_VALIDATION and leaves document bytes unchanged', async () => {
    const { parser, ast, tableId } = await freshTable();
    const xmlBefore = savedDocumentXml(parser, ast);
    const bytesBefore = parser.rezip(ast);
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        operation: 'deleteRow',
      } as any, ast)
    ).rejects.toThrow(/E_VALIDATION: deleteRow requires a row index/);
    const xmlAfter = savedDocumentXml(parser, ast);
    expect(xmlAfter).toBe(xmlBefore);
    const bytesAfter = parser.rezip(ast);
    expect(bytesAfter.equals(bytesBefore)).toBe(true);
  });

  it('deleteColumn with colIndex: 2 removes the second column', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({
      action: 'table',
      target: tableId,
      operation: 'deleteColumn',
      colIndex: 2,
    } as any, ast);
    const saved = savedDocumentXml(parser, ast);
    const tbl = onlyTable(parseXml(saved));
    const rows = directTrs(tbl);
    expect(rows.length).toBe(3);
    for (const r of rows) {
      const cells = directTcs(r);
      expect(cells.length).toBe(1);
    }
    expect(directTcs(rows[0])[0].textContent).toBe('A');
    expect(directTcs(rows[1])[0].textContent).toBe('1');
    expect(directTcs(rows[2])[0].textContent).toBe('3');
  });

  it('no index given on deleteColumn throws E_VALIDATION and leaves document unchanged', async () => {
    const { parser, ast, tableId } = await freshTable();
    const xmlBefore = savedDocumentXml(parser, ast);
    const bytesBefore = parser.rezip(ast);
    await expect(
      executeTable({
        action: 'table',
        target: tableId,
        operation: 'deleteColumn',
      } as any, ast)
    ).rejects.toThrow(/E_VALIDATION: deleteColumn requires a column index/);
    const xmlAfter = savedDocumentXml(parser, ast);
    expect(xmlAfter).toBe(xmlBefore);
    const bytesAfter = parser.rezip(ast);
    expect(bytesAfter.equals(bytesBefore)).toBe(true);
  });

  // Number('') and Number([]) and Number(false) are 0, Number(true) is 1: a malformed index must not become a real one.
  it.each([
    ['deleteRow', { rowIndex: '' }],
    ['deleteRow', { rowIndex: '   ' }],
    ['deleteRow', { rowIndex: true }],
    ['deleteRow', { rowIndex: false }],
    ['deleteRow', { rowIndex: [] }],
    ['deleteRow', { index: '' }],
    ['deleteColumn', { index: '' }],
    ['deleteColumn', { index: true }],
    ['deleteColumn', { index: [] }],
    ['deleteColumn', { colIndex: '' }],
  ] as Array<[string, Record<string, unknown>]>)('%s with %j gives E_VALIDATION and leaves the document unchanged', async (operation, extra) => {
    const { parser, ast, tableId } = await freshTable();
    const bytesBefore = parser.rezip(ast);
    await expect(
      executeTable({ action: 'table', target: tableId, operation, ...extra } as any, ast)
    ).rejects.toThrow(/E_VALIDATION/);
    expect(parser.rezip(ast).equals(bytesBefore)).toBe(true);
  });

  it('a numeric string still works as an index', async () => {
    const { parser, ast, tableId } = await freshTable();
    await executeTable({ action: 'table', target: tableId, operation: 'deleteRow', rowIndex: '2' } as any, ast);
    const rows = directTrs(onlyTable(parseXml(savedDocumentXml(parser, ast))));
    expect(rows.map((r) => directTcs(r).map((tc) => tc.textContent))).toEqual([['A', 'B'], ['3', '4']]);
  });
});


