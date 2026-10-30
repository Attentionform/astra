import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocumentAST,
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executeTable,
  TableOperation,
  validateOperationStrict,
} from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

function createMockTableAST(): DocumentAST {
  const xmlStr = `
    <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:tblPr/>
      <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Header 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Header 2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cell 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Cell 2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  `;
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  const tblEl = doc.documentElement;

  return {
    metadata: {},
    sections: [{
      id: 'sec_root',
      heading: null,
      nodes: [{
        id: 't_001',
        type: 'table',
        xml: tblEl,
        metadata: {
          text: 'Header 1 Header 2 Cell 1 Cell 2',
          tableDimensions: { rows: 2, cols: 2 }
        } as any,
        version: 1,
        checksum: 'chk_tbl'
      }],
      pageSetup: {},
      version: 1,
      checksum: 'sec_chk_1'
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1
  };
}

describe('Advanced Table Operations', () => {
  it('should support setColumnWidths and setColumnWidth', async () => {
    const ast = createMockTableAST();
    const op: TableOperation = {
      action: 'table',
      target: 't_001',
      params: {
        op: 'setColumnWidths',
        widths: [2000, 4500]
      }
    };

    await executeTable(op, ast);
    const tblEl = ast.sections[0].nodes[0].xml;
    const gridCols = tblEl.getElementsByTagName('w:gridCol');
    expect(gridCols[0].getAttribute('w:w')).toBe('2000');
    expect(gridCols[1].getAttribute('w:w')).toBe('4500');

    // Single column width
    await executeTable({
      action: 'table',
      target: 't_001',
      params: {
        op: 'setColumnWidth',
        colIndex: 1,
        width: 2500
      }
    }, ast);

    expect(gridCols[0].getAttribute('w:w')).toBe('2500');
  });

  it('should support addColumn and deleteColumn', async () => {
    const ast = createMockTableAST();

    // Add Column
    await executeTable({
      action: 'table',
      target: 't_001',
      params: {
        op: 'addColumn',
        afterIndex: 2,
        width: 3200
      }
    }, ast);

    const tblEl = ast.sections[0].nodes[0].xml;
    expect(tblEl.getElementsByTagName('w:gridCol').length).toBe(3);
    const rows = tblEl.getElementsByTagName('w:tr');
    expect(rows[0].getElementsByTagName('w:tc').length).toBe(3);

    // Delete Column
    await executeTable({
      action: 'table',
      target: 't_001',
      params: {
        op: 'deleteColumn',
        index: 1
      }
    }, ast);

    expect(tblEl.getElementsByTagName('w:gridCol').length).toBe(2);
    expect(rows[0].getElementsByTagName('w:tc').length).toBe(2);
  });

  it('should support setBorder and mergeCells', async () => {
    const ast = createMockTableAST();

    // Set Border
    await executeTable({
      action: 'table',
      target: 't_001',
      params: {
        op: 'setBorder',
        border: 'all',
        color: '4338CA',
        sz: 8
      }
    }, ast);

    const tblEl = ast.sections[0].nodes[0].xml;
    const tblBorders = tblEl.getElementsByTagName('w:tblBorders')[0];
    expect(tblBorders).toBeDefined();
    const topBorder = tblBorders.getElementsByTagName('w:top')[0];
    expect(topBorder.getAttribute('w:color')).toBe('4338CA');

    // Merge Cells
    await executeTable({
      action: 'table',
      target: 't_001',
      params: {
        op: 'mergeCells',
        fromRow: 1,
        fromCol: 1,
        toRow: 1,
        toCol: 2
      }
    }, ast);

    const firstCell = tblEl.getElementsByTagName('w:tc')[0];
    const hMerge = firstCell.getElementsByTagName('w:hMerge')[0];
    expect(hMerge.getAttribute('w:val')).toBe('restart');
  });
});

function createMock3ColTableAST(): DocumentAST {
  const xmlStr = `
    <w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:tblPr/>
      <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>R0C0</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R0C1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R0C2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>R1C0</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R1C1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>R1C2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  `;
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  const tblEl = doc.documentElement;

  return {
    metadata: {},
    sections: [{
      id: 'sec_root',
      heading: null,
      nodes: [{
        id: 't_001',
        type: 'table',
        xml: tblEl,
        metadata: {
          text: 'R0C0 R0C1 R0C2 R1C0 R1C1 R1C2',
          tableDimensions: { rows: 2, cols: 3 }
        } as any,
        version: 1,
        checksum: 'chk_tbl'
      }],
      pageSetup: {},
      version: 1,
      checksum: 'sec_chk_1'
    }],
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    settings: { trackChanges: false, defaultAuthor: 'ASTra' },
    nextChangeId: 1
  };
}

function cellText(ast: DocumentAST, row: number, col: number): string {
  const tblEl = ast.sections[0].nodes[0].xml as any;
  const trs = tblEl.getElementsByTagName('w:tr');
  const tcs = (trs[row] as any).getElementsByTagName('w:tc');
  const ts = (tcs[col] as any).getElementsByTagName('w:t');
  let out = '';
  for (let i = 0; i < ts.length; i++) out += ts[i].textContent || '';
  return out;
}

describe('updateCell columnIndex alias', () => {
  it('writes column 2 when columnIndex:3 is nested in params', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 2, columnIndex: 3, text: 'NESTED' }
    } as any, ast);
    expect(cellText(ast, 1, 2)).toBe('NESTED');
    expect(cellText(ast, 1, 0)).toBe('R1C0');
    expect(cellText(ast, 1, 1)).toBe('R1C1');
  });

  it('writes column 2 when columnIndex:3 is flat on the op', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      operation: 'updateCell',
      rowIndex: 1,
      columnIndex: 3,
      text: 'FLAT'
    } as any, ast);
    expect(cellText(ast, 0, 2)).toBe('FLAT');
    expect(cellText(ast, 0, 0)).toBe('R0C0');
  });

  it('colIndex:3 still writes column 2', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, colIndex: 3, text: 'LEGACY' }
    }, ast);
    expect(cellText(ast, 0, 2)).toBe('LEGACY');
    expect(cellText(ast, 0, 0)).toBe('R0C0');
  });

  it('colIndex wins when both colIndex and columnIndex are present', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, colIndex: 2, columnIndex: 3, text: 'WIN' }
    } as any, ast);
    expect(cellText(ast, 0, 1)).toBe('WIN');
    expect(cellText(ast, 0, 2)).toBe('R0C2');
  });

  it('rejects negative columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, columnIndex: -1, text: 'BAD' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('rejects non-integer columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, columnIndex: 1.5, text: 'BAD' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('rejects out-of-bounds columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, columnIndex: 10, text: 'BAD' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });
});

describe('columnIndex alias across cell/column paths', () => {
  it('splitCell accepts columnIndex and strips the merge', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'mergeCells', fromRow: 1, fromCol: 1, toRow: 1, toCol: 2 }
    }, ast);
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', rowIndex: 1, columnIndex: 2 }
    } as any, ast);
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', rowIndex: 1, columnIndex: 1 }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const row0tcs = (trs[0] as any).getElementsByTagName('w:tc');
    let hMerges = 0;
    for (let i = 0; i < row0tcs.length; i++) {
      hMerges += (row0tcs[i] as any).getElementsByTagName('w:hMerge').length;
    }
    expect(hMerges).toBe(0);
  });

  it('deleteColumn accepts columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', columnIndex: 1 }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    expect(tblEl.getElementsByTagName('w:gridCol').length).toBe(2);
    expect(cellText(ast, 0, 0)).toBe('R0C1');
    expect(cellText(ast, 0, 1)).toBe('R0C2');
  });

  it('setCellShading accepts columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellShading', rowIndex: 1, columnIndex: 3, fill: 'FF0000' }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[0] as any).getElementsByTagName('w:tc');
    expect((tcs[0] as any).getElementsByTagName('w:shd').length).toBe(0);
    expect((tcs[1] as any).getElementsByTagName('w:shd').length).toBe(0);
    const shd = (tcs[2] as any).getElementsByTagName('w:shd')[0];
    expect(shd.getAttribute('w:fill')).toBe('FF0000');
  });

  it('setCellAlignment accepts columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellAlignment', rowIndex: 1, columnIndex: 3, align: 'center' }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[0] as any).getElementsByTagName('w:tc');
    const jc = (tcs[2] as any).getElementsByTagName('w:jc')[0];
    expect(jc.getAttribute('w:val')).toBe('center');
    expect((tcs[0] as any).getElementsByTagName('w:jc').length).toBe(0);
  });

  it('setCellPadding accepts columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellPadding', rowIndex: 2, columnIndex: 3, padding: 120 }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[1] as any).getElementsByTagName('w:tc');
    const mar = (tcs[2] as any).getElementsByTagName('w:tcMar')[0];
    expect(mar).toBeDefined();
    expect(mar.getElementsByTagName('w:top')[0].getAttribute('w:w')).toBe('120');
    expect((tcs[0] as any).getElementsByTagName('w:tcMar').length).toBe(0);
  });

  it('setVerticalAlign accepts columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setVerticalAlign', rowIndex: 2, columnIndex: 3, vAlign: 'bottom' }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[1] as any).getElementsByTagName('w:tc');
    expect((tcs[2] as any).getElementsByTagName('w:vAlign')[0].getAttribute('w:val')).toBe('bottom');
    expect((tcs[0] as any).getElementsByTagName('w:vAlign').length).toBe(0);
  });
});

describe('splitCell columnIndex validation', () => {
  it('rejects negative columnIndex with colIndex vocabulary', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', rowIndex: 1, columnIndex: -1 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*colIndex/);
  });

  it('rejects non-integer columnIndex with colIndex vocabulary', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', rowIndex: 1, columnIndex: 1.5 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*colIndex/);
  });

  it('rejects out-of-bounds columnIndex with colIndex vocabulary', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', rowIndex: 1, columnIndex: 10 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*colIndex/);
  });
});

describe('deleteColumn columnIndex bounds', () => {
  it('rejects negative columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', columnIndex: -1 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('rejects out-of-bounds columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', columnIndex: 10 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('rejects non-integer columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', columnIndex: 1.5 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });
});

describe('row/col short aliases', () => {
  it('updateCell {row:2,col:2} writes the second row and column and leaves the first alone', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', row: 2, col: 2, text: 'HIT' }
    } as any, ast);
    expect(cellText(ast, 1, 1)).toBe('HIT');
    expect(cellText(ast, 0, 0)).toBe('R0C0');
    expect(cellText(ast, 1, 0)).toBe('R1C0');
    expect(cellText(ast, 0, 1)).toBe('R0C1');
  });

  it('updateCell flat row/col on the op writes the second row and column', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      operation: 'updateCell',
      row: 2,
      col: 3,
      text: 'FLAT'
    } as any, ast);
    expect(cellText(ast, 1, 2)).toBe('FLAT');
    expect(cellText(ast, 0, 0)).toBe('R0C0');
  });

  it('updateCell {rowIndex:1,colIndex:1} still writes the first row and column when aliases disagree', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'ZERO' }
    }, ast);
    expect(cellText(ast, 0, 0)).toBe('ZERO');
    expect(cellText(ast, 1, 1)).toBe('R1C1');
  });

  it('precedence: rowIndex wins over row, colIndex wins over col and columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, row: 2, colIndex: 1, col: 2, columnIndex: 3, text: 'P' }
    } as any, ast);
    expect(cellText(ast, 0, 0)).toBe('P');
    expect(cellText(ast, 1, 0)).toBe('R1C0');
    expect(cellText(ast, 0, 1)).toBe('R0C1');
    expect(cellText(ast, 0, 2)).toBe('R0C2');
  });

  it('precedence: col wins over columnIndex', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', rowIndex: 1, col: 2, columnIndex: 3, text: 'Q' }
    } as any, ast);
    expect(cellText(ast, 0, 1)).toBe('Q');
    expect(cellText(ast, 0, 2)).toBe('R0C2');
  });

  it('precedence: nested params win over flat within the same name', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      operation: 'updateCell',
      rowIndex: 1,
      colIndex: 1,
      text: 'FLAT',
      params: { op: 'updateCell', rowIndex: 2, colIndex: 3, text: 'NEST' }
    } as any, ast);
    expect(cellText(ast, 1, 2)).toBe('NEST');
    expect(cellText(ast, 0, 0)).toBe('R0C0');
  });

  it.each([
    [{ row: -1, col: 1 }],
    [{ row: 1, col: -1 }],
    [{ row: 1.5, col: 1 }],
    [{ row: 1, col: 1.5 }],
    [{ row: 'oops', col: 1 }],
    [{ row: 1, col: 'oops' }],
  ])('updateCell rejects bad row/col %o with E_VALIDATION', async (coords) => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'updateCell', ...(coords as object), text: 'BAD' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });
});

describe('row/col aliases across cell paths', () => {
  it('splitCell accepts row/col and strips the merge', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'mergeCells', fromRow: 1, fromCol: 1, toRow: 1, toCol: 2 }
    }, ast);
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', row: 1, col: 2 }
    } as any, ast);
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', row: 1, col: 1 }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const row0tcs = (trs[0] as any).getElementsByTagName('w:tc');
    let hMerges = 0;
    for (let i = 0; i < row0tcs.length; i++) {
      hMerges += (row0tcs[i] as any).getElementsByTagName('w:hMerge').length;
    }
    expect(hMerges).toBe(0);
  });

  it('splitCell rejects bad row/col', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', row: -1, col: 1 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'splitCell', row: 1, col: 1.5 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('deleteColumn accepts col', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', col: 1 }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    expect(tblEl.getElementsByTagName('w:gridCol').length).toBe(2);
    expect(cellText(ast, 0, 0)).toBe('R0C1');
    expect(cellText(ast, 0, 1)).toBe('R0C2');
  });

  it('deleteColumn rejects bad col', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', col: -1 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'deleteColumn', col: 1.5 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('setCellShading accepts row/col', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellShading', row: 1, col: 3, fill: 'FF0000' }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[0] as any).getElementsByTagName('w:tc');
    expect((tcs[0] as any).getElementsByTagName('w:shd').length).toBe(0);
    expect((tcs[2] as any).getElementsByTagName('w:shd')[0].getAttribute('w:fill')).toBe('FF0000');
  });

  it('setCellShading rejects bad row/col', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellShading', row: 1, col: -1, fill: 'FF0000' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('setCellAlignment accepts row/col', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellAlignment', row: 1, col: 3, align: 'center' }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[0] as any).getElementsByTagName('w:tc');
    expect((tcs[2] as any).getElementsByTagName('w:jc')[0].getAttribute('w:val')).toBe('center');
    expect((tcs[0] as any).getElementsByTagName('w:jc').length).toBe(0);
  });

  it('setCellAlignment rejects bad row/col', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellAlignment', row: 1, col: 1.5, align: 'center' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('setCellPadding accepts row/col', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellPadding', row: 2, col: 3, padding: 120 }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[1] as any).getElementsByTagName('w:tc');
    expect((tcs[2] as any).getElementsByTagName('w:tcMar')[0].getElementsByTagName('w:top')[0].getAttribute('w:w')).toBe('120');
    expect((tcs[0] as any).getElementsByTagName('w:tcMar').length).toBe(0);
  });

  it('setCellPadding rejects bad row/col', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setCellPadding', row: 'oops', col: 1, padding: 120 }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
  });

  it('setVerticalAlign accepts row/col', async () => {
    const ast = createMock3ColTableAST();
    await executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setVerticalAlign', row: 2, col: 3, vAlign: 'bottom' }
    } as any, ast);
    const tblEl = ast.sections[0].nodes[0].xml as any;
    const trs = tblEl.getElementsByTagName('w:tr');
    const tcs = (trs[1] as any).getElementsByTagName('w:tc');
    expect((tcs[2] as any).getElementsByTagName('w:vAlign')[0].getAttribute('w:val')).toBe('bottom');
    expect((tcs[0] as any).getElementsByTagName('w:vAlign').length).toBe(0);
  });

  it('setVerticalAlign rejects bad row/col', async () => {
    const ast = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setVerticalAlign', row: 1.5, col: 1, vAlign: 'center' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION/);
    const ast2 = createMock3ColTableAST();
    await expect(executeTable({
      action: 'table',
      target: 't_001',
      params: { op: 'setVerticalAlign', row: 1, col: -1, vAlign: 'center' }
    } as any, ast2)).rejects.toThrow(/E_VALIDATION/);
  });
});

/** Live-document harness: blank doc + one inserted table, so captions can be asserted in saved bytes. */
async function freshCaptionTable(): Promise<{ parser: DocxParser; ast: any; tableId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id;
  await executeInsert({
    action: 'insert', type: 'table', anchor, position: 'after',
    headers: ['A', 'B'], rows: [['1', '2']],
  } as any, ast, parser);
  const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
  expect(tblNode).toBeDefined();
  return { parser, ast, tableId: tblNode.id as string };
}

/** Index of w:tbl and of the Caption-styled paragraph carrying `text` among body children in saved bytes. */
function savedCaptionOrder(parser: DocxParser, ast: any, text: string): { tableIdx: number; captionIdx: number } {
  const xml = new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const body = doc.getElementsByTagName('w:body')[0] as Element;
  const kids: Element[] = [];
  for (let i = 0; i < body.childNodes.length; i++) {
    const c = body.childNodes[i];
    if (c.nodeType === 1) kids.push(c as Element);
  }
  let tableIdx = -1;
  let captionIdx = -1;
  kids.forEach((k, i) => {
    if (k.tagName === 'w:tbl') tableIdx = i;
    if (k.tagName === 'w:p') {
      const styles = k.getElementsByTagName('w:pStyle');
      const isCaption = styles.length > 0 &&
        (styles[0].getAttribute('w:val') === 'Caption' || styles[0].getAttribute('val') === 'Caption');
      const ts = k.getElementsByTagName('w:t');
      let t = '';
      for (let j = 0; j < ts.length; j++) t += ts[j].textContent || '';
      if (isCaption && t === text) captionIdx = i;
    }
  });
  return { tableIdx, captionIdx };
}

describe('setCaption captionPosition', () => {
  it("captionPosition below places the caption after the table and stays VALID", async () => {
    const { parser, ast, tableId } = await freshCaptionTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Below', captionPosition: 'below' }
    }, ast);
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Below');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx + 1);
    expect((ast.sections[0].nodes.find((n: any) => n.id === tableId) as any).metadata.captionPosition).toBe('below');
    const res = await validateDocxBuffer(parser.rezip(ast));
    expect(res.ok).toBe(true);
  }, 120000);

  it('explicit above places the caption before the table', async () => {
    const { parser, ast, tableId } = await freshCaptionTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Above', captionPosition: 'above' }
    }, ast);
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Above');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx - 1);
    expect((ast.sections[0].nodes.find((n: any) => n.id === tableId) as any).metadata.captionPosition).toBe('above');
  }, 120000);

  it('omitted captionPosition defaults to above', async () => {
    const { parser, ast, tableId } = await freshCaptionTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Default' }
    }, ast);
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Default');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx - 1);
    expect((ast.sections[0].nodes.find((n: any) => n.id === tableId) as any).metadata.captionPosition).toBe('above');
  }, 120000);

  it('bad captionPosition fails closed with E_VALIDATION at engine and schema gate', async () => {
    const { ast, tableId } = await freshCaptionTable();
    await expect(executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Bad', captionPosition: 'sideways' }
    } as any, ast)).rejects.toThrow(/E_VALIDATION.*captionPosition/);
    expect(() => validateOperationStrict({
      action: 'table', target: tableId,
      operation: 'setCaption', captionText: 'Table 1: Bad', captionPosition: 'sideways',
    } as any)).toThrow(/E_VALIDATION/);
  });
});

/** Count every Caption-styled paragraph in saved bytes (stays exactly one across re-sets). */
function countSavedCaptions(parser: DocxParser, ast: any): number {
  const xml = new AdmZip(parser.rezip(ast)).readAsText('word/document.xml');
  const doc = new DOMParser().parseFromString(xml, 'text/xml');
  const body = doc.getElementsByTagName('w:body')[0] as Element;
  const ps = body.getElementsByTagName('w:p');
  let n = 0;
  for (let i = 0; i < ps.length; i++) {
    const styles = (ps[i] as Element).getElementsByTagName('w:pStyle');
    if (styles.length > 0 &&
      (styles[0].getAttribute('w:val') === 'Caption' || styles[0].getAttribute('val') === 'Caption')) n++;
  }
  return n;
}

/** Caption paragraph ids among live AST section nodes (exactly one expected). */
function astCaptionIds(ast: any): string[] {
  const ids: string[] = [];
  for (const s of ast.sections) {
    for (const n of s.nodes) {
      if (n.type !== 'paragraph' || !n.xml) continue;
      const pPr = (n.xml as Element).getElementsByTagName('w:pPr')[0] as Element | undefined;
      const pStyle = pPr?.getElementsByTagName('w:pStyle')[0] as Element | undefined;
      const v = pStyle?.getAttribute('w:val') ?? pStyle?.getAttribute('val');
      if (v === 'Caption') ids.push(n.id);
    }
  }
  return ids;
}

describe('setCaption re-set converges to one caption', () => {
  it('above→below moves the single caption after the table', async () => {
    const { parser, ast, tableId } = await freshCaptionTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Move Me', captionPosition: 'above' }
    }, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Move Me', captionPosition: 'below' }
    }, ast);
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Move Me');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx + 1);
    expect(countSavedCaptions(parser, ast)).toBe(1);
    expect(astCaptionIds(ast)).toHaveLength(1);
    expect((ast.sections[0].nodes.find((n: any) => n.id === tableId) as any).metadata.captionPosition).toBe('below');
  }, 120000);

  it('below→above moves the single caption before the table', async () => {
    const { parser, ast, tableId } = await freshCaptionTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Move Back', captionPosition: 'below' }
    }, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Move Back', captionPosition: 'above' }
    }, ast);
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Move Back');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx - 1);
    expect(countSavedCaptions(parser, ast)).toBe(1);
    expect(astCaptionIds(ast)).toHaveLength(1);
    expect((ast.sections[0].nodes.find((n: any) => n.id === tableId) as any).metadata.captionPosition).toBe('above');
  }, 120000);

  it('same-side re-set twice replaces text with exactly one caption', async () => {
    const { parser, ast, tableId } = await freshCaptionTable();
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: First Write', captionPosition: 'above' }
    }, ast);
    await executeTable({
      action: 'table', target: tableId,
      params: { op: 'setCaption', captionText: 'Table 1: Second Write', captionPosition: 'above' }
    }, ast);
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Second Write');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx - 1);
    expect(savedCaptionOrder(parser, ast, 'Table 1: First Write').captionIdx).toBe(-1);
    expect(countSavedCaptions(parser, ast)).toBe(1);
    expect(astCaptionIds(ast)).toHaveLength(1);
  }, 120000);
});

describe('insert table honors captionText/captionPosition', () => {
  it('places the caption below the table at insert time', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['A'], rows: [['1']],
      captionText: 'Table 1: Inserted Below', captionPosition: 'below',
    } as any, ast, parser);
    const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
    expect(tblNode).toBeDefined();
    const { tableIdx, captionIdx } = savedCaptionOrder(parser, ast, 'Table 1: Inserted Below');
    expect(tableIdx).toBeGreaterThanOrEqual(0);
    expect(captionIdx).toBe(tableIdx + 1);
    expect(countSavedCaptions(parser, ast)).toBe(1);
    expect((tblNode as any).metadata.captionPosition).toBe('below');
  }, 120000);

  it('bad captionPosition on insert fails closed with E_VALIDATION at engine and schema gate', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await expect(executeInsert({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['A'], rows: [['1']],
      captionText: 'Table 1: Bad Insert', captionPosition: 'sideways',
    } as any, ast, parser)).rejects.toThrow(/E_VALIDATION.*captionPosition/);
    expect(() => validateOperationStrict({
      action: 'insert', type: 'table', anchor, position: 'after',
      headers: ['A'], rows: [['1']],
      captionText: 'Table 1: Bad Insert', captionPosition: 'sideways',
    } as any)).toThrow(/E_VALIDATION/);
  });
});
