/**
 * Tracked-changes default wave: fresh-node revisions + tracked cell edits.
 *
 * Every markup assertion reads serialized bytes (node XML or rezip saved
 * word/document.xml), never in-memory flags alone.
 */
import { describe, it, expect, vi } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeDelete,
  executeHyperlink,
  executeInsert,
  executeReplace,
  replaceText,
  executeTable,
  wordDiff,
  emitTrackedCellReplace,
  validateOperationStrict,
} from '@attentionform/astra-engine';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const serializer = new XMLSerializer();

function makeParaNode(text: string): Element {
  return new DOMParser().parseFromString(
    `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
    'text/xml',
  ).documentElement as unknown as Element;
}

function makeReplaceAst(
  text: string,
  opts?: { version?: number; trackChanges?: boolean; nextChangeId?: number },
): any {
  const p = makeParaNode(text);
  return {
    metadata: {},
    sections: [{
      id: 'sec_root', heading: null,
      nodes: [{
        id: 'p_1', type: 'paragraph', xml: p,
        metadata: {
          text, hasCitation: false, hasComment: false, hasBookmark: false,
          hasTrackedChanges: false, hasEquation: false, hasField: false,
          hasImage: false, hasHyperlink: false,
        },
        version: opts?.version ?? 1, checksum: 'c',
      }],
      pageSetup: {}, version: 1, checksum: 's',
    }],
    styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
    settings: { trackChanges: opts?.trackChanges ?? true, defaultAuthor: 'ASTra' },
    nextChangeId: opts?.nextChangeId ?? 1,
  };
}

function nodeXml(ast: any): string {
  return serializer.serializeToString(ast.sections[0].nodes[0].xml);
}

/** Blank doc + one 2x2 table (headers A/B, rows 1/2 and 3/4) inserted live. */
async function freshTable(trackChanges = true): Promise<{ parser: DocxParser; ast: any; tableId: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  ast.settings.trackChanges = trackChanges;
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

function savedBuffer(parser: DocxParser, ast: any): Buffer {
  return parser.rezip(ast);
}

function tableCell(doc: Document, row: number, col: number): Element {
  const tbl = doc.getElementsByTagName('w:tbl')[0] as unknown as Element;
  const trs = Array.from(tbl.childNodes).filter(
    (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tr',
  ) as Element[];
  const tcs = Array.from(trs[row].childNodes).filter(
    (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tc',
  ) as Element[];
  return tcs[col];
}

function parseXml(xml: string): Document {
  return new DOMParser().parseFromString(xml, 'text/xml');
}

describe('tracked-changes default: replace on fresh nodes', () => {
  it('fresh (v1) node + tracking on emits w:ins and w:del with the old text preserved', () => {
    const ast = makeReplaceAst('old text here');
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
    expect(xml).toContain('w:delText');
    expect(xml).toContain('old');
    expect(xml).toContain('new');
    // Old baseline must survive inside delText, not be overwritten.
    expect(xml).toMatch(/<w:delText[^>]*>old<\/w:delText>/);
  });

  it('emitted revisions carry author, date and sequential change ids', () => {
    const ast = makeReplaceAst('alpha beta');
    executeReplace({ action: 'replace', target: 'p_1', text: 'alpha gamma' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('w:author="ASTra"');
    expect(xml).toContain('w:date="');
    expect(xml).toContain('w:id="1"');
    expect(xml).toContain('w:id="2"');
    expect(ast.nextChangeId).toBe(3);
  });

  it('settings trackChanges=false gives a plain replace with no revisions', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('w:del');
    expect(xml).not.toContain('delText');
    expect(xml).toContain('new text here');
    expect(ast.sections[0].nodes[0].metadata.text).toBe('new text here');
    expect(ast.sections[0].nodes[0].version).toBe(2);
  });

  it('per-op trackChanges=false overrides tracking-on settings with a plain replace', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: true });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here', trackChanges: false } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('w:del');
    expect(xml).toContain('new text here');
  });

  it('tracking-off output is byte-identical to the plain replaceText path', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    const viaOp = nodeXml(ast);
    const plain = makeParaNode('old text here');
    replaceText(plain as unknown as Element, 'new text here', false);
    const viaPlain = serializer.serializeToString(plain);
    expect(viaOp).toBe(viaPlain);
  });

  it('tracked replace bumps version 1->2 and updates metadata text', () => {
    const ast = makeReplaceAst('old text here');
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    expect(ast.sections[0].nodes[0].version).toBe(2);
    expect(ast.sections[0].nodes[0].metadata.text).toBe('new text here');
  });

  it('already-edited (v2) node + tracking on still emits revisions', () => {
    const ast = makeReplaceAst('second text', { version: 2 });
    executeReplace({ action: 'replace', target: 'p_1', text: 'third text' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
    expect(ast.sections[0].nodes[0].version).toBe(3);
  });

  it('no-op tracked replace (same text) mints no change ids and no revisions', () => {
    const ast = makeReplaceAst('same text');
    executeReplace({ action: 'replace', target: 'p_1', text: 'same text' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('<w:del');
    expect(ast.nextChangeId).toBe(1);
  });

  it('tracking-off replace consumes no change ids', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    expect(ast.nextChangeId).toBe(1);
  });

  it('end-to-end saved bytes contain revisions for a first edit with tracking on', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    ast.settings.trackChanges = true;
    expect(ast.settings.trackChanges).toBe(true);
    const target = ast.sections[0].nodes[0];
    expect(target.version).toBe(1);
    const oldText: string = target.metadata.text;
    executeReplace({ action: 'replace', target: target.id, text: `${oldText} appended` } as any, ast);
    const xml = savedDocumentXml(parser, ast);
    expect(xml).toContain('<w:ins');
    expect(xml).toContain('appended');
  });
});

describe('tracked-changes default: table cell edits', () => {
  it('updateCell with tracking on emits w:ins/w:del scoped to the cell in saved bytes', async () => {
    const { parser, ast, tableId } = await freshTable(true);
    const before = ast.nextChangeId;
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'ONE' } } as any, ast);
    const xml = savedDocumentXml(parser, ast);
    const cellXml = serializer.serializeToString(tableCell(parseXml(xml), 1, 0));
    expect(cellXml).toContain('<w:del');
    expect(cellXml).toContain('<w:ins');
    expect(cellXml).toMatch(/<w:delText[^>]*>1<\/w:delText>/);
    expect(cellXml).toContain('ONE');
    // Structural whole-cell markers are for row/column insert/delete, not
    // text edits: the cell properties must not gain cellIns/cellDel.
    expect(cellXml).not.toContain('w:cellIns');
    expect(cellXml).not.toContain('w:cellDel');
    expect(ast.nextChangeId).toBeGreaterThan(before);
  });

  it('updateCell with tracking off stays a plain overwrite in saved bytes', async () => {
    const { parser, ast, tableId } = await freshTable(false);
    const before = ast.nextChangeId;
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'ONE' } } as any, ast);
    const xml = savedDocumentXml(parser, ast);
    const cellXml = serializer.serializeToString(tableCell(parseXml(xml), 1, 0));
    expect(cellXml).not.toContain('w:ins');
    expect(cellXml).not.toContain('w:del');
    expect(cellXml).toContain('ONE');
    expect(ast.nextChangeId).toBe(before);
  });

  it('tracked cell edit keeps valid cell structure: revisions sit inside w:p, cell keeps a paragraph', async () => {
    const { parser, ast, tableId } = await freshTable(true);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 3, colIndex: 2, text: 'FOUR' } } as any, ast);
    const xml = savedDocumentXml(parser, ast);
    const doc = parseXml(xml);
    const cell = tableCell(doc, 2, 1);
    const ownPs = Array.from(cell.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:p',
    ) as Element[];
    expect(ownPs.length).toBeGreaterThanOrEqual(1);
    for (const p of ownPs) {
      const revs = Array.from(p.childNodes).filter(
        (n) => n.nodeType === 1 && ((n as Element).tagName === 'w:ins' || (n as Element).tagName === 'w:del'),
      );
      for (const r of revs) {
        expect((r as Element).getAttribute('w:author')).toBe('ASTra');
        expect((r as Element).getAttribute('w:id')).toMatch(/^\d+$/);
      }
    }
    const cellText = cell.textContent || '';
    expect(cellText).toContain('FOUR');
  });

  it('tracked cell edit document passes the OOXML validator', async () => {
    const { parser, ast, tableId } = await freshTable(true);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'ONE' } } as any, ast);
    const buf = savedBuffer(parser, ast);
    const res = await validateDocxBuffer(buf);
    expect(res.ok).toBe(true);
  }, 60000);
});

describe('tracked changes wave 1B fixes', () => {
  it('Issue 1: tableOp strict schema accepts trackChanges flag', () => {
    const res = validateOperationStrict({
      action: 'table',
      target: 't_1',
      op: 'updateCell',
      rowIndex: 1,
      colIndex: 1,
      text: 'hello',
      trackChanges: false,
    });
    expect(res.ok).toBe(true);
    expect((res.operation as any).trackChanges).toBe(false);
  });

  it('Issue 1: per-op trackChanges in params overrides settings in replace', () => {
    // trackChanges=false in params overrides settings.trackChanges=true
    const ast1 = makeReplaceAst('old text', { trackChanges: true });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text', params: { trackChanges: false } } as any, ast1);
    expect(nodeXml(ast1)).not.toContain('w:ins');
    expect(nodeXml(ast1)).not.toContain('w:del');

    // trackChanges=true in params overrides settings.trackChanges=false
    const ast2 = makeReplaceAst('old text', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text', params: { trackChanges: true } } as any, ast2);
    expect(nodeXml(ast2)).toContain('w:ins');
    expect(nodeXml(ast2)).toContain('w:del');

    // trackChanges=true flat overrides settings.trackChanges=false
    const ast3 = makeReplaceAst('old text', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text', trackChanges: true } as any, ast3);
    expect(nodeXml(ast3)).toContain('w:ins');
    expect(nodeXml(ast3)).toContain('w:del');
  });

  it('Issue 1: per-op trackChanges (flat & params) overrides settings in table updateCell', async () => {
    // flat trackChanges: false overrides settings=true
    const t1 = await freshTable(true);
    const before1 = t1.ast.nextChangeId;
    await executeTable({ action: 'table', target: t1.tableId, trackChanges: false, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'VAL' } } as any, t1.ast);
    const cellXml1 = serializer.serializeToString(tableCell(parseXml(savedDocumentXml(t1.parser, t1.ast)), 1, 0));
    expect(cellXml1).not.toContain('w:ins');
    expect(cellXml1).not.toContain('w:del');
    expect(t1.ast.nextChangeId).toBe(before1);

    // params.trackChanges: false overrides settings=true
    const t2 = await freshTable(true);
    const before2 = t2.ast.nextChangeId;
    await executeTable({ action: 'table', target: t2.tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'VAL', trackChanges: false } } as any, t2.ast);
    const cellXml2 = serializer.serializeToString(tableCell(parseXml(savedDocumentXml(t2.parser, t2.ast)), 1, 0));
    expect(cellXml2).not.toContain('w:ins');
    expect(cellXml2).not.toContain('w:del');
    expect(t2.ast.nextChangeId).toBe(before2);

    // flat trackChanges: true overrides settings=false
    const t3 = await freshTable(false);
    const before3 = t3.ast.nextChangeId;
    await executeTable({ action: 'table', target: t3.tableId, trackChanges: true, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'VAL' } } as any, t3.ast);
    const cellXml3 = serializer.serializeToString(tableCell(parseXml(savedDocumentXml(t3.parser, t3.ast)), 1, 0));
    expect(cellXml3).toContain('w:ins');
    expect(cellXml3).toContain('w:del');
    expect(t3.ast.nextChangeId).toBeGreaterThan(before3);

    // params.trackChanges: true overrides settings=false
    const t4 = await freshTable(false);
    const before4 = t4.ast.nextChangeId;
    await executeTable({ action: 'table', target: t4.tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'VAL', trackChanges: true } } as any, t4.ast);
    const cellXml4 = serializer.serializeToString(tableCell(parseXml(savedDocumentXml(t4.parser, t4.ast)), 1, 0));
    expect(cellXml4).toContain('w:ins');
    expect(cellXml4).toContain('w:del');
    expect(t4.ast.nextChangeId).toBeGreaterThan(before4);
  });

  it('Issue 2: repeated replace does not leave ghost empty <w:ins> or <w:del> wrappers behind', () => {
    const ast = makeReplaceAst('initial text');
    executeReplace({ action: 'replace', target: 'p_1', text: 'second text' } as any, ast);
    executeReplace({ action: 'replace', target: 'p_1', text: 'third text' } as any, ast);
    executeReplace({ action: 'replace', target: 'p_1', text: 'fourth text' } as any, ast);
    const pEl = ast.sections[0].nodes[0].xml as Element;
    // Check all w:ins and w:del wrappers - none should be empty shells
    for (let i = 0; i < pEl.childNodes.length; i++) {
      const child = pEl.childNodes[i] as Element;
      if (child.nodeType === 1 && (child.tagName === 'w:ins' || child.tagName === 'w:del')) {
        expect(child.hasChildNodes()).toBe(true);
      }
    }
  });

  it('Issue 3: replace on non-paragraph node throws E_VALIDATION', () => {
    const tableEl = new DOMParser().parseFromString(
      `<w:tbl xmlns:w="${W}"><w:tr><w:tc><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
      'text/xml',
    ).documentElement as unknown as Element;
    const ast = {
      sections: [{
        id: 'sec_1',
        nodes: [{
          id: 'tbl_1',
          type: 'table',
          xml: tableEl,
          metadata: { text: 'cell' },
          version: 1,
        }],
      }],
      settings: { trackChanges: true },
      nextChangeId: 1,
    } as any;

    expect(() => {
      executeReplace({ action: 'replace', target: 'tbl_1', text: 'cannot replace table directly' } as any, ast);
    }).toThrowError(/E_VALIDATION/);

    try {
      executeReplace({ action: 'replace', target: 'tbl_1', text: 'cannot replace table directly' } as any, ast);
    } catch (err: any) {
      expect(err.code).toBe('E_VALIDATION');
      expect(err.retryable).toBe(false);
    }
  });

  it('Issue 4: wordDiff OOM guard triggers on massive inputs without blowing up memory', () => {
    // 2000 words * 2000 words = 4,000,000 > 2,500,000 threshold
    const wordsA = Array(2000).fill('alpha').join(' ');
    const wordsB = Array(2000).fill('beta').join(' ');
    const diff = wordDiff(wordsA, wordsB);
    expect(diff).toHaveLength(2);
    expect(diff[0]).toEqual({ type: 'delete', text: wordsA });
    expect(diff[1]).toEqual({ type: 'insert', text: wordsB });
  });

  it('Issue 5: metadata.hasTrackedChanges is set to true on paragraph and table replace', async () => {
    // Paragraph replace
    const ast = makeReplaceAst('old paragraph text');
    expect(ast.sections[0].nodes[0].metadata.hasTrackedChanges).toBe(false);
    executeReplace({ action: 'replace', target: 'p_1', text: 'new paragraph text' } as any, ast);
    expect(ast.sections[0].nodes[0].metadata.hasTrackedChanges).toBe(true);

    // Table updateCell
    const { ast: tblAst, tableId } = await freshTable(true);
    const tblNode = tblAst.sections[0].nodes.find((n: any) => n.id === tableId);
    expect(tblNode.metadata.hasTrackedChanges).toBe(false);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'REVISED_CELL' } } as any, tblAst);
    expect(tblNode.metadata.hasTrackedChanges).toBe(true);
  });

  it('wordDiff on identical strings exceeding threshold returns equal chunk and executeReplace emits 0 revisions', () => {
    // 2000 words * 2000 words = 4,000,000 > 2,500,000 threshold
    const largeText = Array(2000).fill('alpha').join(' ');
    const diff = wordDiff(largeText, largeText);
    expect(diff).toEqual([{ type: 'equal', text: largeText }]);

    const ast = makeReplaceAst(largeText, { trackChanges: true });
    executeReplace({ action: 'replace', target: 'p_1', text: largeText } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('w:ins');
    expect(xml).not.toContain('w:del');
    expect(ast.nextChangeId).toBe(1);
    expect(ast.sections[0].nodes[0].metadata.hasTrackedChanges).toBe(false);
  });
});

describe('tracked structural table ops fail closed under explicit trackChanges:true', () => {
  function liveRowCount(ast: any, tableId: string): number {
    const node = ast.sections[0].nodes.find((n: any) => n.id === tableId);
    return Array.from((node.xml as Element).childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tr',
    ).length;
  }

  function liveColCount(ast: any, tableId: string): number {
    const node = ast.sections[0].nodes.find((n: any) => n.id === tableId);
    const tbl = node.xml as Element;
    const firstTr = Array.from(tbl.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tr',
    ) as Element;
    return Array.from(firstTr.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tc',
    ).length;
  }

  async function expectFailClosed(
    opName: string,
    params: Record<string, unknown>,
    marker: RegExp,
  ): Promise<void> {
    const { ast, tableId } = await freshTable(true);
    const rowsBefore = liveRowCount(ast, tableId);
    const colsBefore = liveColCount(ast, tableId);
    const textBefore: string = ast.sections[0].nodes.find((n: any) => n.id === tableId).metadata.text;
    let caught: any = null;
    try {
      await executeTable({ action: 'table', target: tableId, trackChanges: true, params: { op: opName, ...params } } as any, ast);
    } catch (err: any) {
      caught = err;
    }
    expect(caught).not.toBeNull();
    expect(String(caught.message)).toMatch(/E_VALIDATION/);
    expect(String(caught.message)).toMatch(/trackChanges/);
    expect(String(caught.message)).toMatch(marker);
    expect(caught.code).toBe('E_VALIDATION');
    expect(caught.retryable).toBe(false);
    // Fail-closed means nothing applied: shape and text identical.
    expect(liveRowCount(ast, tableId)).toBe(rowsBefore);
    expect(liveColCount(ast, tableId)).toBe(colsBefore);
    expect(ast.sections[0].nodes.find((n: any) => n.id === tableId).metadata.text).toBe(textBefore);
  }

  it('addRow names the missing structural marker and changes nothing', async () => {
    await expectFailClosed('addRow', { afterIndex: 1 }, /cellIns/);
  });

  it('deleteRow names the missing structural marker and changes nothing', async () => {
    await expectFailClosed('deleteRow', { index: 1 }, /cellDel/);
  });

  it('addColumn names the missing structural marker and changes nothing', async () => {
    await expectFailClosed('addColumn', { afterIndex: 1 }, /cellIns/);
  });

  it('deleteColumn names the missing structural marker and changes nothing', async () => {
    await expectFailClosed('deleteColumn', { index: 1 }, /cellDel/);
  });

  it('mergeCells names the missing structural marker and changes nothing', async () => {
    await expectFailClosed(
      'mergeCells',
      { fromRow: 1, fromCol: 1, toRow: 2, toCol: 2 },
      /cellMerge/,
    );
  });

  it('splitCell names the missing structural marker and changes nothing', async () => {
    await expectFailClosed('splitCell', { rowIndex: 1, colIndex: 1 }, /tcPr/);
  });

  it('params-nested trackChanges:true also fails closed', async () => {
    const { ast, tableId } = await freshTable(true);
    const rowsBefore = liveRowCount(ast, tableId);
    let caught: any = null;
    try {
      await executeTable({ action: 'table', target: tableId, params: { op: 'addRow', afterIndex: 1, trackChanges: true } } as any, ast);
    } catch (err: any) {
      caught = err;
    }
    expect(caught).not.toBeNull();
    expect(String(caught.message)).toMatch(/E_VALIDATION.*trackChanges/);
    expect(caught.code).toBe('E_VALIDATION');
    expect(caught.retryable).toBe(false);
    expect(liveRowCount(ast, tableId)).toBe(rowsBefore);
  });

  it('explicit trackChanges:false still applies structural ops outright', async () => {
    const { ast, tableId } = await freshTable(true);
    const rowsBefore = liveRowCount(ast, tableId);
    await executeTable({ action: 'table', target: tableId, trackChanges: false, params: { op: 'addRow', afterIndex: 1 } } as any, ast);
    expect(liveRowCount(ast, tableId)).toBe(rowsBefore + 1);
  });

  it('ambient settings tracking without a per-op flag keeps the legacy apply path', async () => {
    // The default blank document ships with tracking enabled; structural ops
    // without an explicit per-op flag apply directly (documented compat).
    const { ast, tableId } = await freshTable(true);
    const rowsBefore = liveRowCount(ast, tableId);
    await executeTable({ action: 'table', target: tableId, params: { op: 'addRow', afterIndex: 1 } } as any, ast);
    expect(liveRowCount(ast, tableId)).toBe(rowsBefore + 1);
  });

  it('updateCell with an explicit per-op flag still emits paragraph-level revisions', async () => {
    const { parser, ast, tableId } = await freshTable(false);
    await executeTable({ action: 'table', target: tableId, trackChanges: true, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'ONE' } } as any, ast);
    const cellXml = serializer.serializeToString(tableCell(parseXml(savedDocumentXml(parser, ast)), 1, 0));
    expect(cellXml).toContain('<w:del');
    expect(cellXml).toContain('<w:ins');
    expect(cellXml).toContain('ONE');
  });
});

describe('tracked delete', () => {
  it('explicit trackChanges:true on a paragraph records w:del and retains the node', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    const before = ast.nextChangeId;
    const res = executeDelete({ action: 'delete', target: target.id, trackChanges: true } as any, ast, parser);
    expect(res.deleted).toEqual([target.id]);
    expect(res.skipped).toEqual([]);
    // Retained, not removed: revision markup carries the deletion.
    const kept = ast.sections[0].nodes.find((n: any) => n.id === target.id);
    expect(kept).toBeDefined();
    const xml = serializer.serializeToString(kept.xml);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('w:delText');
    expect(kept.metadata.text).toBe('');
    expect(kept.metadata.hasTrackedChanges).toBe(true);
    expect(kept.version).toBe(2);
    expect(ast.nextChangeId).toBeGreaterThan(before);
    // Saved bytes carry the deletion and validate.
    expect(savedDocumentXml(parser, ast)).toContain('<w:del');
    const vres = await validateDocxBuffer(savedBuffer(parser, ast));
    expect(vres.ok).toBe(true);
  }, 60000);

  it('params-nested trackChanges:true on a paragraph also records w:del', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    const res = executeDelete({ action: 'delete', target: target.id, params: { trackChanges: true } } as any, ast, parser);
    expect(res.deleted).toEqual([target.id]);
    expect(ast.sections[0].nodes.find((n: any) => n.id === target.id)).toBeDefined();
    expect(serializer.serializeToString(target.xml)).toContain('<w:del');
  });

  it('explicit trackChanges:true on a table node fails closed and removes nothing', async () => {
    const { ast, tableId } = await freshTable(false);
    const nodesBefore = ast.sections[0].nodes.length;
    let caught: any = null;
    try {
      executeDelete({ action: 'delete', target: tableId, trackChanges: true } as any, ast);
    } catch (err: any) {
      caught = err;
    }
    expect(caught).not.toBeNull();
    expect(String(caught.message)).toMatch(/E_VALIDATION/);
    expect(String(caught.message)).toMatch(/trackChanges/);
    expect(String(caught.message)).toMatch(/cellDel/);
    expect(caught.code).toBe('E_VALIDATION');
    expect(caught.retryable).toBe(false);
    expect(ast.sections[0].nodes.length).toBe(nodesBefore);
    expect(ast.sections[0].nodes.find((n: any) => n.id === tableId)).toBeDefined();
  });

  it('explicit trackChanges:false deletes outright (legacy removal)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    const nodesBefore = ast.sections[0].nodes.length;
    const res = executeDelete({ action: 'delete', target: target.id, trackChanges: false } as any, ast, parser);
    expect(res.deleted).toEqual([target.id]);
    expect(ast.sections[0].nodes.length).toBe(nodesBefore - 1);
    expect(ast.sections[0].nodes.find((n: any) => n.id === target.id)).toBeUndefined();
  });

  it('missing targets under explicit tracking keep the legacy warn+skip path', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const res = executeDelete({ action: 'delete', target: 'ghost_tracked', trackChanges: true } as any, ast, parser);
      expect(res.deleted).toEqual([]);
      expect(res.skipped).toEqual(['ghost_tracked']);
    } finally {
      warnSpy.mockRestore();
    }
  });
});

describe('tracked replace hyperlink shells and chained history', () => {
  function emptyShells(p: Element, tag: string): number {
    const els = p.getElementsByTagName(tag);
    let empty = 0;
    for (let i = 0; i < els.length; i++) {
      const el = els[i] as unknown as Element;
      let hasEl = false;
      for (let j = 0; j < el.childNodes.length; j++) {
        if (el.childNodes[j].nodeType === 1) {
          hasEl = true;
          break;
        }
      }
      if (!hasEl) empty++;
    }
    return empty;
  }

  it('paragraph replace re-emits the hyperlink wrapper so the LINK survives', () => {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
      `<w:hyperlink r:id="rId5" w:anchor="bm1" w:history="1"><w:r><w:t xml:space="preserve">click here</w:t></w:r></w:hyperlink>` +
      `<w:r><w:t xml:space="preserve"> and more</w:t></w:r></w:p>`,
      'text/xml',
    ).documentElement as unknown as Element;
    const ast = makeReplaceAst('click here and more');
    ast.sections[0].nodes[0].xml = p;
    executeReplace({ action: 'replace', target: 'p_1', text: 'click here and updated' } as any, ast);
    const out = ast.sections[0].nodes[0].xml as Element;
    expect(emptyShells(out, 'w:hyperlink')).toBe(0);
    expect(emptyShells(out, 'w:ins')).toBe(0);
    expect(emptyShells(out, 'w:del')).toBe(0);
    const xml = serializer.serializeToString(out);
    // Old link-adjacent text survives as revision content, new text is present.
    expect(xml).toContain('click here');
    expect(xml).toMatch(/<w:delText[^>]*>more<\/w:delText>/);
    expect(xml).toContain('updated');
    // LINK survives, not just its text: exactly one live wrapper with identical attrs.
    const hls = out.getElementsByTagName('w:hyperlink');
    expect(hls.length).toBeGreaterThan(0);
    const hl = hls[0] as unknown as Element;
    expect(hl.getAttribute('r:id')).toBe('rId5');
    expect(hl.getAttribute('w:anchor')).toBe('bm1');
    expect(hl.getAttribute('w:history')).toBe('1');
    expect(hl.childNodes.length).toBeGreaterThan(0);
  });

  it('3-chain paragraph replaces retain every link of deletion history', () => {
    const ast = makeReplaceAst('red green blue');
    executeReplace({ action: 'replace', target: 'p_1', text: 'RED green blue' } as any, ast);
    executeReplace({ action: 'replace', target: 'p_1', text: 'RED GREEN blue' } as any, ast);
    executeReplace({ action: 'replace', target: 'p_1', text: 'RED GREEN BLUE' } as any, ast);
    const out = ast.sections[0].nodes[0].xml as Element;
    const xml = serializer.serializeToString(out);
    expect(xml).toMatch(/<w:delText[^>]*>red<\/w:delText>/);
    expect(xml).toMatch(/<w:delText[^>]*>green<\/w:delText>/);
    expect(xml).toMatch(/<w:delText[^>]*>blue<\/w:delText>/);
    // Visible text (plain w:t + insertions, excluding w:delText) is the final revision.
    const visibleDoc = parseXml(xml);
    const visibleText = Array.from(visibleDoc.getElementsByTagName('w:t'))
      .map((t) => t.textContent || '')
      .join('');
    expect(visibleText).toBe('RED GREEN BLUE');
    expect((xml.match(/<w:del[\s>]/g) || []).length).toBeGreaterThanOrEqual(3);
    expect(emptyShells(out, 'w:ins')).toBe(0);
    expect(emptyShells(out, 'w:del')).toBe(0);
    expect(ast.sections[0].nodes[0].metadata.text).toBe('RED GREEN BLUE');
  });

  it('3-chain cell replaces retain every link of deletion history in saved bytes', async () => {
    const { parser, ast, tableId } = await freshTable(true);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'one' } } as any, ast);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'ONE' } } as any, ast);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: '1ST' } } as any, ast);
    const cellXml = serializer.serializeToString(tableCell(parseXml(savedDocumentXml(parser, ast)), 1, 0));
    expect(cellXml).toMatch(/<w:delText[^>]*>1<\/w:delText>/);
    expect(cellXml).toMatch(/<w:delText[^>]*>one<\/w:delText>/);
    expect(cellXml).toMatch(/<w:delText[^>]*>ONE<\/w:delText>/);
    expect(cellXml).toContain('1ST');
    expect(cellXml).not.toContain('w:cellIns');
    expect(cellXml).not.toContain('w:cellDel');
    const buf = savedBuffer(parser, ast);
    const res = await validateDocxBuffer(buf);
    expect(res.ok).toBe(true);
  }, 60000);

  it('tracked cell replace re-emits the hyperlink wrapper with identical anchor', async () => {
    const { parser, ast, tableId } = await freshTable(true);
    const tblNode = ast.sections[0].nodes.find((n: any) => n.id === tableId);
    const tblEl = tblNode.xml as Element;
    const trs = Array.from(tblEl.childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tr',
    ) as Element[];
    const tcs = Array.from(trs[1].childNodes).filter(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:tc',
    ) as Element[];
    const p = Array.from(tcs[0].childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:p',
    ) as Element;
    const r = Array.from(p.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:r',
    ) as Element;
    const doc = tblEl.ownerDocument!;
    // Internal anchor link (no r:id, so no rel entry is needed for validation).
    const hl = doc.createElementNS(W, 'w:hyperlink');
    hl.setAttribute('w:anchor', 'bmCell');
    hl.setAttribute('w:history', '1');
    p.insertBefore(hl, r);
    hl.appendChild(r);
    await executeTable({ action: 'table', target: tableId, params: { op: 'updateCell', rowIndex: 2, colIndex: 1, text: 'ONE' } } as any, ast);
    const savedXml = savedDocumentXml(parser, ast);
    const cellXml = serializer.serializeToString(tableCell(parseXml(savedXml), 1, 0));
    const cellDoc = parseXml(cellXml);
    const hls = cellDoc.getElementsByTagName('w:hyperlink');
    expect(hls.length).toBeGreaterThan(0);
    const live = hls[0] as unknown as Element;
    expect(live.getAttribute('w:anchor')).toBe('bmCell');
    expect(live.getAttribute('w:history')).toBe('1');
    for (let i = 0; i < hls.length; i++) {
      const el = hls[i] as unknown as Element;
      let hasEl = false;
      for (let j = 0; j < el.childNodes.length; j++) {
        if (el.childNodes[j].nodeType === 1) {
          hasEl = true;
          break;
        }
      }
      expect(hasEl).toBe(true);
    }
    expect(cellXml).toContain('<w:del');
    expect(cellXml).toContain('<w:ins');
    expect(cellXml).toContain('ONE');
    const res = await validateDocxBuffer(savedBuffer(parser, ast));
    expect(res.ok).toBe(true);
  }, 60000);

  it('paragraph tracked replace on an anchored link validates with the anchor intact', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    const p = target.xml as Element;
    const r = Array.from(p.childNodes).find(
      (n) => n.nodeType === 1 && (n as Element).tagName === 'w:r',
    ) as Element;
    const doc = p.ownerDocument!;
    const hl = doc.createElementNS(W, 'w:hyperlink');
    hl.setAttribute('w:anchor', 'bmPara');
    hl.setAttribute('w:history', '1');
    p.insertBefore(hl, r);
    hl.appendChild(r);
    const oldText: string = target.metadata.text;
    executeReplace({ action: 'replace', target: target.id, text: `${oldText} updated` } as any, ast);
    const out = target.xml as Element;
    const hls = out.getElementsByTagName('w:hyperlink');
    expect(hls.length).toBeGreaterThan(0);
    expect((hls[0] as unknown as Element).getAttribute('w:anchor')).toBe('bmPara');
    expect((hls[0] as unknown as Element).getAttribute('w:history')).toBe('1');
    expect(savedDocumentXml(parser, ast)).toContain('w:hyperlink');
    const res = await validateDocxBuffer(savedBuffer(parser, ast));
    expect(res.ok).toBe(true);
  }, 60000);

  it('paragraph tracked replace on an external link keeps href target and validates', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    executeHyperlink({ action: 'hyperlink', target: target.id, text: 'click here', url: 'https://example.com' } as any, ast, parser);
    const beforeHl = (target.xml as Element).getElementsByTagName('w:hyperlink')[0] as unknown as Element;
    const beforeRid = beforeHl.getAttribute('r:id');
    expect(beforeRid).toBeTruthy();
    const hrefBefore: string = (ast.relationships as any)[beforeRid!]?.target;
    expect(hrefBefore).toBe('https://example.com');
    // Refresh visible text (hyperlink op appends without touching metadata.text).
    const tEls = (target.xml as Element).getElementsByTagName('w:t');
    let visible = '';
    for (let i = 0; i < tEls.length; i++) visible += (tEls[i] as unknown as Element).textContent || '';
    target.metadata.text = visible;
    executeReplace({ action: 'replace', target: target.id, text: `${visible} updated` } as any, ast);
    const out = target.xml as Element;
    const hls = out.getElementsByTagName('w:hyperlink');
    expect(hls.length).toBeGreaterThan(0);
    const afterRid = (hls[0] as unknown as Element).getAttribute('r:id');
    expect(afterRid).toBe(beforeRid);
    expect((ast.relationships as any)[afterRid!]?.target).toBe('https://example.com');
    const res = await validateDocxBuffer(savedBuffer(parser, ast));
    expect(res.ok).toBe(true);
  }, 60000);
});



describe('tracked replace with pretty-printed whitespace before pPr', () => {
  it('chained replaces keep w:pPr first with history intact', async () => {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">\n      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>\n      <w:r><w:t xml:space="preserve">Astra Badge</w:t></w:r>\n    </w:p>`,
      'text/xml',
    ).documentElement as unknown as Element;
    const ast: any = {
      metadata: {},
      sections: [{
        id: 'sec_root', heading: null,
        nodes: [{
          id: 'p_ws', type: 'paragraph', xml: p,
          metadata: {
            text: 'Astra Badge', hasCitation: false, hasComment: false, hasBookmark: false,
            hasTrackedChanges: false, hasEquation: false, hasField: false,
            hasImage: false, hasHyperlink: false,
          },
          version: 1, checksum: 'c',
        }],
        pageSetup: {}, version: 1, checksum: 's',
      }],
      styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
      relationships: {}, headers: {}, footers: {}, mediaFiles: {},
      settings: { trackChanges: true, defaultAuthor: 'ASTra' }, nextChangeId: 1,
    };
    executeReplace({ action: 'replace', target: 'p_ws', text: 'EDIT ONE' } as any, ast);
    executeReplace({ action: 'replace', target: 'p_ws', text: 'EDIT TWO' } as any, ast);
    const xml = serializer.serializeToString(ast.sections[0].nodes[0].xml);
    const pPrIdx = xml.indexOf('<w:pPr');
    const delIdx = xml.indexOf('<w:del');
    expect(pPrIdx).toBeGreaterThanOrEqual(0);
    expect(delIdx).toBeGreaterThanOrEqual(0);
    expect(pPrIdx).toBeLessThan(delIdx);
    expect(xml).toContain('EDIT ');
    expect(xml).toContain('>TWO</w:t>');
  }, 60000);
});

describe('replace follows document setting with explicit override', () => {
  it('stays silent when the document setting is off', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
    expect(xml).toContain('new text here');
  });

  it('stays silent when the document setting is absent', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: false });
    delete (ast.settings as any).trackChanges;
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
    expect(xml).toContain('new text here');
  });

  it('marks revisions when the caller passes true while the document setting is off', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here', trackChanges: true } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
  });

  it('marks revisions when the document setting is on and no caller flag is given', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: true });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
  });

  it('blank text replacement stays silent and updates bookkeeping', () => {
    const ast = makeReplaceAst('hello', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: '' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
    expect(ast.sections[0].nodes[0].metadata.text).toBe('');
    expect(ast.sections[0].nodes[0].version).toBe(2);
  });

  it('stays silent when the caller passes false while the document setting is on', () => {
    const ast = makeReplaceAst('old text here', { trackChanges: true });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text here', trackChanges: false } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('<w:ins');
    expect(xml).not.toContain('<w:del');
    expect(xml).toContain('new text here');
  });

  it('params-nested flag overrides the document setting in either direction', () => {
    const silent = makeReplaceAst('old text', { trackChanges: true });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text', params: { trackChanges: false } } as any, silent);
    expect(nodeXml(silent)).not.toContain('<w:ins');
    expect(nodeXml(silent)).not.toContain('<w:del');

    const marked = makeReplaceAst('old text', { trackChanges: false });
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text', params: { trackChanges: true } } as any, marked);
    expect(nodeXml(marked)).toContain('<w:ins');
    expect(nodeXml(marked)).toContain('<w:del');
  });
});

describe('replace carries note anchors onto fresh runs', () => {
  function noteAst(kind: 'footnote' | 'endnote', id: string, text: string, on: boolean): any {
    const refTag = kind === 'footnote' ? 'w:footnoteReference' : 'w:endnoteReference';
    const styleVal = kind === 'footnote' ? 'FootnoteReference' : 'EndnoteReference';
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}"><w:r><w:t xml:space="preserve">${text} </w:t></w:r>` +
      `<w:r><w:rPr><w:rStyle w:val="${styleVal}"/><w:vertAlign w:val="superscript"/></w:rPr>` +
      `<${refTag} w:id="${id}"/></w:r></w:p>`,
      'text/xml',
    ).documentElement as unknown as Element;
    const ast = makeReplaceAst(`${text} `, { trackChanges: on });
    ast.sections[0].nodes[0].xml = p;
    return ast;
  }

  it('keeps footnote anchor on plain rewrite', () => {
    const ast = noteAst('footnote', '7', 'hello', false);
    executeReplace({ action: 'replace', target: 'p_1', text: 'bye' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('w:footnoteReference');
    expect(xml).toContain('w:id="7"');
    expect(xml).toContain('bye');
  });

  it('keeps footnote anchor on marked rewrite', () => {
    const ast = noteAst('footnote', '8', 'hello world', true);
    executeReplace({ action: 'replace', target: 'p_1', text: 'bye world' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
    expect(xml).toContain('w:footnoteReference');
    expect(xml).toContain('w:id="8"');
  });

  it('keeps endnote anchor on plain rewrite', () => {
    const ast = noteAst('endnote', '9', 'hi', false);
    executeReplace({ action: 'replace', target: 'p_1', text: 'yo' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('w:endnoteReference');
    expect(xml).toContain('w:id="9"');
    expect(xml).toContain('yo');
  });

  it('keeps endnote anchor on marked rewrite', () => {
    const ast = noteAst('endnote', '10', 'hello world', true);
    executeReplace({ action: 'replace', target: 'p_1', text: 'bye world' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
    expect(xml).toContain('w:endnoteReference');
    expect(xml).toContain('w:id="10"');
  });
});

describe('replace reset style clears inherited text properties', () => {
  const italicMark = /<w:i[\s/>]/;
  function styledAst(on: boolean): any {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}"><w:r><w:rPr><w:color w:val="FF00FF"/><w:i/><w:sz w:val="24"/></w:rPr>` +
      `<w:t xml:space="preserve">styled</w:t></w:r></w:p>`,
      'text/xml',
    ).documentElement as unknown as Element;
    const ast = makeReplaceAst('styled', { trackChanges: on });
    ast.sections[0].nodes[0].xml = p;
    return ast;
  }

  it('drops color and italic when reset is requested', () => {
    const ast = styledAst(false);
    executeReplace({ action: 'replace', target: 'p_1', text: 'plain', preserveFormatting: true, resetStyle: true } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).not.toContain('FF00FF');
    expect(xml).not.toMatch(italicMark);
    expect(xml).toContain('plain');
  });

  it('keeps color and italic when only preserve is requested', () => {
    const ast = styledAst(false);
    executeReplace({ action: 'replace', target: 'p_1', text: 'plain', preserveFormatting: true } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('FF00FF');
    expect(xml).toMatch(italicMark);
  });

  it('drops color and italic on marked rewrite when reset is requested', () => {
    const ast = styledAst(true);
    executeReplace({ action: 'replace', target: 'p_1', text: 'new text', resetStyle: true } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
    expect(xml).not.toContain('FF00FF');
    expect(xml).not.toMatch(italicMark);
  });

  it('keeps footnote anchor while dropping styling on marked rewrite when reset is requested', () => {
    const p = new DOMParser().parseFromString(
      `<w:p xmlns:w="${W}"><w:r><w:rPr><w:color w:val="FF00FF"/><w:i/><w:sz w:val="24"/></w:rPr>` +
      `<w:t xml:space="preserve">styled </w:t></w:r>` +
      `<w:r><w:rPr><w:rStyle w:val="FootnoteReference"/><w:vertAlign w:val="superscript"/></w:rPr>` +
      `<w:footnoteReference w:id="11"/></w:r></w:p>`,
      'text/xml',
    ).documentElement as unknown as Element;
    const ast = makeReplaceAst('styled ', { trackChanges: true });
    ast.sections[0].nodes[0].xml = p;
    executeReplace({ action: 'replace', target: 'p_1', text: 'plain', resetStyle: true } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:del');
    expect(xml).toContain('<w:ins');
    expect(xml).not.toContain('FF00FF');
    expect(xml).not.toMatch(italicMark);
    expect(xml).toContain('w:footnoteReference');
    expect(xml).toContain('w:id="11"');
    expect(xml).toContain('plain');
  });
});

describe('tracked replace keeps boundary whitespace visible in print output', () => {
  function visibleText(xml: string): string {
    const doc = parseXml(xml);
    const ts = doc.getElementsByTagName('w:t');
    let out = '';
    for (let i = 0; i < ts.length; i++) out += (ts[i] as Element).textContent || '';
    return out;
  }

  it('tab between adjacent revision runs carries xml:space preserve', () => {
    const ast = makeReplaceAst('a b');
    executeReplace({ action: 'replace', target: 'p_1', text: 'a\tb' } as any, ast);
    const xml = nodeXml(ast);
    // The space and the tab sit in adjacent del/ins runs.
    expect(xml).toMatch(/<w:del\b[\s\S]*?<\/w:del><w:ins\b/);
    // The tab run is preserved so print/PDF keeps the separation.
    expect(xml).toContain('xml:space="preserve">\t</w:t>');
    expect(visibleText(xml)).toBe('a\tb');
  });

  it('plain space boundaries keep preserve markers and exact visible text', () => {
    const ast = makeReplaceAst('hello world');
    executeReplace({ action: 'replace', target: 'p_1', text: 'hello brave world' } as any, ast);
    const xml = nodeXml(ast);
    expect(visibleText(xml)).toBe('hello brave world');
    const doc = parseXml(xml);
    for (const tag of ['w:t', 'w:delText']) {
      const els = doc.getElementsByTagName(tag);
      for (let i = 0; i < els.length; i++) {
        const text = (els[i] as Element).textContent || '';
        if (text.startsWith(' ') || text.endsWith(' ')) {
          expect((els[i] as Element).getAttribute('xml:space')).toBe('preserve');
        }
      }
    }
  });

  it('shared-whitespace del/ins joints gain no extra space', () => {
    const ast = makeReplaceAst('a b');
    executeReplace({ action: 'replace', target: 'p_1', text: 'a\tb' } as any, ast);
    const doc = parseXml(nodeXml(ast));
    const delTexts = doc.getElementsByTagName('w:delText');
    expect(delTexts.length).toBe(1);
    expect((delTexts[0] as Element).textContent).toBe(' ');
  });
});

describe('tracked replace keeps the original text exact', () => {
  function visibleText(xml: string): string {
    const doc = parseXml(xml);
    let out = '';
    const walk = (n: Node): void => {
      const el = n as Element;
      if (n.nodeType === 1 && el.tagName === 'w:del') {
        return;
      }
      if (n.nodeType === 1 && el.tagName === 'w:t') {
        out += el.textContent || '';
        return;
      }
      if (!n.childNodes) return;
      for (let i = 0; i < n.childNodes.length; i++) walk(n.childNodes[i]);
    };
    walk(doc.documentElement);
    return out;
  }

  function originalText(xml: string): string {
    const doc = parseXml(xml);
    let out = '';
    const walk = (n: Node): void => {
      const el = n as Element;
      if (n.nodeType === 1 && el.tagName === 'w:ins') {
        return;
      }
      if (n.nodeType === 1 && (el.tagName === 'w:t' || el.tagName === 'w:delText')) {
        out += el.textContent || '';
        return;
      }
      if (!n.childNodes) return;
      for (let i = 0; i < n.childNodes.length; i++) walk(n.childNodes[i]);
    };
    walk(doc.documentElement);
    return out;
  }

  function delTexts(xml: string): string[] {
    const doc = parseXml(xml);
    const els = doc.getElementsByTagName('w:delText');
    const out: string[] = [];
    for (let i = 0; i < els.length; i++) out.push((els[i] as Element).textContent || '');
    return out;
  }

  it('leading substitution keeps original text exact', () => {
    const ast = makeReplaceAst('New Document');
    executeReplace({ action: 'replace', target: 'p_1', text: 'Tracking Document' } as any, ast);
    const xml = nodeXml(ast);
    expect(visibleText(xml)).toBe('Tracking Document');
    expect(originalText(xml)).toBe('New Document');
    expect(delTexts(xml).join('')).toBe('New');
  });

  it('prefix substitution keeps original text exact', () => {
    const ast = makeReplaceAst('Document stuff');
    executeReplace({ action: 'replace', target: 'p_1', text: 'defaults stuff' } as any, ast);
    const xml = nodeXml(ast);
    expect(visibleText(xml)).toBe('defaults stuff');
    expect(originalText(xml)).toBe('Document stuff');
    expect(delTexts(xml).join('')).toBe('Document');
  });

  it('inner substitution keeps original text exact', () => {
    const ast = makeReplaceAst('New Old Document');
    executeReplace({ action: 'replace', target: 'p_1', text: 'New Tracking Document' } as any, ast);
    const xml = nodeXml(ast);
    expect(visibleText(xml)).toBe('New Tracking Document');
    expect(originalText(xml)).toBe('New Old Document');
    expect(delTexts(xml).join('')).toBe('Old');
  });

  it('comma-delimited substitution keeps original text exact without phantom padding', () => {
    const ast = makeReplaceAst('environment, deliberating');
    executeReplace({ action: 'replace', target: 'p_1', text: 'surroundings, deliberating' } as any, ast);
    const xml = nodeXml(ast);
    expect(visibleText(xml)).toBe('surroundings, deliberating');
    expect(originalText(xml)).toBe('environment, deliberating');
    expect(delTexts(xml).join('')).toBe('environment,');
  });

  it('table cell tracked replace keeps original text exact', () => {
    const tc = new DOMParser().parseFromString(
      `<w:tc xmlns:w="${W}"><w:p><w:r><w:t xml:space="preserve">cell old text</w:t></w:r></w:p></w:tc>`,
      'text/xml',
    ).documentElement as unknown as Element;
    let nextId = 1;
    emitTrackedCellReplace(tc, 'cell old text', 'cell new text', 'ASTra', '2026-04-18T00:00:00Z', {
      getNextId: () => nextId++,
    });
    const xml = serializer.serializeToString(tc);
    expect(visibleText(xml)).toBe('cell new text');
    expect(originalText(xml)).toBe('cell old text');
    expect(delTexts(xml).join('')).toBe('old');
  });
});
