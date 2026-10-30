import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';
import {
  CHART_NS,
  createBlankDocxBuffer,
  createChartElement,
  createDrawingElement,
  createTextElement,
  createWordElement,
  deterministicCommentIso,
  deterministicDate,
  deterministicRevisionIso,
  DocumentManager,
  DocxParser,
  DRAWING_NS,
  executeComment,
  executeDelete,
  executeReplace,
  executeTable,
  executeTrackedChange,
  OperationEngine,
  PPR_CHILD_ORDER,
  resolveDeterministicAuthor,
  sortChildrenBySchema,
  validateDocxBuffer,
  WORD_NS,
} from '@attentionform/astra-engine';

import path from 'path';
import os from 'os';
import fs from 'fs';

describe('Track A: Namespace & Schema Correctness Suite', () => {
  it('A1: createWordElement creates elements with WORD_NS', () => {
    const doc = new DOMParser().parseFromString('<root/>', 'text/xml');
    const p = createWordElement(doc, 'p');
    expect(p.namespaceURI).toBe(WORD_NS);
    expect(p.tagName).toBe('w:p');

    const withPrefix = createWordElement(doc, 'w:r');
    expect(withPrefix.namespaceURI).toBe(WORD_NS);
    expect(withPrefix.tagName).toBe('w:r');
  });

  it('A1: createDrawingElement and createChartElement create correct namespaces', () => {
    const doc = new DOMParser().parseFromString('<root/>', 'text/xml');
    const blip = createDrawingElement(doc, 'blip');
    expect(blip.namespaceURI).toBe(DRAWING_NS);
    expect(blip.tagName).toBe('a:blip');

    const chart = createChartElement(doc, 'plotArea');
    expect(chart.namespaceURI).toBe(CHART_NS);
    expect(chart.tagName).toBe('c:plotArea');
  });

  it('A2: table and insert operations use WORD_NS for all created elements', async () => {
    const blankBuf = createBlankDocxBuffer('blank');
    const parser = new DocxParser(blankBuf);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const firstNode = ast.sections[0].nodes[0];

    // Insert a new paragraph
    await engine.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'paragraph',
          anchor: firstNode.id,
          position: 'after',
          text: 'Inserted paragraph via ns'
        }
      ]
    });

    const updatedAst = await parser.getAST();
    const insertedNode = updatedAst.sections[0].nodes[1];
    expect(insertedNode).toBeDefined();
    expect(insertedNode.xml?.namespaceURI).toBe(WORD_NS);

    // Insert a table
    await engine.executeBatch({
      operations: [
        {
          action: 'insert',
          type: 'table',
          anchor: insertedNode.id,
          position: 'after',
          headers: ['H1', 'H2'],
          rows: [['R1C1', 'R1C2']]
        }
      ]
    });

    const astWithTbl = await parser.getAST();
    const tblNode = astWithTbl.sections[0].nodes.find(n => n.type === 'table');
    expect(tblNode).toBeDefined();
    expect(tblNode!.xml?.namespaceURI).toBe(WORD_NS);

    // Run table operation: addRow, setColumnWidths, setHeaderRepeat
    await engine.executeBatch({
      operations: [
        {
          action: 'table',
          target: tblNode!.id,
          params: { op: 'addRow', afterIndex: 1 }
        },
        {
          action: 'table',
          target: tblNode!.id,
          params: { op: 'setHeaderRepeat', rowIndex: 1 }
        }
      ]
    });

    const docXml = parser.readDocumentXml();
    expect(docXml).toContain('w:tblHeader');
  });

  it('A3: sortChildrenBySchema preserves vendor tags and comments without 9999 exile', () => {
    const xml = `<w:pPr xmlns:w="${WORD_NS}" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
      <!-- comment before -->
      <w:jc w:val="center"/>
      <w14:paraId w:val="12345678"/>
      <!-- comment inside -->
      <w:pStyle w:val="Heading1"/>
      <w:spacing w:after="200"/>
    </w:pPr>`;

    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const pPr = doc.documentElement;
    sortChildrenBySchema(pPr, PPR_CHILD_ORDER);

    const childNodes = Array.from(pPr.childNodes);
    // Comments must still exist in childNodes
    const comments = childNodes.filter(n => n.nodeType === 8);
    expect(comments.length).toBe(2);
    expect(comments[0].textContent).toContain('comment before');
    expect(comments[1].textContent).toContain('comment inside');

    // w14:paraId must NOT be tossed to the end or removed
    const elementTagNames = childNodes.filter(n => n.nodeType === 1).map(n => (n as Element).tagName);
    expect(elementTagNames).toContain('w14:paraId');

    // Schema order for recognized standard tags: pStyle must precede spacing which must precede jc
    const pStyleIdx = elementTagNames.indexOf('w:pStyle');
    const spacingIdx = elementTagNames.indexOf('w:spacing');
    const jcIdx = elementTagNames.indexOf('w:jc');

    expect(pStyleIdx).toBeLessThan(spacingIdx);
    expect(spacingIdx).toBeLessThan(jcIdx);
  });

  it('A4: DocumentManager validateDocxBuffer rejects corrupted documents with Error instance and E_VALIDATION', async () => {
    const docManager = new DocumentManager();
    const testDir = path.join(os.tmpdir(), `astra-test-${Date.now()}`);
    fs.mkdirSync(testDir, { recursive: true });
    const testPath = path.join(testDir, 'test.docx');

    await docManager.createNewDocument('blank', testPath);

    // Corrupt the document XML intentionally
    const parser = (docManager as any).parser;
    parser.reloadDocumentXml('<w:document><w:body><w:invalidElement/></w:body></w:document>');

    // saveAsDocument should throw E_VALIDATION with Error prototype
    let caughtErr: any = null;
    try {
      await docManager.saveDocument();
    } catch (err: any) {
      caughtErr = err;
    }
    expect(caughtErr).not.toBeNull();
    expect(caughtErr instanceof Error).toBe(true);
    expect(caughtErr.code).toBe('E_VALIDATION');
    expect(caughtErr.retryable).toBe(false);
    expect(typeof caughtErr.message).toBe('string');
    expect(caughtErr.stack).toBeDefined();

    // Clean up
    fs.rmSync(testDir, { recursive: true, force: true });
  });

  it('A5: createTextElement in operations/base uses WORD_NS', () => {
    const doc = new DOMParser().parseFromString('<root/>', 'text/xml');
    const t = createTextElement(doc, 'Hello world');
    expect(t.namespaceURI).toBe(WORD_NS);
    expect(t.tagName).toBe('w:t');
    expect(t.textContent).toBe('Hello world');
  });

  it('A6: template fidelity for all templates in createBlankDocxBuffer and createNewDocument', () => {
    const templates = ['blank', 'novel', 'report', 'financial', 'academic', 'newsletter', 'api'];
    for (const tpl of templates) {
      const parser = DocxParser.createNewDocument(tpl);
      const summary = parser.getStructuralSummary();
      expect(summary.nodes.length).toBeGreaterThanOrEqual(2);

      // Verify that non-blank templates have distinct, non-blank title/preview
      if (tpl !== 'blank') {
        const text = summary.nodes.map(n => n.text).join(' ');
        expect(text).not.toContain('Blank Document');
      }

      // Verify createBlankDocxBuffer returns valid buffer
      const buf = createBlankDocxBuffer(tpl);
      expect(buf).toBeInstanceOf(Buffer);
      expect(buf.length).toBeGreaterThan(1000);
    }
  });

  it('A7: styles.xml in generated templates strictly orders w:color before w:sz', () => {
    const parser = DocxParser.createNewDocument('blank');
    const stylesXml = (parser as any).zip.getEntry('word/styles.xml')?.getData().toString('utf8');
    expect(stylesXml).toBeDefined();

    const stylesDoc = new DOMParser().parseFromString(stylesXml!, 'text/xml');
    const rPrNodes = Array.from(stylesDoc.getElementsByTagName('w:rPr'));
    expect(rPrNodes.length).toBeGreaterThan(0);

    for (const rPr of rPrNodes) {
      const childElements = Array.from(rPr.childNodes).filter(n => n.nodeType === 1) as Element[];
      const tags = childElements.map(e => e.tagName);
      const colorIdx = tags.indexOf('w:color');
      const szIdx = tags.indexOf('w:sz');
      if (colorIdx !== -1 && szIdx !== -1) {
        expect(colorIdx).toBeLessThan(szIdx);
      }
    }
  });
});

describe('R-7 deterministic revision stamps via SOURCE_DATE_EPOCH', () => {
  const EPOCH_SECS = '1700000000';
  const PINNED_REV = '2023-11-14T22:13:20.000Z';
  const PINNED_COMMENT = '2023-11-14T22:13:20Z';
  const DEFAULT_REV = new Date(Date.UTC(1980, 0, 1, 0, 0, 0)).toISOString();
  const DEFAULT_COMMENT = new Date(Date.UTC(1980, 0, 1, 0, 0, 0)).toISOString().replace(/\.\d{3}Z$/, 'Z');
  let savedEpoch: string | undefined;
  let savedRep: string | undefined;
  let savedNo: string | undefined;

  beforeEach(() => {
    savedEpoch = process.env.SOURCE_DATE_EPOCH;
    savedRep = process.env.ASTRA_REPRODUCIBLE;
    savedNo = process.env.ASTRA_NO_REPRODUCIBLE;
    delete process.env.ASTRA_REPRODUCIBLE;
    delete process.env.ASTRA_NO_REPRODUCIBLE;
  });

  afterEach(() => {
    if (savedEpoch === undefined) delete process.env.SOURCE_DATE_EPOCH;
    else process.env.SOURCE_DATE_EPOCH = savedEpoch;
    if (savedRep === undefined) delete process.env.ASTRA_REPRODUCIBLE;
    else process.env.ASTRA_REPRODUCIBLE = savedRep;
    if (savedNo === undefined) delete process.env.ASTRA_NO_REPRODUCIBLE;
    else process.env.ASTRA_NO_REPRODUCIBLE = savedNo;
  });

  function readDateFromDocumentXml(buf: Buffer): string | null {
    const xml = new AdmZip(buf).readAsText('word/document.xml');
    return xml.match(/w:date="([^"]+)"/)?.[1] ?? null;
  }

  function readDateFromCommentsXml(buf: Buffer): string | null {
    const xml = new AdmZip(buf).readAsText('word/comments.xml');
    return xml.match(/w:date="([^"]+)"/)?.[1] ?? null;
  }

  it('writes the current time when nothing asks for pinning', async () => {
    delete process.env.SOURCE_DATE_EPOCH;
    const before = Date.now();
    const rev = Date.parse(deterministicRevisionIso());
    const comment = Date.parse(deterministicCommentIso());
    const after = Date.now();
    expect(rev).toBeGreaterThanOrEqual(before);
    expect(rev).toBeLessThanOrEqual(after);
    expect(comment).toBeGreaterThanOrEqual(before - 1000);
    expect(comment).toBeLessThanOrEqual(after);

    const base = createBlankDocxBuffer('blank');
    const parser = new DocxParser(Buffer.from(base));
    const ast: any = await parser.getAST();
    executeDelete({ action: 'delete', target: ast.sections[0].nodes[0].id, trackChanges: true } as any, ast, parser);
    const stamp = readDateFromDocumentXml(parser.rezip(ast));
    expect(stamp).not.toBe(DEFAULT_REV);
    expect(Math.abs(Date.now() - Date.parse(stamp!))).toBeLessThan(60000);
  });

  it('a valid SOURCE_DATE_EPOCH alone turns pinning on; a bad one alone does not', () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    expect(deterministicRevisionIso()).toBe(PINNED_REV);
    for (const bad of ['', '   ', 'not-a-number', '-5']) {
      process.env.SOURCE_DATE_EPOCH = bad;
      expect(deterministicRevisionIso()).not.toBe(DEFAULT_REV);
      expect(Math.abs(Date.now() - Date.parse(deterministicRevisionIso()))).toBeLessThan(60000);
    }
  });

  it('ASTRA_NO_REPRODUCIBLE wins over a SOURCE_DATE_EPOCH', () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    expect(deterministicRevisionIso()).not.toBe(PINNED_REV);
  });

  it('helper pins dates when env set, 1980 when pinned without an epoch, wall-clock on opt-out', () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    expect(deterministicDate().getTime()).toBe(1700000000 * 1000);
    expect(deterministicRevisionIso()).toBe(PINNED_REV);
    expect(deterministicCommentIso()).toBe(PINNED_COMMENT);
    expect(resolveDeterministicAuthor(undefined, undefined)).toBe('ASTra');
    expect(resolveDeterministicAuthor('Jane', 'ASTra')).toBe('Jane');
    expect(resolveDeterministicAuthor('', 'ASTra')).toBe('ASTra');

    // Pinned but with no epoch given: the fixed 1980 stamp.
    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_REPRODUCIBLE = '1';
    expect(deterministicRevisionIso()).toBe(DEFAULT_REV);
    expect(deterministicCommentIso()).toBe(DEFAULT_COMMENT);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const before = Date.now();
    const iso = deterministicRevisionIso();
    const after = Date.now();
    const parsed = Date.parse(iso);
    expect(parsed).toBeGreaterThanOrEqual(before - 5000);
    expect(parsed).toBeLessThanOrEqual(after + 5000);
    expect(iso).not.toBe(PINNED_REV);
    expect(iso).not.toBe(DEFAULT_REV);
    expect(deterministicCommentIso()).not.toBe(PINNED_COMMENT);
    expect(deterministicCommentIso()).not.toBe(DEFAULT_COMMENT);
  });

  it('invalid env values fall back to the 1980 stamp while pinned (wall-clock only on opt-out)', () => {
    process.env.ASTRA_REPRODUCIBLE = '1';
    for (const bad of ['', '   ', 'not-a-number', '-5', 'NaN', 'Infinity']) {
      process.env.SOURCE_DATE_EPOCH = bad;
      expect(deterministicRevisionIso()).toBe(DEFAULT_REV);
      expect(deterministicRevisionIso()).not.toBe(PINNED_REV);
    }
    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const parsed = Date.parse(deterministicRevisionIso());
    const now = Date.now();
    expect(Math.abs(now - parsed)).toBeLessThan(60000);
  });

  it('same input+ops twice with env set yields byte-identical docx (tracked delete)', async () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    const base = createBlankDocxBuffer('blank');
    async function runOnce(): Promise<Buffer> {
      const parser = new DocxParser(Buffer.from(base));
      const ast: any = await parser.getAST();
      const target = ast.sections[0].nodes[0];
      executeDelete({ action: 'delete', target: target.id, trackChanges: true } as any, ast, parser);
      return parser.rezip(ast);
    }
    const a = await runOnce();
    const b = await runOnce();
    expect(Buffer.compare(a, b)).toBe(0);
    expect(readDateFromDocumentXml(a)).toBe(PINNED_REV);
    expect(readDateFromDocumentXml(b)).toBe(PINNED_REV);
    expect((await validateDocxBuffer(a)).ok).toBe(true);
    expect((await validateDocxBuffer(b)).ok).toBe(true);
  }, 60000);

  it('pinned without an epoch is the 1980 stamp; opt-out restores wall-clock', async () => {
    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_REPRODUCIBLE = '1';

    const base = createBlankDocxBuffer('blank');
    const parser = new DocxParser(Buffer.from(base));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    executeDelete({ action: 'delete', target: target.id, trackChanges: true } as any, ast, parser);
    const buf = parser.rezip(ast);
    expect(readDateFromDocumentXml(buf)).toBe(DEFAULT_REV);
    expect((await validateDocxBuffer(buf)).ok).toBe(true);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const parser2 = new DocxParser(Buffer.from(base));
    const ast2: any = await parser2.getAST();
    const target2 = ast2.sections[0].nodes[0];
    const before = Date.now();
    executeDelete({ action: 'delete', target: target2.id, trackChanges: true } as any, ast2, parser2);
    const after = Date.now();
    const buf2 = parser2.rezip(ast2);
    const stamp = readDateFromDocumentXml(buf2);
    expect(stamp).not.toBeNull();
    expect(stamp).not.toBe(PINNED_REV);
    expect(stamp).not.toBe(DEFAULT_REV);
    const parsed = Date.parse(stamp!);
    expect(parsed).toBeGreaterThanOrEqual(before - 5000);
    expect(parsed).toBeLessThanOrEqual(after + 5000);
    expect((await validateDocxBuffer(buf2)).ok).toBe(true);
  }, 60000);

  it('comment w:date is pinned with env, 1980 when pinned without an epoch, wall-clock on opt-out', async () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    const base = createBlankDocxBuffer('blank');
    async function commentOnce(text: string): Promise<Buffer> {
      const parser = new DocxParser(Buffer.from(base));
      const ast: any = await parser.getAST();
      const target = ast.sections[0].nodes[0];
      executeComment({ action: 'comment', operation: 'insert', target: target.id, text } as any, ast, parser);
      return parser.rezip(ast);
    }
    const pinned1 = await commentOnce('hello pinned');
    const pinned2 = await commentOnce('hello pinned');
    expect(readDateFromCommentsXml(pinned1)).toBe(PINNED_COMMENT);
    expect(readDateFromCommentsXml(pinned2)).toBe(PINNED_COMMENT);
    expect(Buffer.compare(pinned1, pinned2)).toBe(0);
    expect((await validateDocxBuffer(pinned1)).ok).toBe(true);

    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_REPRODUCIBLE = '1';
    const defaulted = await commentOnce('hello default');
    expect(readDateFromCommentsXml(defaulted)).toBe(DEFAULT_COMMENT);
    expect((await validateDocxBuffer(defaulted)).ok).toBe(true);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const before = Date.now();
    const unpinned = await commentOnce('hello wall-clock');
    const after = Date.now();
    const stamp = readDateFromCommentsXml(unpinned);
    expect(stamp).not.toBeNull();
    expect(stamp).not.toBe(PINNED_COMMENT);
    expect(stamp).not.toBe(DEFAULT_COMMENT);
    const parsed = Date.parse(stamp!);
    expect(parsed).toBeGreaterThanOrEqual(before - 5000);
    expect(parsed).toBeLessThanOrEqual(after + 5000);
    expect((await validateDocxBuffer(unpinned)).ok).toBe(true);
  }, 60000);

  it('replace op twice with env set yields byte-identical docx', async () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    const base = createBlankDocxBuffer('blank');
    async function runOnce(): Promise<Buffer> {
      const parser = new DocxParser(Buffer.from(base));
      const ast: any = await parser.getAST();
      const target = ast.sections[0].nodes[0];
      executeReplace({ action: 'replace', target: target.id, text: 'pinned replacement', trackChanges: true } as any, ast);
      return parser.rezip(ast);
    }
    const a = await runOnce();
    const b = await runOnce();
    expect(Buffer.compare(a, b)).toBe(0);
    expect(readDateFromDocumentXml(a)).toBe(PINNED_REV);
    expect(readDateFromDocumentXml(b)).toBe(PINNED_REV);
    expect((await validateDocxBuffer(a)).ok).toBe(true);
    expect((await validateDocxBuffer(b)).ok).toBe(true);
  }, 60000);

  it('replace pinned without an epoch is 1980; opt-out restores wall-clock; invalid env stays pinned', async () => {
    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_REPRODUCIBLE = '1';

    const base = createBlankDocxBuffer('blank');
    async function replaceOnce(text: string): Promise<{ buf: Buffer; before: number; after: number }> {
      const parser = new DocxParser(Buffer.from(base));
      const ast: any = await parser.getAST();
      const target = ast.sections[0].nodes[0];
      const before = Date.now();
      executeReplace({ action: 'replace', target: target.id, text, trackChanges: true } as any, ast);
      const after = Date.now();
      return { buf: parser.rezip(ast), before, after };
    }
    const def = await replaceOnce('default replacement');
    expect(readDateFromDocumentXml(def.buf)).toBe(DEFAULT_REV);
    expect((await validateDocxBuffer(def.buf)).ok).toBe(true);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const wall = await replaceOnce('wall-clock replacement');
    const stamp = readDateFromDocumentXml(wall.buf);
    expect(stamp).not.toBeNull();
    expect(stamp).not.toBe(PINNED_REV);
    expect(stamp).not.toBe(DEFAULT_REV);
    const parsed = Date.parse(stamp!);
    expect(parsed).toBeGreaterThanOrEqual(wall.before - 5000);
    expect(parsed).toBeLessThanOrEqual(wall.after + 5000);
    expect((await validateDocxBuffer(wall.buf)).ok).toBe(true);
    delete process.env.ASTRA_NO_REPRODUCIBLE;

    process.env.SOURCE_DATE_EPOCH = 'bogus-value';
    const fallback = await replaceOnce('invalid-env replacement');
    expect(readDateFromDocumentXml(fallback.buf)).toBe(DEFAULT_REV);
    expect((await validateDocxBuffer(fallback.buf)).ok).toBe(true);
  }, 60000);

  it('table updateCell twice with env set yields equal w:date + valid output', async () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    const setup = new DocxParser(createBlankDocxBuffer('blank'));
    const setupAst: any = await setup.getAST();
    const anchor = setupAst.sections[0].nodes[0].id;
    const engine = new OperationEngine(setupAst, setup);
    await engine.executeBatch({
      operations: [
        { action: 'insert', type: 'table', anchor, position: 'after', headers: ['H1', 'H2'], rows: [['R1C1', 'R1C2']] }
      ]
    } as any);
    const base = setup.rezip(setupAst);
    function readAllDates(buf: Buffer): string[] {
      const xml = new AdmZip(buf).readAsText('word/document.xml');
      return [...xml.matchAll(/w:date="([^"]+)"/g)].map(m => m[1]);
    }
    async function runOnce(): Promise<Buffer> {
      const parser = new DocxParser(Buffer.from(base));
      const ast: any = await parser.getAST();
      const tbl = ast.sections[0].nodes.find((n: any) => n.type === 'table');
      await executeTable({ action: 'table', target: tbl.id, trackChanges: true, params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'pinned cell' } } as any, ast);
      return parser.rezip(ast);
    }
    const a = await runOnce();
    const b = await runOnce();
    const datesA = readAllDates(a);
    const datesB = readAllDates(b);
    expect(datesA.length).toBeGreaterThan(0);
    expect(datesA).toEqual(datesB);
    for (const d of datesA) expect(d).toBe(PINNED_REV);
    expect((await validateDocxBuffer(a)).ok).toBe(true);
    expect((await validateDocxBuffer(b)).ok).toBe(true);

    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_REPRODUCIBLE = '1';
    const defParser = new DocxParser(Buffer.from(base));
    const defAst: any = await defParser.getAST();
    const defTbl = defAst.sections[0].nodes.find((n: any) => n.type === 'table');
    await executeTable({ action: 'table', target: defTbl.id, trackChanges: true, params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'default cell' } } as any, defAst);
    const defBuf = defParser.rezip(defAst);
    for (const s of readAllDates(defBuf)) expect(s).toBe(DEFAULT_REV);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const parser = new DocxParser(Buffer.from(base));
    const ast: any = await parser.getAST();
    const tbl = ast.sections[0].nodes.find((n: any) => n.type === 'table');
    const before = Date.now();
    await executeTable({ action: 'table', target: tbl.id, trackChanges: true, params: { op: 'updateCell', rowIndex: 1, colIndex: 1, text: 'wall-clock cell' } } as any, ast);
    const after = Date.now();
    const buf = parser.rezip(ast);
    const stamps = readAllDates(buf);
    expect(stamps.length).toBeGreaterThan(0);
    for (const s of stamps) {
      expect(s).not.toBe(PINNED_REV);
      expect(s).not.toBe(DEFAULT_REV);
      const p = Date.parse(s);
      expect(p).toBeGreaterThanOrEqual(before - 5000);
      expect(p).toBeLessThanOrEqual(after + 5000);
    }
    expect((await validateDocxBuffer(buf)).ok).toBe(true);
  }, 60000);

  it('trackedChange insert honors env pin, explicit date, 1980 when pinned without an epoch, and opt-out wall-clock', async () => {
    process.env.SOURCE_DATE_EPOCH = EPOCH_SECS;
    const base = createBlankDocxBuffer('blank');
    async function insertOnce(text: string, date?: string): Promise<Buffer> {
      const parser = new DocxParser(Buffer.from(base));
      const ast: any = await parser.getAST();
      const target = ast.sections[0].nodes[0];
      executeTrackedChange({ action: 'trackedChange', operation: 'insert', target: target.id, text, date } as any, ast, parser);
      return parser.rezip(ast);
    }
    const pinned = await insertOnce('pinned tc');
    expect(readDateFromDocumentXml(pinned)).toBe(PINNED_REV);
    expect((await validateDocxBuffer(pinned)).ok).toBe(true);

    const explicit = '2020-01-02T03:04:05.000Z';
    const explicitBuf = await insertOnce('explicit date tc', explicit);
    expect(readDateFromDocumentXml(explicitBuf)).toBe(explicit);
    expect((await validateDocxBuffer(explicitBuf)).ok).toBe(true);

    delete process.env.SOURCE_DATE_EPOCH;
    process.env.ASTRA_REPRODUCIBLE = '1';
    const defBuf = await insertOnce('default tc');
    expect(readDateFromDocumentXml(defBuf)).toBe(DEFAULT_REV);

    process.env.ASTRA_NO_REPRODUCIBLE = '1';
    const parser = new DocxParser(Buffer.from(base));
    const ast: any = await parser.getAST();
    const target = ast.sections[0].nodes[0];
    const before = Date.now();
    executeTrackedChange({ action: 'trackedChange', operation: 'insert', target: target.id, text: 'wall-clock tc' } as any, ast, parser);
    const after = Date.now();
    const buf = parser.rezip(ast);
    const stamp = readDateFromDocumentXml(buf);
    expect(stamp).not.toBeNull();
    expect(stamp).not.toBe(PINNED_REV);
    expect(stamp).not.toBe(DEFAULT_REV);
    const parsed = Date.parse(stamp!);
    expect(parsed).toBeGreaterThanOrEqual(before - 5000);
    expect(parsed).toBeLessThanOrEqual(after + 5000);
    expect((await validateDocxBuffer(buf)).ok).toBe(true);
  }, 60000);
});
