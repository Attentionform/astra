import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import fs from 'fs';
import fsp from 'fs/promises';
import os from 'os';
import path from 'path';
import { createHash } from 'crypto';
import {
  initAstraRepo,
  commitChange,
  getHistory,
  undo,
  undoPackage,
  checkoutCommit,
  isTrackedPackageFile,
  DEFAULT_COMMIT_AUTHOR,
  DocxParser,
  OperationEngine,
  validateDocxBuffer,
  validateDocxFile,
  DocumentManager,
  registerTrustedDocPath,
} from '@attentionform/astra-engine';

const SANDBOX_BASE = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-engine-real-test-'));
const TEST_GALLERY_DIR = path.resolve(__dirname, '../../test-gallery');

function sha256(content: Buffer | string): string {
  return createHash('sha256').update(content).digest('hex');
}

function collectExtras(parser: DocxParser): Record<string, Buffer> {
  const snap = parser.capturePackageSnapshot();
  const extras: Record<string, Buffer> = {};
  for (const [name, buf] of snap) {
    if (isTrackedPackageFile(name)) {
      extras[name] = buf;
    }
  }
  return extras;
}

function reinjectAndRebuild(parser: DocxParser, xml: string, extra: Map<string, Buffer>): Buffer {
  const merged = parser.capturePackageSnapshot();
  for (const key of [...merged.keys()]) {
    if (key !== 'word/document.xml' && isTrackedPackageFile(key)) {
      merged.delete(key);
    }
  }
  for (const [k, v] of extra) {
    merged.set(k, v);
  }
  merged.set('word/document.xml', Buffer.from(xml, 'utf8'));
  parser.restorePackageSnapshot(merged);
  return parser.rezip();
}

describe('Engine Real-World Git Versioning & Revert Verification', () => {
  beforeAll(async () => {
    await fsp.mkdir(SANDBOX_BASE, { recursive: true });
  });

  afterAll(async () => {
    await fsp.rm(SANDBOX_BASE, { recursive: true, force: true });
  });

  describe('1. Lifecycle Execution on Real Doc: T1_coffee_onepager.docx (Charts + Footers)', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case1-t1');
    const docPath = path.join(caseDir, 'T1_coffee_onepager.docx');

    let stage0Xml: string;
    let stage0DocHash: string;
    let stage0StylesHash: string;
    let stage0ChartHash: string;
    let stage0RelsHash: string;
    let stage0CommitOid: string;
    let stage1CommitOid: string;
    let stage2CommitOid: string;
    let stage3CommitOid: string;

    it('Stage 0 (Baseline): Load real doc, initialize git repo, and record baseline hashes', async () => {
      await fsp.mkdir(caseDir, { recursive: true });
      await fsp.copyFile(path.join(TEST_GALLERY_DIR, 'T1_coffee_onepager.docx'), docPath);
      registerTrustedDocPath(docPath, 'dialog');

      const initialBuf = await fsp.readFile(docPath);
      const parser = new DocxParser(initialBuf);
      const snap = parser.capturePackageSnapshot();

      stage0Xml = parser.readDocumentXml();
      stage0DocHash = sha256(stage0Xml);
      expect(stage0Xml).toContain('w:document');

      // Verify baseline parts
      expect(snap.has('word/styles.xml')).toBe(true);
      stage0StylesHash = sha256(snap.get('word/styles.xml')!);

      expect(snap.has('word/charts/chart1.xml')).toBe(true);
      stage0ChartHash = sha256(snap.get('word/charts/chart1.xml')!);

      expect(snap.has('word/_rels/document.xml.rels')).toBe(true);
      stage0RelsHash = sha256(snap.get('word/_rels/document.xml.rels')!);

      const stage0Extras = collectExtras(parser);
      expect(Object.keys(stage0Extras)).toContain('word/styles.xml');
      expect(Object.keys(stage0Extras)).toContain('word/charts/chart1.xml');
      expect(Object.keys(stage0Extras)).toContain('word/_rels/document.xml.rels');
      expect(Object.keys(stage0Extras)).toContain('[Content_Types].xml');

      const repoDir = await initAstraRepo(docPath, stage0Xml, stage0Extras);
      expect(fs.existsSync(repoDir)).toBe(true);

      const history = await getHistory(docPath);
      expect(history.length).toBe(1);
      expect(history[0].message).toBe('Initial state\n');
      expect(history[0].author).toBe('ASTra');
      expect(history[0].oid).toMatch(/^[0-9a-f]{40}$/);
      stage0CommitOid = history[0].oid;
    }, 30_000);

    it('Stage 1: Apply batch mutations (replace text, style heading, insert paragraph), save & commit', async () => {
      const buf = await fsp.readFile(docPath);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const firstNode = ast.sections[0].nodes[0];
      const lastNode = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];

      const batchRes = await engine.executeBatch({
        operations: [
          {
            action: 'replace',
            target: firstNode.id,
            text: 'Specialty Coffee Global Performance 2026',
          } as never,
          {
            action: 'style',
            target: firstNode.id,
            properties: {
              bold: true,
              fontSize: 18,
              color: '1F497D',
            },
          } as never,
          {
            action: 'insert',
            type: 'paragraph',
            anchor: lastNode.id,
            position: 'after',
            text: 'Stage 1 Innovation: Direct-trade sourcing improved net operating margins by 14.8%.',
          } as never,
        ],
        sequential: true,
      });

      expect(batchRes.success).toBe(true);
      expect(batchRes.appliedCount).toBe(3);

      const updatedBuf = parser.rezip();
      await fsp.writeFile(docPath, updatedBuf);

      const stage1Xml = parser.readDocumentXml();
      expect(stage1Xml).toContain('Specialty Coffee Global Performance 2026');
      expect(stage1Xml).toContain('Stage 1 Innovation: Direct-trade sourcing');

      stage1CommitOid = await commitChange(
        docPath,
        stage1Xml,
        'Stage 1: Batch headline edits, styling, and innovation paragraph',
        collectExtras(parser),
        undefined,
        'QA Automation Lead'
      );

      expect(stage1CommitOid).toMatch(/^[0-9a-f]{40}$/);
      const history = await getHistory(docPath);
      expect(history.length).toBe(2);
      expect(history[0].oid).toBe(stage1CommitOid);
      expect(history[0].author).toBe('QA Automation Lead');
      expect(history[0].message).toBe('Stage 1: Batch headline edits, styling, and innovation paragraph\n');

      const valRes = await validateDocxBuffer(updatedBuf);
      expect(valRes.ok).toBe(true);
      expect(valRes.errors.length).toBe(0);
    }, 30_000);

    it('Stage 2: Add formatted table with headers, rows, and cell shading, save & commit', async () => {
      const buf = await fsp.readFile(docPath);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const lastNode = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];

      const tableRes = await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'table',
            anchor: lastNode.id,
            position: 'after',
            headers: ['Segment', 'Q1 Revenue', 'Gross Margin', 'YoY Growth'],
            rows: [
              ['Single-Origin Whole Bean', '$4,850,000', '64.2%', '+22.1%'],
              ['Cold Brew Concentrate', '$3,120,000', '71.5%', '+31.4%'],
              ['Nitro Draft Kegs', '$1,980,000', '58.9%', '+12.7%'],
            ],
          } as never,
        ],
      });
      expect(tableRes.success).toBe(true);

      const tableNode = ast.sections[0].nodes.find((n) => n.type === 'table');
      expect(tableNode).toBeDefined();

      const shadingRes = await engine.executeBatch({
        operations: [
          {
            action: 'table',
            target: tableNode!.id,
            operation: 'setCellShading',
            rowIndex: 1,
            colIndex: 1,
            fill: '003366',
          } as never,
          {
            action: 'table',
            target: tableNode!.id,
            operation: 'setCellShading',
            rowIndex: 1,
            colIndex: 2,
            fill: '003366',
          } as never,
          {
            action: 'table',
            target: tableNode!.id,
            operation: 'setCellShading',
            rowIndex: 1,
            colIndex: 3,
            fill: '003366',
          } as never,
          {
            action: 'table',
            target: tableNode!.id,
            operation: 'setCellShading',
            rowIndex: 1,
            colIndex: 4,
            fill: '003366',
          } as never,
        ],
      });
      expect(shadingRes.success).toBe(true);

      const updatedBuf = parser.rezip();
      await fsp.writeFile(docPath, updatedBuf);

      const stage2Xml = parser.readDocumentXml();
      expect(stage2Xml).toContain('Single-Origin Whole Bean');
      expect(stage2Xml).toContain('Cold Brew Concentrate');
      expect(stage2Xml).toContain('003366');

      stage2CommitOid = await commitChange(
        docPath,
        stage2Xml,
        'Stage 2: Added quarterly financial performance table with cell shading',
        collectExtras(parser),
        undefined,
        'Operations Lead'
      );

      const history = await getHistory(docPath);
      expect(history.length).toBe(3);
      expect(history[0].oid).toBe(stage2CommitOid);

      const valRes = await validateDocxBuffer(updatedBuf);
      expect(valRes.ok).toBe(true);
    }, 30_000);

    it('Stage 3: Add executive header & footer, save & commit', async () => {
      const buf = await fsp.readFile(docPath);
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const hfRes = await engine.executeBatch({
        operations: [
          {
            action: 'headerFooter',
            operation: 'set',
            targetType: 'header',
            pageType: 'default',
            text: 'CONFIDENTIAL EXECUTIVE BRIEFING - GLOBAL ROASTERY',
          } as never,
          {
            action: 'headerFooter',
            operation: 'set',
            targetType: 'footer',
            pageType: 'default',
            text: 'Internal Distribution Only - Serial 009281',
          } as never,
        ],
      });
      expect(hfRes.success).toBe(true);

      const updatedBuf = parser.rezip();
      await fsp.writeFile(docPath, updatedBuf);

      const stage3Xml = parser.readDocumentXml();
      const extras = collectExtras(parser);
      expect(Object.keys(extras)).toEqual(expect.arrayContaining(['word/header1.xml']));

      stage3CommitOid = await commitChange(
        docPath,
        stage3Xml,
        'Stage 3: Injected executive header and custom footer markings',
        extras,
        undefined,
        'Executive Secretary'
      );

      const history = await getHistory(docPath);
      expect(history.length).toBe(4);
      expect(history[0].oid).toBe(stage3CommitOid);

      const valRes = await validateDocxBuffer(updatedBuf);
      expect(valRes.ok).toBe(true);
    }, 30_000);

    it('Revert 1 Step: undoPackage() rewinds branch pointer, Stage 3 vanishes, Stage 2 table remains 100%', async () => {
      const undoRes = await undoPackage(docPath);

      // Verify returned XML does NOT contain Stage 3 headers
      expect(undoRes.xml).not.toContain('CONFIDENTIAL EXECUTIVE BRIEFING');

      // Verify returned XML DOES contain Stage 2 table and Stage 1 text
      expect(undoRes.xml).toContain('Single-Origin Whole Bean');
      expect(undoRes.xml).toContain('Specialty Coffee Global Performance 2026');

      // Verify git history was rewound to length 3
      const history = await getHistory(docPath);
      expect(history.length).toBe(3);
      expect(history[0].oid).toBe(stage2CommitOid);

      // Rebuild package zip from undo result
      const liveBuf = await fsp.readFile(docPath);
      const liveParser = new DocxParser(liveBuf);
      const revertedZipBuf = reinjectAndRebuild(liveParser, undoRes.xml, undoRes.extra);
      await fsp.writeFile(docPath, revertedZipBuf);

      // Extract and verify cleanly
      const verifiedParser = new DocxParser(revertedZipBuf);
      const revertedSnap = verifiedParser.capturePackageSnapshot();
      const currentDocXml = verifiedParser.readDocumentXml();

      expect(currentDocXml).not.toContain('CONFIDENTIAL EXECUTIVE BRIEFING');
      expect(currentDocXml).toContain('Single-Origin Whole Bean');
      expect(currentDocXml).toContain('Specialty Coffee Global Performance 2026');
      expect(currentDocXml).toContain('003366');

      // Verify chart from Stage 0 is still intact
      expect(revertedSnap.has('word/charts/chart1.xml')).toBe(true);
      expect(sha256(revertedSnap.get('word/charts/chart1.xml')!)).toBe(stage0ChartHash);

      // ECMA-376 schema validation
      const valRes = await validateDocxBuffer(revertedZipBuf);
      expect(valRes.ok).toBe(true);
      expect(valRes.errors.length).toBe(0);
    }, 30_000);

    it('Redo / Revert by Commit OID: checkoutCommit(stage0CommitOid) restores baseline document.xml and AST', async () => {
      const checkoutRes = await checkoutCommit(docPath, stage0CommitOid);

      // Verify returned XML matches Stage 0 baseline hash
      const checkedXmlHash = sha256(checkoutRes.xml);
      expect(checkedXmlHash).toBe(stage0DocHash);

      // Re-inject into package snapshot and rebuild
      const currentBuf = await fsp.readFile(docPath);
      const parser = new DocxParser(currentBuf);
      const stage0RebuiltBuf = reinjectAndRebuild(parser, checkoutRes.xml, checkoutRes.extra);
      await fsp.writeFile(docPath, stage0RebuiltBuf);

      // Verify reconstructed docx
      const restoredParser = new DocxParser(stage0RebuiltBuf);
      const restoredXml = restoredParser.readDocumentXml();
      expect(sha256(restoredXml)).toBe(stage0DocHash);

      // Verify AST matches baseline structure
      const baselineParser = new DocxParser(await fsp.readFile(path.join(TEST_GALLERY_DIR, 'T1_coffee_onepager.docx')));
      const baselineAst = await baselineParser.getAST();
      const restoredAst = await restoredParser.getAST();

      expect(restoredAst.sections.length).toBe(baselineAst.sections.length);
      expect(restoredAst.sections[0].nodes.length).toBe(baselineAst.sections[0].nodes.length);
      expect(restoredAst.sections[0].nodes.map((n) => n.metadata.text)).toEqual(
        baselineAst.sections[0].nodes.map((n) => n.metadata.text)
      );

      // Verify chart and styles SHA-256 match Stage 0 baseline
      const restoredSnap = restoredParser.capturePackageSnapshot();
      expect(sha256(restoredSnap.get('word/styles.xml')!)).toBe(stage0StylesHash);
      expect(sha256(restoredSnap.get('word/charts/chart1.xml')!)).toBe(stage0ChartHash);
      expect(sha256(restoredSnap.get('word/_rels/document.xml.rels')!)).toBe(stage0RelsHash);

      // OOXML validation check
      const valRes = await validateDocxBuffer(stage0RebuiltBuf);
      expect(valRes.ok).toBe(true);
      expect(valRes.errors.length).toBe(0);
    }, 30_000);
  });

  describe('2. Lifecycle Execution on Real Doc: T4_budget.docx (Financial Sheet)', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case2-t4');
    const docPath = path.join(caseDir, 'T4_budget.docx');

    let stage0DocHash: string;
    let stage0CommitOid: string;

    it('Stage 0 -> Stage 1 -> Stage 2 -> Revert full lifecycle on T4', async () => {
      await fsp.mkdir(caseDir, { recursive: true });
      await fsp.copyFile(path.join(TEST_GALLERY_DIR, 'T4_budget.docx'), docPath);
      registerTrustedDocPath(docPath, 'dialog');

      const initialBuf = await fsp.readFile(docPath);
      const parser = new DocxParser(initialBuf);
      const stage0Xml = parser.readDocumentXml();
      stage0DocHash = sha256(stage0Xml);

      await initAstraRepo(docPath, stage0Xml, collectExtras(parser));
      const h0 = await getHistory(docPath);
      stage0CommitOid = h0[0].oid;

      // Stage 1: Batch edits
      let ast = await parser.getAST();
      let engine = new OperationEngine(ast, parser);
      const lastNode = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];

      await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'paragraph',
            anchor: lastNode.id,
            position: 'after',
            text: 'FY2026 Revised Operating Expenses Target: $14.2M.',
          } as never,
        ],
      });
      const stage1Xml = parser.readDocumentXml();
      await fsp.writeFile(docPath, parser.rezip());
      await commitChange(docPath, stage1Xml, 'Stage 1: Added OpEx target', collectExtras(parser));

      // Stage 2: Formatted table
      ast = await parser.getAST();
      engine = new OperationEngine(ast, parser);
      const updatedLast = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
      await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'table',
            anchor: updatedLast.id,
            position: 'after',
            headers: ['Department', 'Budget', 'Actual', 'Variance'],
            rows: [
              ['Engineering', '$6,500,000', '$6,250,000', '-$250,000'],
              ['Product Design', '$2,100,000', '$2,050,000', '-$50,000'],
            ],
          } as never,
        ],
      });
      const stage2Xml = parser.readDocumentXml();
      await fsp.writeFile(docPath, parser.rezip());
      await commitChange(docPath, stage2Xml, 'Stage 2: Added budget breakdown table', collectExtras(parser));

      const h2 = await getHistory(docPath);
      expect(h2.length).toBe(3);

      // Revert 1 step
      const undoRes = await undoPackage(docPath);
      expect(undoRes.xml).not.toContain('Product Design');
      expect(undoRes.xml).toContain('FY2026 Revised Operating Expenses Target');

      const liveParser = new DocxParser(await fsp.readFile(docPath));
      const revertedBuf = reinjectAndRebuild(liveParser, undoRes.xml, undoRes.extra);
      await fsp.writeFile(docPath, revertedBuf);

      const val1 = await validateDocxBuffer(revertedBuf);
      expect(val1.ok).toBe(true);

      // Checkout Stage 0 by OID
      const coRes = await checkoutCommit(docPath, stage0CommitOid);
      expect(sha256(coRes.xml)).toBe(stage0DocHash);

      const stage0Buf = reinjectAndRebuild(liveParser, coRes.xml, coRes.extra);
      const val0 = await validateDocxBuffer(stage0Buf);
      expect(val0.ok).toBe(true);
    }, 30_000);
  });

  describe('3. Binary Media & OOXML Parts Preservation: T5_image_doc.docx', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case3-t5-media');
    const docPath = path.join(caseDir, 'T5_image_doc.docx');

    it('preserves binary images byte-for-byte across git commits, checkouts, and undos', async () => {
      await fsp.mkdir(caseDir, { recursive: true });
      await fsp.copyFile(path.join(TEST_GALLERY_DIR, 'T5_image_doc.docx'), docPath);
      registerTrustedDocPath(docPath, 'dialog');

      const buf = await fsp.readFile(docPath);
      const parser = new DocxParser(buf);
      const snap = parser.capturePackageSnapshot();

      const imageFiles = [...snap.keys()].filter((f) => f.startsWith('word/media/'));
      expect(imageFiles.length).toBeGreaterThan(0);
      const imagePath = imageFiles[0];
      const origImageBytes = snap.get(imagePath)!;
      const origImageHash = sha256(origImageBytes);

      const extras0 = collectExtras(parser);
      expect(extras0[imagePath]).toBeDefined();
      expect(sha256(extras0[imagePath])).toBe(origImageHash);

      const initialXml = parser.readDocumentXml();
      await initAstraRepo(docPath, initialXml, extras0);

      const h0 = await getHistory(docPath);
      const stage0Oid = h0[0].oid;

      // Make a text mutation
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);
      const lastNode = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];
      await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'paragraph',
            anchor: lastNode.id,
            position: 'after',
            text: 'Caption: Figure 1 High-Resolution Microscopy Scan',
          } as never,
        ],
      });

      const updatedXml = parser.readDocumentXml();
      await fsp.writeFile(docPath, parser.rezip());
      const stage1Oid = await commitChange(docPath, updatedXml, 'Added figure caption', collectExtras(parser));

      // Checkout stage 0 commit by OID
      const co0 = await checkoutCommit(docPath, stage0Oid);
      const coImageBytes = co0.extra.get(imagePath);
      expect(coImageBytes).toBeDefined();
      expect(sha256(coImageBytes!)).toBe(origImageHash);
      expect(coImageBytes!.equals(origImageBytes)).toBe(true);

      // Checkout stage 1 commit by OID
      const co1 = await checkoutCommit(docPath, stage1Oid);
      const co1ImageBytes = co1.extra.get(imagePath);
      expect(co1ImageBytes).toBeDefined();
      expect(sha256(co1ImageBytes!)).toBe(origImageHash);
      expect(co1ImageBytes!.equals(origImageBytes)).toBe(true);

      // Rebuild and validate OOXML
      const restoredBuf = reinjectAndRebuild(parser, co0.xml, co0.extra);
      const valRes = await validateDocxBuffer(restoredBuf);
      expect(valRes.ok).toBe(true);
    });
  });

  describe('4. Rich Multi-Part OOXML Preservation: Charts, Embeddings, Themes, Rels', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case4-rich-parts');
    const docPath = path.join(caseDir, 'rich_doc.docx');

    it('round-trips custom binary embeddings, chart spaces, themes, and rels intact', async () => {
      await fsp.mkdir(caseDir, { recursive: true });

      const v1 = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>V1 Base</w:t></w:r></w:p></w:body></w:document>';
      const v2 = '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>V2 Rich</w:t></w:r></w:p></w:body></w:document>';

      const chartBytes = Buffer.from('<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>', 'utf8');
      const oleBinBytes = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00, 0x5a, 0x99, 0xff]);
      const themeBytes = Buffer.from('<?xml version="1.0" encoding="UTF-8"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"/>', 'utf8');
      const docRelBytes = Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>', 'utf8');
      const headerRelBytes = Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>', 'utf8');

      const richExtras: Record<string, Buffer> = {
        'word/charts/chart1.xml': chartBytes,
        'word/embeddings/oleObject1.bin': oleBinBytes,
        'word/theme/theme1.xml': themeBytes,
        'word/_rels/document.xml.rels': docRelBytes,
        'word/_rels/header1.xml.rels': headerRelBytes,
      };

      await initAstraRepo(docPath, v1);
      const s1Oid = await commitChange(docPath, v2, 'Commit with rich parts', richExtras);

      const history = await getHistory(docPath);
      expect(history.length).toBe(2);

      const headCheckout = await checkoutCommit(docPath, s1Oid);
      expect(headCheckout.xml).toBe(v2);

      for (const [relPath, expectedBuf] of Object.entries(richExtras)) {
        const actualBuf = headCheckout.extra.get(relPath);
        expect(actualBuf, `Part ${relPath} should exist in checked out extra`).toBeDefined();
        expect(actualBuf!.equals(expectedBuf), `Part ${relPath} bytes should match exactly`).toBe(true);
      }
    });
  });

  describe('5. Comprehensive Engine Git Features & Error Gates', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case5-git-features');
    const docPath = path.join(caseDir, 'feature_test.docx');

    it('verifies getHistory metadata: timestamps, messages, author resolution, commit order', async () => {
      await fsp.mkdir(caseDir, { recursive: true });

      const startTime = Date.now() - 2000;
      await initAstraRepo(docPath, '<w:p>Rev 1</w:p>');
      await commitChange(docPath, '<w:p>Rev 2</w:p>', 'Second change', undefined, undefined, 'Specialist Alice');
      await commitChange(docPath, '<w:p>Rev 3</w:p>', 'Third change with default author');

      const history = await getHistory(docPath);
      expect(history.length).toBe(3);

      // Order: newest first
      expect(history[0].message).toBe('Third change with default author\n');
      expect(history[0].author).toBe(DEFAULT_COMMIT_AUTHOR);
      expect(history[0].timestamp.getTime()).toBeGreaterThanOrEqual(startTime);

      expect(history[1].message).toBe('Second change\n');
      expect(history[1].author).toBe('Specialist Alice');

      expect(history[2].message).toBe('Initial state\n');
      expect(history[2].author).toBe('ASTra');

      for (const c of history) {
        expect(c.oid).toMatch(/^[0-9a-f]{40}$/);
        expect(c.timestamp).toBeInstanceOf(Date);
      }
    });

    it('verifies checkoutCommit error handling: unknown OIDs, uninitialized doc paths, invalid OIDs', async () => {
      // Unknown OID on initialized repo
      await expect(
        checkoutCommit(docPath, '11223344556677889900aabbccddeeff11223344')
      ).rejects.toThrow(/not found in history/);

      // Empty or non-string OID
      await expect(checkoutCommit(docPath, '')).rejects.toThrow(/not found in history/);

      // Uninitialized docPath
      const uninitDocPath = path.join(caseDir, 'uninitialized.docx');
      await expect(
        getHistory(uninitDocPath)
      ).rejects.toThrow(/no history exists yet/);

      await expect(
        checkoutCommit(uninitDocPath, '11223344556677889900aabbccddeeff11223344')
      ).rejects.toThrow(/no history exists yet/);

      await expect(
        undoPackage(uninitDocPath)
      ).rejects.toThrow(/no history exists yet/);
    });

    it('verifies undo bounds: attempting to undo at initial state throws domain error', async () => {
      const singleCommitDoc = path.join(caseDir, 'single_commit.docx');
      await initAstraRepo(singleCommitDoc, '<w:p>Only 1</w:p>');

      await expect(undo(singleCommitDoc)).rejects.toThrow('Cannot undo: already at the initial state.');
      await expect(undoPackage(singleCommitDoc)).rejects.toThrow('Cannot undo: already at the initial state.');
    });

    it('verifies concurrent checkouts leave working tree and HEAD completely untouched', async () => {
      const history = await getHistory(docPath);
      const oids = history.map((c) => c.oid);

      const [c0, c1, c2] = await Promise.all([
        checkoutCommit(docPath, oids[0]),
        checkoutCommit(docPath, oids[1]),
        checkoutCommit(docPath, oids[2]),
      ]);

      expect(c0.xml).toBe('<w:p>Rev 3</w:p>');
      expect(c1.xml).toBe('<w:p>Rev 2</w:p>');
      expect(c2.xml).toBe('<w:p>Rev 1</w:p>');

      const afterHistory = await getHistory(docPath);
      expect(afterHistory.map((c) => c.oid)).toEqual(oids);
    });
  });

  describe('6. Zero System Git Dependency: PATH="" Guarantee', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case6-no-path');
    const docPath = path.join(caseDir, 'zero_git_dep.docx');

    it('executes full git versioning and checkout cycle under process.env.PATH=""', async () => {
      await fsp.mkdir(caseDir, { recursive: true });

      const originalPath = process.env.PATH;
      try {
        process.env.PATH = '';

        // Initialize repo
        await initAstraRepo(docPath, '<w:p>No System Git v1</w:p>', {
          'word/styles.xml': '<w:styles><w:style w:styleId="Normal"/></w:styles>',
        });

        // Commit change
        const s1Oid = await commitChange(
          docPath,
          '<w:p>No System Git v2</w:p>',
          'Committed without system git in PATH',
          { 'word/styles.xml': '<w:styles><w:style w:styleId="Heading1"/></w:styles>' }
        );

        // Get history
        const history = await getHistory(docPath);
        expect(history.length).toBe(2);
        expect(history[0].oid).toBe(s1Oid);

        // Checkout intermediate commit
        const co0 = await checkoutCommit(docPath, history[1].oid);
        expect(co0.xml).toBe('<w:p>No System Git v1</w:p>');
        expect(co0.extra.get('word/styles.xml')?.toString('utf8')).toContain('Normal');

        // Undo 1 step
        const undoRes = await undoPackage(docPath);
        expect(undoRes.xml).toBe('<w:p>No System Git v1</w:p>');

        const historyAfterUndo = await getHistory(docPath);
        expect(historyAfterUndo.length).toBe(1);
      } finally {
        process.env.PATH = originalPath;
      }
    });
  });

  describe('7. DocumentManager End-to-End Git Lifecycle Integration', () => {
    const caseDir = path.join(SANDBOX_BASE, 'case7-doc-manager');
    const docPath = path.join(caseDir, 'dm_lifecycle.docx');

    it('manages open, commitTransaction, undoLastAction, redoLastAction, and validates on disk', async () => {
      await fsp.mkdir(caseDir, { recursive: true });
      await fsp.copyFile(path.join(TEST_GALLERY_DIR, 'T4_budget.docx'), docPath);
      registerTrustedDocPath(docPath, 'dialog');

      const dm = new DocumentManager();
      await dm.openDocument(docPath);

      const history0 = await dm.getVersionHistory();
      expect(history0.length).toBe(1);

      // Stage 1 edit
      const parser = dm.getParser();
      let ast = await parser.getAST();
      let engine = new OperationEngine(ast, parser);
      const lastNode1 = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];

      await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'paragraph',
            anchor: lastNode1.id,
            position: 'after',
            text: 'DM Stage 1: Quarterly Strategic Realignment Notes',
          } as never,
        ],
      });
      await dm.commitTransaction('DM Stage 1: Added realignment paragraph');

      const valFile1 = await validateDocxFile(docPath);
      expect(valFile1.ok).toBe(true);

      // Stage 2 edit: add table
      ast = await parser.getAST();
      engine = new OperationEngine(ast, parser);
      const lastNode2 = ast.sections[0].nodes[ast.sections[0].nodes.length - 1];

      await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'table',
            anchor: lastNode2.id,
            position: 'after',
            headers: ['Milestone', 'Target Date', 'Confidence'],
            rows: [['Phase 1 Cutover', '2026-10-01', 'High']],
          } as never,
        ],
      });
      await dm.commitTransaction('DM Stage 2: Added milestone schedule table');

      const valFile2 = await validateDocxFile(docPath);
      expect(valFile2.ok).toBe(true);

      const history2 = await dm.getVersionHistory();
      expect(history2.length).toBe(3);

      // Undo 1 step via DocumentManager
      const undoRes = await dm.undoLastAction();
      expect(undoRes.success).toBe(true);

      const diskBufAfterUndo = await fsp.readFile(docPath);
      const diskParserAfterUndo = new DocxParser(diskBufAfterUndo);
      const diskXmlAfterUndo = diskParserAfterUndo.readDocumentXml();

      expect(diskXmlAfterUndo).not.toContain('Phase 1 Cutover');
      expect(diskXmlAfterUndo).toContain('DM Stage 1: Quarterly Strategic Realignment Notes');

      const valUndo = await validateDocxFile(docPath);
      expect(valUndo.ok).toBe(true);

      // Redo 1 step via DocumentManager
      const redoRes = await dm.redoLastAction();
      expect(redoRes.success).toBe(true);

      const diskBufAfterRedo = await fsp.readFile(docPath);
      const diskParserAfterRedo = new DocxParser(diskBufAfterRedo);
      const diskXmlAfterRedo = diskParserAfterRedo.readDocumentXml();

      expect(diskXmlAfterRedo).toContain('Phase 1 Cutover');

      const valRedo = await validateDocxFile(docPath);
      expect(valRedo.ok).toBe(true);
    }, 30_000);
  });
});
