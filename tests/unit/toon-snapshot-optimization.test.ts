import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import AdmZip from 'adm-zip';
import * as path from 'path';
import * as fsp from 'fs/promises';
import os from 'os';
import {
  DocumentManager,
  DocxParser,
  ToonGenerator,
} from '@attentionform/astra-engine';

describe('TOON Token Optimization & Atomic Snapshot History Suite', () => {
  let mockDocxBuffer: Buffer;
  let tmpDir: string;
  let testDocPath: string;

  beforeEach(async () => {
    tmpDir = await fsp.mkdtemp(path.join(os.tmpdir(), 'astra-toon-test-'));
    testDocPath = path.join(tmpDir, 'test_sample.docx');

    const zip = new AdmZip();
    const mockXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="11111111">
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Introduction and Executive Overview</w:t></w:r>
    </w:p>
    <w:p w14:paraId="22222222">
      <w:r><w:t>This is a detailed paragraph explaining high-performance TOON token optimization and snapshot history in ASTra document engine.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tblPr/>
      <w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>
      <w:tr><w:tc><w:p><w:r><w:t>Sample Table Cell Content</w:t></w:r></w:p></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>`;
    zip.addFile('word/document.xml', Buffer.from(mockXml, 'utf8'));
    zip.addFile(
      '[Content_Types].xml',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        'utf8'
      )
    );
    zip.addFile(
      '_rels/.rels',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
        'utf8'
      )
    );
    zip.writeZip(testDocPath);
    mockDocxBuffer = zip.toBuffer();
  });

  afterEach(async () => {
    await fsp.rm(tmpDir, { recursive: true, force: true });
  });

  describe('1. High-Performance TOON Token Optimization', () => {
    it('should generate TOON table and render compact token-optimized text', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const ast = await parser.getAST();
      const generator = new ToonGenerator();

      const table = generator.generateToon(ast);
      expect(table.rows.length).toBe(3);

      const standardText = generator.renderToonText(table);
      const compactText = generator.renderToonText(table, { compact: true, omitColumns: ['flags', 'path'] });

      expect(compactText.length).toBeLessThan(standardText.length);
      expect(compactText).not.toContain('---');
    });

    it('should truncate text length and estimate tokens accurately', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const ast = await parser.getAST();
      const generator = new ToonGenerator();

      const truncatedTable = generator.generateToon(ast, undefined, { maxTextLength: 20 });
      expect(truncatedTable.rows[1].text).toContain('...');
      expect(truncatedTable.rows[1].text.length).toBeLessThanOrEqual(23);

      const tokens = generator.estimateTableTokens(truncatedTable);
      expect(tokens).toBeGreaterThan(0);
    });

    it('should dynamically enforce a target token budget with generateTokenBudgetToon', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const ast = await parser.getAST();
      const generator = new ToonGenerator();

      const res = generator.generateTokenBudgetToon(ast, 80);
      expect(res.tokenCount).toBeLessThanOrEqual(80);
      expect(res.tokenCount).toBe(generator.estimateTokens(res.rendered));
      expect(res.rendered).toBeTruthy();
      expect(res.table.rows.length).toBe(3);
      expect(res.truncated).toBe(false);
      expect(res.hasMore).toBe(false);
      expect(res.nextOffset).toBeNull();
    });

    it('should page a large document through the whole row set within budget', async () => {
      const filler = 'Budget paging probe paragraph carrying filler words to pad every TOON row. ';
      const paras = Array.from(
        { length: 250 },
        (_, i) =>
          `<w:p w14:paraId="B${String(i).padStart(5, '0')}"><w:r><w:t>${i} ${filler.repeat(4)}</w:t></w:r></w:p>`
      ).join('');
      const zip = new AdmZip();
      zip.addFile(
        'word/document.xml',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>${paras}</w:body></w:document>`,
          'utf8'
        )
      );
      zip.addFile('[Content_Types].xml', Buffer.from('<Types></Types>', 'utf8'));
      const parser = new DocxParser(zip.toBuffer());
      const ast = await parser.getAST();
      const generator = new ToonGenerator();
      const full = generator.generateToon(ast);
      expect(full.rows.length).toBe(250);

      const seen: string[] = [];
      let offset = 0;
      let pages = 0;
      for (;;) {
        const page = generator.generateTokenBudgetToon(ast, 1500, { offset });
        expect(page.tokenCount).toBe(generator.estimateTokens(page.rendered));
        expect(page.tokenCount).toBeLessThanOrEqual(1500);
        expect(page.table.rows.length).toBeGreaterThan(0);
        for (const row of page.table.rows) seen.push(row.id);
        pages += 1;
        if (!page.hasMore) {
          expect(page.nextOffset).toBeNull();
          break;
        }
        expect(typeof page.nextOffset).toBe('number');
        expect(page.nextOffset).toBeGreaterThan(offset);
        offset = page.nextOffset as number;
        if (pages > 300) throw new Error('budget paging did not terminate');
      }
      expect(seen).toEqual(full.rows.map((r) => r.id));
    });

    it('should return a single over-budget row with flags instead of looping', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const ast = await parser.getAST();
      const generator = new ToonGenerator();
      const single = {
        ...ast,
        sections: [{ ...ast.sections[0], nodes: [ast.sections[0].nodes[1]] }],
      };
      const res = generator.generateTokenBudgetToon(single, 20);
      expect(res.table.rows.length).toBe(1);
      expect(res.truncated).toBe(true);
      expect(res.hasMore).toBe(false);
      expect(res.nextOffset).toBeNull();
    });

    it('should leave small documents byte-identical with a quiet cursor', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const ast = await parser.getAST();
      const generator = new ToonGenerator();

      const before = generator.generateTokenBudgetToon(ast, 100000);
      const table = generator.generateToon(ast);
      const rendered = generator.renderToonText(table, { maxTextLength: 120 });
      expect(before.rendered).toBe(rendered);
      expect(before.table.rows.length).toBe(table.rows.length);
      expect(before.tokenCount).toBe(generator.estimateTokens(rendered));
      expect(before.truncated).toBe(false);
      expect(before.hasMore).toBe(false);
      expect(before.nextOffset).toBeNull();
    });

    it('should calculate diffs with ADDED and MODIFIED flags', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const ast = await parser.getAST();
      const generator = new ToonGenerator();

      const oldTable = generator.generateToon(ast);
      ast.sections[0].nodes[1].version += 1;
      ast.sections[0].nodes[1].metadata.text = 'Modified text paragraph content';

      const newTable = generator.generateToon(ast, undefined, { useCache: false });
      const diff = generator.generateDiff(oldTable, newTable);

      expect(diff.rows.length).toBe(1);
      expect(diff.rows[0].flags).toContain('[~MODIFIED]');
    });
  });

  describe('2. DocxParser Atomic Snapshot & XML Reload', () => {
    it('should create and restore atomic snapshot cleanly', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const snapshot = parser.createAtomicSnapshot('Test Snapshot');

      expect(snapshot.label).toBe('Test Snapshot');
      expect(snapshot.xmlContent).toContain('Introduction and Executive Overview');

      // Update node
      parser.updateNode('11111111', `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="11111111"><w:r><w:t>Mutated Title</w:t></w:r></w:p>`);
      expect(parser.readNode('11111111')).toContain('Mutated Title');

      // Restore atomic snapshot
      parser.restoreAtomicSnapshot(snapshot);
      expect(parser.readNode('11111111')).toContain('Introduction and Executive Overview');
    });

    it('should safely reload document XML without security flaws', async () => {
      const parser = new DocxParser(mockDocxBuffer);
      const newXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="19999999"><w:r><w:t>Reloaded Document XML</w:t></w:r></w:p>
  </w:body>
</w:document>`;

      parser.reloadDocumentXml(newXml);
      expect(parser.readNode('19999999')).toContain('Reloaded Document XML');

      // XXE Rejection check
      const xxeXml = `<!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>`;
      expect(() => parser.reloadDocumentXml(xxeXml)).toThrow(/Security Error: XXE/);
    });
  });

  describe('3. DocumentManager Atomic Snapshot Stack & Transactions', () => {
    it('should manage in-memory snapshot history stack and support undo/redo', async () => {
      const docManager = new DocumentManager();
      await docManager.openDocument(testDocPath);

      expect(docManager.canUndo()).toBe(false);
      expect(docManager.getSnapshotHistory().length).toBe(1);

      // Perform update
      await docManager.updateNode('11111111', `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="11111111"><w:r><w:t>Updated P1 Title</w:t></w:r></w:p>`, 'Update P1');

      expect(docManager.canUndo()).toBe(true);
      expect(docManager.getSnapshotHistory().length).toBe(2);

      await docManager.undoLastAction();
      expect(docManager.getParser().readNode('11111111')).toContain('Introduction and Executive Overview');
    });

    it('should handle atomic transactions with commit and rollback', async () => {
      const docManager = new DocumentManager();
      await docManager.openDocument(testDocPath);

      docManager.beginTransaction('Batch Edit');
      docManager.getParser().updateNode('11111111', `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="11111111"><w:r><w:t>Transaction Edit P1</w:t></w:r></w:p>`);

      expect(docManager.getParser().readNode('11111111')).toContain('Transaction Edit P1');

      // Rollback transaction
      docManager.rollbackTransaction();
      expect(docManager.getParser().readNode('11111111')).toContain('Introduction and Executive Overview');

      // Begin another and commit
      docManager.beginTransaction('Successful Batch');
      docManager.getParser().updateNode('11111111', `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" w14:paraId="11111111"><w:r><w:t>Committed P1</w:t></w:r></w:p>`);
      await docManager.commitTransaction('Commit Batch');

      expect(docManager.getParser().readNode('11111111')).toContain('Committed P1');
    });
  });
});
