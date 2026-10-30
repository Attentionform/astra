import { describe, it, expect, vi } from 'vitest';
import AdmZip from 'adm-zip';
import { XMLSerializer } from '@xmldom/xmldom';
import {
  DeleteOperation,
  DocxParser,
  executeDelete,
  executePageSetup,
  executeTable,
  OperationEngine,
  PageSetupOperation,
  TableOperation,
} from '@attentionform/astra-engine';

function createMockDocxWithTable(tableId: string = 't_1'): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`)
  );
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>`)
  );
  zip.addFile(
    'word/document.xml',
    Buffer.from(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p data-astra-id="p_001" w14:paraId="11111111">
      <w:r><w:t>First Paragraph</w:t></w:r>
    </w:p>
    <w:tbl data-astra-id="${tableId}">
      <w:tblPr/>
      <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Col 1 Header</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Col 2 Header</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Col 3 Header</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Row 1 Col 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Row 1 Col 2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Row 1 Col 3</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Row 2 Col 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Row 2 Col 2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Row 2 Col 3</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Row 3 Col 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Row 3 Col 2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Row 3 Col 3</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:p data-astra-id="p_002" w14:paraId="22222222">
      <w:r><w:t>Second Paragraph</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Sprint 1 Resilient Operations Unit Tests', () => {
  // Requirement 1: table.ts resilient flat top-level op properties
  describe('1. table.ts - Flat top-level op properties', () => {
    it('executes setShading with flat op properties, headerFill #1E3A5F, fill #F8FAFC, and zebra striping', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const op: TableOperation = {
        action: 'table',
        target: 't_1',
        op: 'setShading',
        headerFill: '#1E3A5F',
        fill: '#F8FAFC',
        zebra: true
      } as any;

      await executeTable(op, ast);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      expect(tableNode).toBeDefined();

      const tblEl = tableNode!.xml as Element;
      const rows = Array.from(tblEl.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tr') as Element[];
      expect(rows.length).toBe(4);

      // Header row (index 0): headerFill #1E3A5F -> stripped '#' becomes '1E3A5F'
      const headerCells = Array.from(rows[0].childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
      for (const tc of headerCells) {
        const shd = tc.getElementsByTagName('w:shd')[0];
        expect(shd).toBeDefined();
        expect(shd.getAttribute('w:val')).toBe('clear');
        expect(shd.getAttribute('w:color')).toBe('auto');
        expect(shd.getAttribute('w:fill')).toBe('1E3A5F');
      }

      // Data row 1 (odd index): fill #F8FAFC -> 'F8FAFC'
      const row1Cells = Array.from(rows[1].childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
      for (const tc of row1Cells) {
        const shd = tc.getElementsByTagName('w:shd')[0];
        expect(shd).toBeDefined();
        expect(shd.getAttribute('w:fill')).toBe('F8FAFC');
      }

      // Data row 2 (even index): alternate zebra fill 'FFFFFF'
      const row2Cells = Array.from(rows[2].childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
      for (const tc of row2Cells) {
        const shd = tc.getElementsByTagName('w:shd')[0];
        expect(shd).toBeDefined();
        expect(shd.getAttribute('w:fill')).toBe('FFFFFF');
      }

      // Data row 3 (odd index): fill 'F8FAFC'
      const row3Cells = Array.from(rows[3].childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
      for (const tc of row3Cells) {
        const shd = tc.getElementsByTagName('w:shd')[0];
        expect(shd).toBeDefined();
        expect(shd.getAttribute('w:fill')).toBe('F8FAFC');
      }
    });

    it('executes setShading with flat op properties and zebra: false', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const op: TableOperation = {
        action: 'table',
        target: 't_1',
        op: 'setShading',
        fill: '#E2E8F0',
        zebra: false
      } as any;

      await executeTable(op, ast);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const tblEl = tableNode!.xml as Element;
      const rows = Array.from(tblEl.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tr') as Element[];

      for (const row of rows) {
        const cells = Array.from(row.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
        for (const tc of cells) {
          const shd = tc.getElementsByTagName('w:shd')[0];
          expect(shd).toBeDefined();
          expect(shd.getAttribute('w:fill')).toBe('E2E8F0');
        }
      }
    });

    it('executes setHeaderRepeat using flat top-level op and rowIndex', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const op: TableOperation = {
        action: 'table',
        target: 't_1',
        op: 'setHeaderRepeat',
        rowIndex: 1
      } as any;

      await executeTable(op, ast);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const tblEl = tableNode!.xml as Element;
      const rows = Array.from(tblEl.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tr') as Element[];

      const headerTrPr = rows[0].getElementsByTagName('w:trPr')[0];
      expect(headerTrPr).toBeDefined();
      expect(headerTrPr.getElementsByTagName('w:tblHeader').length).toBe(1);

      const dataTrPr = rows[1].getElementsByTagName('w:trPr')[0];
      if (dataTrPr) {
        expect(dataTrPr.getElementsByTagName('w:tblHeader').length).toBe(0);
      }
    });

    it('executes setCantSplit using flat top-level op applying w:cantSplit to all rows', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const op: TableOperation = {
        action: 'table',
        target: 't_1',
        op: 'setCantSplit'
      } as any;

      await executeTable(op, ast);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const tblEl = tableNode!.xml as Element;
      const rows = Array.from(tblEl.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tr') as Element[];

      for (let i = 0; i < rows.length; i++) {
        const trPr = rows[i].getElementsByTagName('w:trPr')[0];
        expect(trPr).toBeDefined();
        expect(trPr.getElementsByTagName('w:cantSplit').length).toBe(1);
      }
    });

    it('executes setColumnWidths using flat top-level op and widths array', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const op: TableOperation = {
        action: 'table',
        target: 't_1',
        op: 'setColumnWidths',
        widths: [2500, 3500, 4000]
      } as any;

      await executeTable(op, ast);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const tblEl = tableNode!.xml as Element;

      const tblGrid = tblEl.getElementsByTagName('w:tblGrid')[0];
      expect(tblGrid).toBeDefined();
      const gridCols = Array.from(tblGrid.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:gridCol') as Element[];
      expect(gridCols.length).toBe(3);
      expect(gridCols[0].getAttribute('w:w')).toBe('2500');
      expect(gridCols[1].getAttribute('w:w')).toBe('3500');
      expect(gridCols[2].getAttribute('w:w')).toBe('4000');

      const rows = Array.from(tblEl.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tr') as Element[];
      for (const row of rows) {
        const cells = Array.from(row.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
        expect(cells[0].getElementsByTagName('w:tcW')[0].getAttribute('w:w')).toBe('2500');
        expect(cells[1].getElementsByTagName('w:tcW')[0].getAttribute('w:w')).toBe('3500');
        expect(cells[2].getElementsByTagName('w:tcW')[0].getAttribute('w:w')).toBe('4000');
      }
    });

    it('executes setColumnWidth using flat top-level op, colIndex, and width', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const op: TableOperation = {
        action: 'table',
        target: 't_1',
        op: 'setColumnWidth',
        colIndex: 2,
        width: 4500
      } as any;

      await executeTable(op, ast);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const tblEl = tableNode!.xml as Element;

      const tblGrid = tblEl.getElementsByTagName('w:tblGrid')[0];
      const gridCols = Array.from(tblGrid.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:gridCol') as Element[];
      expect(gridCols[1].getAttribute('w:w')).toBe('4500');

      const rows = Array.from(tblEl.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tr') as Element[];
      for (const row of rows) {
        const cells = Array.from(row.childNodes).filter(n => n.nodeType === 1 && (n as Element).tagName === 'w:tc') as Element[];
        expect(cells[1].getElementsByTagName('w:tcW')[0].getAttribute('w:w')).toBe('4500');
      }
    });
  });

  // Requirement 2: operation-engine.ts routing execute_operation for table
  describe('2. operation-engine.ts - Routing execute_operation for table', () => {
    it('routes { action: "execute_operation", type: "table", target: "t_1", op: "setShading", fill: "F8FAFC" } through executeBatch()', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'execute_operation',
            type: 'table',
            target: 't_1',
            op: 'setShading',
            fill: 'F8FAFC'
          } as any
        ]
      });

      expect(result.success).toBe(true);
      expect(result.appliedCount).toBe(1);
      expect(result.rollbackTriggered).toBe(false);
      expect(result.errors.length).toBe(0);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const serialized = new XMLSerializer().serializeToString(tableNode!.xml);
      expect(serialized).toContain('w:fill="F8FAFC"');
    });

    it('routes multiple aliased table operations through OperationEngine.executeBatch()', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'execute_operations',
            type: 'table',
            target: 't_1',
            op: 'setHeaderRepeat',
            rowIndex: 1
          } as any,
          {
            action: 'execute_batch',
            targetType: 'table',
            target: 't_1',
            op: 'setCantSplit'
          } as any,
          {
            action: 'operation',
            opType: 'table',
            target: 't_1',
            op: 'setColumnWidths',
            widths: [2200, 3300, 4400]
          } as any
        ],
        sequential: true
      });

      expect(result.success).toBe(true);
      expect(result.appliedCount).toBe(3);
      expect(result.rollbackTriggered).toBe(false);

      const tableNode = ast.sections[0].nodes.find(n => n.type === 'table');
      const serialized = new XMLSerializer().serializeToString(tableNode!.xml);
      expect(serialized).toContain('w:tblHeader');
      expect(serialized).toContain('w:cantSplit');
      expect(serialized).toContain('w:w="2200"');
      expect(serialized).toContain('w:w="3300"');
      expect(serialized).toContain('w:w="4400"');
    });
  });

  // Requirement 3: delete.ts resilient deletion vs strict mode
  describe('3. delete.ts - Resilient deletion of non-existent node ID vs Strict mode', () => {
    it('deleting a non-existent node ID (pb_fix_01) in a batch reports skipped and does NOT fail batch or trigger rollback', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'delete',
            target: 'pb_fix_01'
          } as DeleteOperation
        ]
      });

      expect(result.success).toBe(true);
      // D2 fix: all-miss delete applies nothing → appliedCount 0 (was 1).
      expect(result.appliedCount).toBe(0);
      expect(result.rollbackTriggered).toBe(false);
      expect(result.errors.length).toBe(0);

      expect(result.skippedIds).toEqual(['pb_fix_01']);
    });

    it('direct executeDelete of non-existent node ID without strict does not throw', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      let res: ReturnType<typeof executeDelete>;
      expect(() => {
        res = executeDelete({
          action: 'delete',
          target: 'pb_fix_01'
        }, ast, parser);
      }).not.toThrow();

      expect(res!.skipped).toEqual(['pb_fix_01']);
      expect(res!.deleted).toEqual([]);
    });

    it('throws an error when deleting a non-existent node ID if strict: true is specified', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});

      expect(() => {
        executeDelete({
          action: 'delete',
          target: 'pb_fix_01',
          strict: true
        } as any, ast, parser);
      }).toThrow('Node pb_fix_01 not found for deletion');

      warnSpy.mockRestore();
    });

    it('triggers batch rollback when deleting a non-existent node ID with strict: true in a batch', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);
      const beforeXml = parser.readDocumentXml();

      const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'paragraph',
            anchor: '11111111',
            position: 'before',
            text: 'TEMP_INSERTION_BEFORE_FAIL'
          } as any,
          {
            action: 'delete',
            target: 'pb_fix_01',
            strict: true
          } as any
        ],
        sequential: true
      });

      expect(result.success).toBe(false);
      expect(result.rollbackTriggered).toBe(true);
      expect(result.errors.some(e => e.error.includes('Node pb_fix_01 not found for deletion'))).toBe(true);

      // Verify DOM was restored to pre-batch snapshot
      expect(parser.readDocumentXml()).toBe(beforeXml);
      expect(parser.readDocumentXml()).not.toContain('TEMP_INSERTION_BEFORE_FAIL');

      warnSpy.mockRestore();
    });

    it('mixed batch allows preceding and succeeding operations to succeed when deleting non-existent node without strict', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const warnSpy = vi.spyOn(console, 'warn').mockImplementation(() => {});

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'insert',
            type: 'paragraph',
            anchor: '11111111',
            position: 'before',
            text: 'RESILIENT_INSERTED_PREFIX'
          } as any,
          {
            action: 'delete',
            target: 'pb_fix_01'
          } as any,
          {
            action: 'replace',
            target: '22222222',
            text: 'RESILIENT_REPLACED_POSTFIX'
          } as any
        ],
        sequential: true
      });

      expect(result.success).toBe(true);
      // D2 fix: all-miss middle delete applies nothing → 2 applies (was 3).
      expect(result.appliedCount).toBe(2);
      expect(result.rollbackTriggered).toBe(false);

      const docXml = parser.readDocumentXml();
      expect(docXml).toContain('RESILIENT_INSERTED_PREFIX');
      expect(docXml).toContain('RESILIENT_REPLACED_POSTFIX');

      warnSpy.mockRestore();
    });
  });

  // Requirement 4: pageSetup.ts headerDistance and footerDistance
  describe('4. pageSetup.ts - headerDistance: 720 and footerDistance: 720', () => {
    it('sets headerDistance: 720 and footerDistance: 720 on <w:pgMar> via executePageSetup with flat properties', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      executePageSetup({
        action: 'pageSetup',
        headerDistance: 720,
        footerDistance: 720
      }, ast, parser);

      const sectPr = parser.getSectPrElements()[0];
      const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
      expect(pgMar).toBeDefined();
      expect(pgMar.getAttribute('w:header')).toBe('720');
      expect(pgMar.getAttribute('w:footer')).toBe('720');

      expect(ast.sections[0].pageSetup?.headerDistance).toBe(720);
      expect(ast.sections[0].pageSetup?.footerDistance).toBe(720);
    });

    it('sets header and footer aliases on <w:pgMar> via flat properties', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      executePageSetup({
        action: 'pageSetup',
        header: 720,
        footer: 720
      } as any, ast, parser);

      const sectPr = parser.getSectPrElements()[0];
      const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
      expect(pgMar).toBeDefined();
      expect(pgMar.getAttribute('w:header')).toBe('720');
      expect(pgMar.getAttribute('w:footer')).toBe('720');
    });

    it('sets headerDistance: 720 and footerDistance: 720 through nested params object', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();

      executePageSetup({
        action: 'pageSetup',
        params: {
          headerDistance: 720,
          footerDistance: 720,
          marginTop: 1080,
          marginBottom: 1080
        }
      }, ast, parser);

      const sectPr = parser.getSectPrElements()[0];
      const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
      expect(pgMar.getAttribute('w:header')).toBe('720');
      expect(pgMar.getAttribute('w:footer')).toBe('720');
      expect(pgMar.getAttribute('w:top')).toBe('1080');
      expect(pgMar.getAttribute('w:bottom')).toBe('1080');
    });

    it('executes pageSetup with headerDistance: 720 and footerDistance: 720 through OperationEngine.executeBatch()', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'pageSetup',
            headerDistance: 720,
            footerDistance: 720,
            marginTop: 720,
            marginBottom: 720
          } as PageSetupOperation
        ]
      });

      expect(result.success).toBe(true);
      expect(result.appliedCount).toBe(1);

      const sectPr = parser.getSectPrElements()[0];
      const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
      expect(pgMar.getAttribute('w:header')).toBe('720');
      expect(pgMar.getAttribute('w:footer')).toBe('720');
      expect(pgMar.getAttribute('w:top')).toBe('720');
      expect(pgMar.getAttribute('w:bottom')).toBe('720');

      const docXml = parser.readDocumentXml();
      expect(docXml).toContain('w:header="720"');
      expect(docXml).toContain('w:footer="720"');
    });

    it('routes generic execute_operation to pageSetup with headerDistance: 720 and footerDistance: 720', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      const result = await engine.executeBatch({
        operations: [
          {
            action: 'execute_operation',
            type: 'pageSetup',
            headerDistance: 720,
            footerDistance: 720
          } as any
        ]
      });

      expect(result.success).toBe(true);
      expect(result.appliedCount).toBe(1);

      const sectPr = parser.getSectPrElements()[0];
      const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
      expect(pgMar.getAttribute('w:header')).toBe('720');
      expect(pgMar.getAttribute('w:footer')).toBe('720');
    });

    it('maintains ECMA-376 schema ordering in sectPr (w:pgSz before w:pgMar before w:cols)', async () => {
      const buffer = createMockDocxWithTable('t_1');
      const parser = new DocxParser(buffer);
      const ast = await parser.getAST();
      const engine = new OperationEngine(ast, parser);

      await engine.executeBatch({
        operations: [
          {
            action: 'pageSetup',
            headerDistance: 720,
            footerDistance: 720,
            paperSize: 'A4',
            orientation: 'portrait',
            columns: 2
          } as any
        ]
      });

      const sectPr = parser.getSectPrElements()[0];
      const children = Array.from(sectPr.childNodes).filter(n => n.nodeType === 1).map(n => (n as Element).tagName);

      const iPgSz = children.indexOf('w:pgSz');
      const iPgMar = children.indexOf('w:pgMar');
      const iCols = children.indexOf('w:cols');

      expect(iPgSz).toBeGreaterThanOrEqual(0);
      expect(iPgMar).toBeGreaterThan(iPgSz);
      expect(iCols).toBeGreaterThan(iPgMar);
    });
  });
});
