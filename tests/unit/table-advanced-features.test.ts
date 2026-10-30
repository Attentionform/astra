import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  OperationEngine,
  TableOperation,
} from '@attentionform/astra-engine';

import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';

function createTableDocx(): Buffer {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl data-astra-id="t_001">
      <w:tblGrid>
        <w:gridCol w:w="3000"/>
        <w:gridCol w:w="3000"/>
      </w:tblGrid>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Header A</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Header B</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Data 1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Data 2</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Advanced Table Features Suite', () => {
  it('applies setHeaderRepeat and setCantSplit on multi-page tables', async () => {
    const buffer = createTableDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const repeatOp: TableOperation = {
      action: 'table',
      target: 't_001',
      operation: 'setHeaderRepeat',
      rowIndex: 1
    };

    const cantSplitOp: TableOperation = {
      action: 'table',
      target: 't_001',
      operation: 'setCantSplit'
    };

    const vAlignOp: TableOperation = {
      action: 'table',
      target: 't_001',
      operation: 'setVerticalAlign',
      params: { vAlign: 'center' }
    };

    const result = await engine.executeBatch({ operations: [repeatOp, cantSplitOp, vAlignOp] });
    expect(result.success).toBe(true);

    const tableNode = ast.sections[0].nodes.find(n => n.id === 't_001');
    expect(tableNode).toBeDefined();

    const xmlStr = new XMLSerializer().serializeToString(tableNode!.xml);
    expect(xmlStr).toContain('w:tblHeader');
    expect(xmlStr).toContain('w:cantSplit');
    expect(xmlStr).toContain('w:vAlign w:val="center"');
  });
});
