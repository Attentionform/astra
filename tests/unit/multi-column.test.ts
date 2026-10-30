import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  OperationEngine,
  SectionOperation,
} from '@attentionform/astra-engine';

import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';

function createMultiNodeDocx(): Buffer {
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
    <w:p data-astra-id="11111111" w14:paraId="11111111"><w:r><w:t>Paragraph 1</w:t></w:r></w:p>
    <w:p data-astra-id="22222222" w14:paraId="22222222"><w:r><w:t>Paragraph 2</w:t></w:r></w:p>
    <w:p data-astra-id="33333333" w14:paraId="33333333"><w:r><w:t>Paragraph 3</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Section & Multi-Column Operation Suite', () => {
  it('configures multi-column section layout (2 columns)', async () => {
    const buffer = createMultiNodeDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const secOp: SectionOperation = {
      action: 'section',
      operation: 'setColumns',
      columns: 2,
      spacing: 720
    };

    const result = await engine.executeBatch({ operations: [secOp] });
    expect(result.success).toBe(true);
    const doc = ast.sections[0].nodes[0].xml.ownerDocument;
    const docXml = new XMLSerializer().serializeToString(doc!);
    expect(docXml).toContain('w:cols');
    expect(docXml).toContain('w:num="2"');
  });

  it('splits section at target node and merges sections', async () => {
    const buffer = createMultiNodeDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    // Split at p_002
    const splitOp: SectionOperation = {
      action: 'section',
      operation: 'split',
      target: '22222222'
    };

    const splitResult = await engine.executeBatch({ operations: [splitOp] });
    expect(splitResult.success).toBe(true);
    expect(ast.sections.length).toBe(2);
    expect(ast.sections[0].nodes.length).toBe(1); // p_001
    expect(ast.sections[1].nodes.length).toBe(2); // p_002, p_003

    // Merge back
    const mergeOp: SectionOperation = {
      action: 'section',
      operation: 'merge',
      sectionIndex: 1
    };

    const mergeResult = await engine.executeBatch({ operations: [mergeOp] });
    expect(mergeResult.success).toBe(true);
    expect(ast.sections.length).toBe(1);
    expect(ast.sections[0].nodes.length).toBe(3);
  });
});
