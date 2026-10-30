import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  OperationEngine,
  StyleOperation,
} from '@attentionform/astra-engine';

import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';

function createMinimalDocx(): Buffer {
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p data-astra-id="p_001" w14:paraId="11111111">
      <w:r><w:t>Advanced typography sample paragraph</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Advanced Typography Suite', () => {
  it('applies highlighting, superscript, and line spacing with schema ordering', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const styleOp: StyleOperation = {
      action: 'style',
      target: '11111111',
      properties: {
        highlight: 'yellow',
        vertAlign: 'superscript',
        lineSpacing: 360, // 1.5x
        spacingBefore: 240, // 12pt
        spacingAfter: 120, // 6pt
        indentLeft: 720 // 0.5 in
      }
    };

    const result = await engine.executeBatch({ operations: [styleOp] });
    expect(result.success).toBe(true);

    const node = ast.sections[0].nodes.find(n => n.id === '11111111');
    expect(node).toBeDefined();

    const xmlStr = new XMLSerializer().serializeToString(node!.xml);
    expect(xmlStr).toContain('w:highlight w:val="yellow"');
    expect(xmlStr).toContain('w:vertAlign w:val="superscript"');
    expect(xmlStr).toContain('w:spacing w:line="360" w:lineRule="auto" w:before="240" w:after="120"');
    expect(xmlStr).toContain('w:ind w:left="720"');
  });

  it('supports double strikethrough and subscript', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const styleOp: StyleOperation = {
      action: 'style',
      target: '11111111',
      properties: {
        doubleStrikethrough: true,
        vertAlign: 'subscript'
      }
    };

    const result = await engine.executeBatch({ operations: [styleOp] });
    expect(result.success).toBe(true);

    const node = ast.sections[0].nodes.find(n => n.id === '11111111');
    const xmlStr = new XMLSerializer().serializeToString(node!.xml);
    expect(xmlStr).toContain('w:dstrike');
    expect(xmlStr).toContain('w:vertAlign w:val="subscript"');
  });
});
