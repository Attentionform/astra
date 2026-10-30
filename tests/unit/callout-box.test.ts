import { describe, it, expect } from 'vitest';
import {
  CalloutOperation,
  DocxParser,
  OperationEngine,
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
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p data-astra-id="p_001">
      <w:r><w:t>Main introduction paragraph</w:t></w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`)
  );
  return zip.toBuffer();
}

describe('Callout Box Operation Suite', () => {
  it('creates an Indigo executive callout box with accent left border and padding', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const calloutOp: CalloutOperation = {
      action: 'callout',
      anchor: 'p_001',
      position: 'after',
      title: 'Key Strategic Finding',
      text: 'Enterprise AI adoption expanded by 180% year-over-year in Q3.',
      theme: 'indigo'
    };

    const result = await engine.executeBatch({ operations: [calloutOp] });
    expect(result.success).toBe(true);
    expect(result.appliedCount).toBe(1);

    const calloutNode = ast.sections[0].nodes.find(n => (n.metadata as any)?.callout?.title === 'Key Strategic Finding');
    expect(calloutNode).toBeDefined();
    expect(calloutNode?.type).toBe('table');

    const xmlStr = new XMLSerializer().serializeToString(calloutNode!.xml);
    expect(xmlStr).toContain('w:tbl');
    expect(xmlStr).toContain('w:fill="EEF2FF"');
    expect(xmlStr).toContain('w:color="4F46E5"');
    expect(xmlStr).toContain('Key Strategic Finding');
    expect(xmlStr).toContain('Enterprise AI adoption');
  });

  it('supports emerald and amber callout themes', async () => {
    const buffer = createMinimalDocx();
    const parser = new DocxParser(buffer);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const calloutOp: CalloutOperation = {
      action: 'callout',
      anchor: 'p_001',
      position: 'after',
      title: 'Security Notice',
      text: 'All credentials are encrypted at rest with AES-256.',
      theme: 'emerald'
    };

    const result = await engine.executeBatch({ operations: [calloutOp] });
    expect(result.success).toBe(true);

    const calloutNode = ast.sections[0].nodes.find(n => (n.metadata as any)?.callout?.title === 'Security Notice');
    const xmlStr = new XMLSerializer().serializeToString(calloutNode!.xml);
    expect(xmlStr).toContain('w:fill="ECFDF5"');
    expect(xmlStr).toContain('w:color="059669"');
  });
});
