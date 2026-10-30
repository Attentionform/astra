import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  executePageSetup,
  OperationEngine,
} from '@attentionform/astra-engine';

import AdmZip from 'adm-zip';

// TX-05 (tests-only): 720 twips = 0.5 inch (1440 twips/inch). Local name only —
// product code has no exported constant, so this keeps the test readable
// without touching product files.
const HALF_INCH_TWIPS = 720;

function createMockDocx(): Buffer {
  const zip = new AdmZip();
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Hello world</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
  return zip.toBuffer();
}

describe('pageSetup Operation Test Suite', () => {
  it('should execute pageSetup with flat properties successfully', async () => {
    const parser = new DocxParser(createMockDocx());
    const ast = await parser.getAST();

    executePageSetup({
      action: 'pageSetup',
      marginTop: HALF_INCH_TWIPS,
      marginBottom: HALF_INCH_TWIPS,
      marginLeft: HALF_INCH_TWIPS,
      marginRight: HALF_INCH_TWIPS,
      paperSize: 'A4',
      orientation: 'portrait'
    }, ast, parser);

    const sectPr = parser.getSectPrElements()[0];
    const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
    expect(pgMar.getAttribute('w:top')).toBe(String(HALF_INCH_TWIPS));
    expect(pgMar.getAttribute('w:bottom')).toBe(String(HALF_INCH_TWIPS));
    expect(pgMar.getAttribute('w:left')).toBe(String(HALF_INCH_TWIPS));
    expect(pgMar.getAttribute('w:right')).toBe(String(HALF_INCH_TWIPS));

    const pgSz = sectPr.getElementsByTagName('w:pgSz')[0];
    expect(pgSz.getAttribute('w:w')).toBe('11906');
    expect(pgSz.getAttribute('w:h')).toBe('16838');

    expect(ast.sections[0].pageSetup?.marginTop).toBe(HALF_INCH_TWIPS);
    expect(ast.sections[0].pageSetup?.paperSize).toBe('A4');
  });

  it('should execute pageSetup with nested params object successfully', async () => {
    const parser = new DocxParser(createMockDocx());
    const ast = await parser.getAST();

    executePageSetup({
      action: 'pageSetup',
      params: {
        marginTop: 1000,
        marginBottom: 1000,
        marginLeft: 1000,
        marginRight: 1000,
        paperSize: 'Letter',
        orientation: 'landscape'
      }
    }, ast, parser);

    const sectPr = parser.getSectPrElements()[0];
    const pgMar = sectPr.getElementsByTagName('w:pgMar')[0];
    expect(pgMar.getAttribute('w:top')).toBe('1000');

    const pgSz = sectPr.getElementsByTagName('w:pgSz')[0];
    expect(pgSz.getAttribute('w:orient')).toBe('landscape');
    expect(parseInt(pgSz.getAttribute('w:w')!, 10)).toBeGreaterThan(parseInt(pgSz.getAttribute('w:h')!, 10));

    expect(ast.sections[0].pageSetup?.marginTop).toBe(1000);
    expect(ast.sections[0].pageSetup?.orientation).toBe('landscape');
  });

  it('should execute via OperationEngine in batch mode', async () => {
    const parser = new DocxParser(createMockDocx());
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);

    const result = await engine.executeBatch({
      operations: [
        {
          action: 'pageSetup',
          marginTop: 500,
          marginBottom: 500,
          marginLeft: 500,
          marginRight: 500,
          paperSize: 'A4'
        }
      ]
    });

    expect(result.success).toBe(true);
    expect(result.appliedCount).toBe(1);
    expect(ast.sections[0].pageSetup?.marginTop).toBe(500);
  });
});
