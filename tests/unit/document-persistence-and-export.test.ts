import { describe, it, expect } from 'vitest';
import {
  convertAstToHtml,
  convertAstToMarkdown,
  DocxParser,
} from '@attentionform/astra-engine';

import AdmZip from 'adm-zip';

function createDocumentWithTableAndHeadings(): Buffer {
  const zip = new AdmZip();
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Clinical Evidence &amp; Prevalence</w:t></w:r>
    </w:p>
    <w:p>
      <w:r><w:t>This evidence synthesis analyzes global health cohorts.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Cohort</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Sample (n)</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Prevalence (%)</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Medical Trainees</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>1,420</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>76.4%</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Consultants</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>850</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>31.2%</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720"/>
    </w:sectPr>
  </w:body>
</w:document>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
  return zip.toBuffer();
}

describe('Document Persistence & High-Fidelity Export Suite', () => {
  it('should accurately convert table XML to GitHub Flavored Markdown', async () => {
    const parser = new DocxParser(createDocumentWithTableAndHeadings());
    const ast = await parser.getAST();

    const markdown = convertAstToMarkdown(ast);
    expect(markdown).toContain('# Clinical Evidence & Prevalence');
    expect(markdown).toContain('| Cohort | Sample (n) | Prevalence (%) |');
    expect(markdown).toContain('| --- | --- | --- |');
    expect(markdown).toContain('| Medical Trainees | 1,420 | 76.4% |');
    expect(markdown).toContain('| Consultants | 850 | 31.2% |');
  });

  it('should accurately convert table XML to styled HTML with print support', async () => {
    const parser = new DocxParser(createDocumentWithTableAndHeadings());
    const ast = await parser.getAST();

    const html = convertAstToHtml(ast, 'Clinical Report', { printMode: true });
    expect(html).toContain('<h1 class="heading">Clinical Evidence &amp; Prevalence</h1>');
    expect(html).toContain('<table class="doc-table">');
    expect(html).toContain('<th style="background-color: #F1F5F9; color: #0F172A;">Cohort</th><th style="background-color: #F1F5F9; color: #0F172A;">Sample (n)</th><th style="background-color: #F1F5F9; color: #0F172A;">Prevalence (%)</th>');
    expect(html).toContain('<td>Medical Trainees</td><td>1,420</td><td>76.4%</td>');
    expect(html).toContain('@page');
  });

  it('should rezip and preserve byte-exact document XML', async () => {
    const parser = new DocxParser(createDocumentWithTableAndHeadings());
    const rezippedBuf = parser.rezip();

    const reparsed = new DocxParser(rezippedBuf);
    const roundtripXml = reparsed.readDocumentXml();

    expect(roundtripXml).toContain('Clinical Evidence &amp; Prevalence');
    expect(roundtripXml).toContain('Medical Trainees');
  });
});
