import AdmZip from 'adm-zip';
import { createBlankDocxBuffer } from '../../packages/engine/src/parser';

// A blank document with a Title and a Subtitle paragraph, for tests that need a few nodes to work on.
// (`blank` itself is one empty paragraph.)
export function seededBlank(): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const xml = zip.readAsText('word/document.xml');
  const seeded = xml.replace(
    /(<w:body[^>]*>)[\s\S]*?(<w:sectPr)/,
    '$1<w:p><w:pPr><w:pStyle w:val="Title"/></w:pPr><w:r><w:t>Blank Document</w:t></w:r></w:p>' +
      '<w:p><w:pPr><w:pStyle w:val="Subtitle"/></w:pPr><w:r><w:t>Start typing your content here...</w:t></w:r></w:p>$2',
  );
  zip.updateFile('word/document.xml', Buffer.from(seeded, 'utf8'));
  return zip.toBuffer();
}
