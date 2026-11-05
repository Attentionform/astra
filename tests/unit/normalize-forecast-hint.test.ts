// Compatibility coverage for deprecated 0.2.0 forecast helpers; CLI/MCP hints no longer call them.
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  forecastNormalize,
  normalizeHintSentence,
  type NormalizeForecast,
} from '@attentionform/astra-engine';

function docxWithMisplacedPPr(): Buffer {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const xml = zip.readAsText('word/document.xml');
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const end = xml.lastIndexOf('<w:sectPr');
  const badP = '<w:p><w:r><w:t>misplaced</w:t></w:r><w:pPr><w:jc w:val="center"/></w:pPr></w:p>';
  zip.updateFile('word/document.xml', Buffer.from(xml.slice(0, start) + badP + xml.slice(end), 'utf8'));
  return zip.toBuffer();
}

describe('forecastNormalize and normalizeHintSentence', () => {
  it('forecastNormalize returns undefined for a valid document with no repairs', async () => {
    const cleanBuf = DocxParser.createNewDocument('blank').rezip();
    const forecast = await forecastNormalize(cleanBuf);
    expect(forecast).toBeUndefined();
  });

  it('forecastNormalize reports changes and remaining error count when repair rules apply', async () => {
    const buf = docxWithMisplacedPPr();
    const forecast = await forecastNormalize(buf);
    expect(forecast).toBeDefined();
    expect(forecast!.changes).toBe(1);
    expect(forecast!.remaining).toBe(0);
  });

  it('normalizeHintSentence formats message when document fully validates after repair', () => {
    const forecast: NormalizeForecast = { changes: 2, remaining: 0 };
    const sentence = normalizeHintSentence(forecast, 2, 'Run `astra normalize doc.docx`: it');
    expect(sentence).toBe('Run `astra normalize doc.docx`: it can repair some of these errors (2 changes) and the document then validates.');
  });

  it('normalizeHintSentence formats message when remaining errors require Word', () => {
    const forecast: NormalizeForecast = { changes: 1, remaining: 3 };
    const sentence = normalizeHintSentence(forecast, 4, 'Run normalize_document: it');
    expect(sentence).toBe('Run normalize_document: it can repair some of these errors (1 change) and leaves 3 of 4 errors, which need Word.');
  });
});
