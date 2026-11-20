/**
 * Locks what `rezip()` does and does not check, so the documentation and the code stay honest:
 * `rezip()` runs integrity, duplicate-paraId, XXE and macro checks. It does NOT run the Microsoft schema
 * validator. Schema validity is checked by `validateDocxBuffer()` or `parser.validateOOXML()`.
 */
import { describe, expect, it } from 'vitest';
import { DocxParser, validateDocxBuffer } from '../src';

const BAD_SIZE = '<w:p><w:r><w:rPr><w:sz w:val="notanumber"/></w:rPr><w:t>bad size</w:t></w:r></w:p>';

function parserWithBadSize(): DocxParser {
  const parser = DocxParser.createNewDocument('blank');
  const id = parser.getStructuralSummary().nodes[0].id;
  parser.updateNode(id, BAD_SIZE);
  return parser;
}

describe('rezip() is not the schema validator', () => {
  it('well-formed but schema-invalid XML passes rezip() and fails validateDocxBuffer()', async () => {
    const bytes = parserWithBadSize().rezip();
    expect(bytes.length).toBeGreaterThan(1000);
    const result = await validateDocxBuffer(bytes);
    expect(result.ok).toBe(false);
    expect(result.errors[0].description).toContain("The value 'notanumber' is not valid");
  });

  it('parser.validateOOXML() reports the same error without writing anything', async () => {
    const result = await parserWithBadSize().validateOOXML();
    expect(result.ok).toBe(false);
    expect(JSON.stringify(result.errors)).toContain('notanumber');
  });

  it('a document that was not edited validates after rezip()', async () => {
    const result = await validateDocxBuffer(DocxParser.createNewDocument('blank').rezip());
    expect(result.ok).toBe(true);
  });
});
