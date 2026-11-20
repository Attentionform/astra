// Check a document against the Office schema, and see what a failure looks like.
import { AdmZip, DocxParser, OoxmlValidationError, assertValidDocx, getEngineErrorCode, validateDocxBuffer } from '@attentionform/astra-engine';

const good = DocxParser.createNewDocument('memo').rezip();
const ok = await validateDocxBuffer(good);
if (!ok.ok) throw new Error('a new memo should be valid');
console.log(`a new memo is valid (${ok.errors.length} errors)`);

// Make a file that is well-formed XML but breaks the schema: a font size that is not a number.
const zip = new AdmZip(good);
const xml = zip.readAsText('word/document.xml');
const broken = xml.replace('</w:p>', '<w:r><w:rPr><w:sz w:val="notanumber"/></w:rPr><w:t>x</w:t></w:r></w:p>');
if (broken === xml) throw new Error('the example could not break the document');
zip.updateFile('word/document.xml', Buffer.from(broken, 'utf8'));
const bad = zip.toBuffer();

// validateDocxBuffer returns the result; you decide what to do.
const result = await validateDocxBuffer(bad);
if (result.ok) throw new Error('the broken file should not be valid');
const first = result.errors[0];
console.log(`not valid: ${first.description}`);
console.log(`  where: ${first.part} ${first.xPath}`);

// assertValidDocx throws, with a stable code you can handle.
try {
  await assertValidDocx(bad);
  throw new Error('assertValidDocx should have thrown');
} catch (err) {
  if (!(err instanceof OoxmlValidationError)) throw err;
  console.log(`assertValidDocx threw ${err.name}, code ${getEngineErrorCode(err)}`);
}
