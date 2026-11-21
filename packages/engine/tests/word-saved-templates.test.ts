import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import AdmZip from 'adm-zip';
import { DocxParser, TEMPLATE_NAMES, composeDocument, validateDocxBuffer } from '../src/internal';

// Each fixture is a built-in template that Microsoft Word opened and saved (see scripts/templates/word-roundtrip/README.md).
const DESIGNED = TEMPLATE_NAMES.filter((n) => n !== 'blank');
const fixture = (name: string, variant = ''): Buffer => readFileSync(join(__dirname, 'fixtures', 'word-saved', variant, `${name}.docx`));
const textOf = (xml: string): string => Array.from(xml.matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g), (m) => m[1]).join('\n');

// '' is Letter. 'a5' is A5 with its default margins. 'custom' is 6 x 9 in with 0.8 in side margins, 0.9 in top and bottom and a 0.25 in gutter.
describe.each(['', 'a5', 'custom'])('a template that Word opened and saved is still recognised (%s page)', (variant) => {
  for (const name of DESIGNED) {
    it(`${name}: compose replaces the placeholder body`, async () => {
      const parser = new DocxParser(fixture(name, variant));
      const res = await composeDocument(parser, '# Quarterly Review\n\nFirst paragraph of the body.\n');
      expect(res.success, JSON.stringify(res.errors)).toBe(true);
      expect(res.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
      const out = parser.rezip();
      const text = textOf(new AdmZip(out).readAsText('word/document.xml'));
      expect(text).toContain('First paragraph of the body.');
      expect(text.match(/Quarterly Review/g)).toHaveLength(1);
      expect(text).not.toMatch(/Summarise the purpose|State in one or two sentences|Open with the reason|Lead story headline|Chapter One/);
      const valid = await validateDocxBuffer(out);
      expect(valid.ok, JSON.stringify(valid.errors)).toBe(true);
    });
  }

  it.runIf(variant === '')('a body edited after Word saved it is kept: TEMPLATE_REGION_EDITED', async () => {
    const zip = new AdmZip(fixture('memo'));
    const doc = zip.readAsText('word/document.xml');
    expect(doc).toContain('[First action, with owner and date]');
    zip.updateFile('word/document.xml', Buffer.from(doc.replace('[First action, with owner and date]', 'Ship it on Friday')));
    const parser = new DocxParser(zip.toBuffer());
    const res = await composeDocument(parser, 'Appended.\n');
    expect(res.warnings.some((w) => w.startsWith('TEMPLATE_REGION_EDITED'))).toBe(true);
    expect(textOf(parser.readDocumentXml())).toContain('Ship it on Friday');
  });
});
