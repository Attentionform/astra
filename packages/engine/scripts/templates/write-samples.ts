/**
 * Writes every built-in template to a folder, for a look in Word or LibreOffice. Not part of the build: the templates are built at run time.
 * Usage: write-samples <outDir> [size] [orientation]   e.g. write-samples out A5 landscape
 */
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser } from '../../src/parser';
import { TEMPLATE_NAMES } from '../../src/ooxml/new-document-parts';

const [out, size, orientation] = process.argv.slice(2);
if (!out) throw new Error('usage: write-samples <outDir> [size] [portrait|landscape]');
fs.mkdirSync(out, { recursive: true });
for (const name of TEMPLATE_NAMES) {
  const page = { ...(size ? { size } : {}), ...(orientation ? { orientation } : {}) } as { size?: string; orientation?: 'portrait' | 'landscape' };
  const file = path.join(out, `${name}${size ? `.${size.replace(/\W+/g, '-')}` : ''}.docx`);
  fs.writeFileSync(file, DocxParser.createNewDocument(name, { page }).rezip());
  console.log(`wrote ${file}`);
}
