// Find out which renderers are installed, then make a PDF and page images when they are.
import { writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { DocxParser, composeDocument, convertDocxToPdf, getVisualArtifactsDir, getVisualBackendStatus, hasPdfBackend, renderDocxToImages } from '@attentionform/astra-engine';
import { assertApplied, outputDir } from './_shared.mjs';

// What is installed. `plan.renderers` and `plan.rasterizers` list the programs a default call would use, in order.
const status = getVisualBackendStatus();
console.log(`renderers: Word ${status.renderers?.word.available ? `(version ${status.renderers.word.version})` : 'not found'}, ` +
  `LibreOffice ${status.renderers?.libreoffice.available ? 'found' : 'not found'}`);
console.log(`page-image programs: Poppler ${status.rasterizers?.poppler.available ? 'found' : 'not found'}, ` +
  `Windows renderer ${status.rasterizers?.windows.available ? 'found' : 'not found'}`);
console.log(`a PDF can be made: ${hasPdfBackend()}; page images can be made: ${status.available}`);

const parser = DocxParser.createNewDocument('report');
assertApplied(await composeDocument(parser, '## Results\n\nThe numbers are in.\n'), 'compose');
const docx = parser.rezip();
const dir = outputDir('10');

if (!hasPdfBackend()) {
  console.log('no renderer is installed, so there is nothing to render here (install LibreOffice, or use Word on Windows)');
} else {
  const pdf = await convertDocxToPdf(docx, join(dir, 'report.pdf'));
  console.log(`wrote ${pdf}`);
}

if (!status.available) {
  console.log(`page images need ${status.missing.join(' and ')}; skipped`);
} else {
  // Page images are written inside the visual artifacts folder; getVisualArtifactsDir makes a folder for this run.
  const result = await renderDocxToImages(docx, { outDir: getVisualArtifactsDir('example-10'), maxPages: 1 });
  if (!result.images.length) throw new Error('no page image was made');
  console.log(`wrote ${result.images.length} page image(s) of ${result.totalPages} to ${result.outDir}, drawn by ${result.backend}`);
  writeFileSync(join(dir, 'status.json'), JSON.stringify(status, null, 2));
}
