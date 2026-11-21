import * as fs from 'node:fs';
import * as path from 'node:path';

// A minimal package with a Title and a Subtitle paragraph, styles, settings, a font table and a theme, and no
// numbering part. Tests of engine behaviour (numbering, sanitising, inserting, metadata) work on it because it gives
// them a few nodes and nothing else. It is not a template: `blank` is one empty paragraph, and the templates are
// tested in packages/engine/tests/templates.test.ts.
export function seededBlank(): Buffer {
  return fs.readFileSync(path.resolve(__dirname, '..', 'fixtures', 'minimal-two-paragraph.docx'));
}
