import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { DocxParser, convertDocxToPdf } from '@attentionform/astra-engine';

/**
 * Writes a fake LibreOffice backend: honors the `--outdir <dir>` + `<docx>`
 * argv contract and copies the input to the expected `<base>.pdf` name.
 */
function fakeLibreOffice(): string {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-fake-lo-'));
  const stub = path.join(dir, 'soffice');
  fs.writeFileSync(
    stub,
    [
      '#!/bin/sh',
      'outdir=""',
      'docx=""',
      'prev=""',
      'for a in "$@"; do',
      '  if [ "$prev" = "--outdir" ]; then outdir="$a"; fi',
      '  case "$a" in *.docx) docx="$a";; esac',
      '  prev="$a"',
      'done',
      'base=$(basename "$docx" .docx)',
      'cp "$docx" "$outdir/$base.pdf"',
      '',
    ].join('\n')
  );
  fs.chmodSync(stub, 0o755);
  return stub;
}

describe('pdf output permissions', () => {
  it('lands 0644 like the sibling exporters (not 0600 LO-temp inheritance)', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-pdfmode-'));
    const buf = DocxParser.createNewDocument('blank').rezip();
    const out = path.join(dir, 'rich.pdf');
    await convertDocxToPdf(buf, out, { binary: fakeLibreOffice() });
    expect(fs.existsSync(out)).toBe(true);
    expect(fs.statSync(out).mode & 0o777).toBe(0o644);
  });
});
