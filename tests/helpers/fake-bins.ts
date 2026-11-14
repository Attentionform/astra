import * as fs from 'fs';
import * as path from 'path';

const PDF_SOURCE = `
function pdf(pages) {
  const objs = []; const kids = [];
  objs.push('<</Type/Catalog/Pages 2 0 R>>');
  for (let i = 0; i < pages; i++) kids.push((3 + i * 2) + ' 0 R');
  objs.push('<</Type/Pages/Kids[' + kids.join(' ') + ']/Count ' + pages + '>>');
  for (let i = 0; i < pages; i++) { objs.push('<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 100]/Contents ' + (4 + i * 2) + ' 0 R>>'); objs.push('<</Length 0>>\\nstream\\n\\nendstream'); }
  let out = '%PDF-1.4\\n'; const offs = [];
  objs.forEach((o, i) => { offs.push(out.length); out += (i + 1) + ' 0 obj\\n' + o + '\\nendobj\\n'; });
  const x = out.length; out += 'xref\\n0 ' + (objs.length + 1) + '\\n0000000000 65535 f \\n';
  for (const o of offs) out += String(o).padStart(10, '0') + ' 00000 n \\n';
  return out + 'trailer\\n<</Size ' + (objs.length + 1) + '/Root 1 0 R>>\\nstartxref\\n' + x + '\\n%%EOF\\n';
}`;

/** Writes an executable Node script (shebang = the current node) and returns its path. POSIX only. */
export function makeFakeBinary(dir: string, name: string, body: string): string {
  const file = path.join(dir, name);
  fs.writeFileSync(file, `#!${process.execPath}\nconst fs = require('fs'); const path = require('path');\n${PDF_SOURCE}\n${body}\n`);
  fs.chmodSync(file, 0o755);
  return file;
}

/**
 * A stand-in for soffice: writes a PDF of FAKE_PDF_PAGES pages next to the requested output dir, and records what
 * it was given (arguments and a copy of the input) when FAKE_LOG / FAKE_COPY_INPUT are set.
 * FAKE_SOFFICE_FAIL=1 makes it exit 1; FAKE_SOFFICE_FAIL_MARKUP=1 only when the margin-comments option is passed.
 * `--version` prints FAKE_SOFFICE_VERSION (default a current version; empty prints nothing), on stderr when
 * FAKE_SOFFICE_VERSION_STREAM=stderr, exits FAKE_SOFFICE_VERSION_EXIT, hangs when FAKE_SOFFICE_VERSION_HANG=1.
 */
export function makeFakeSoffice(dir: string): string {
  return makeFakeBinary(
    dir,
    'soffice',
    `
const args = process.argv.slice(2);
if (args.includes('--version')) {
  if (process.env.FAKE_SOFFICE_VERSION_HANG === '1') setInterval(() => undefined, 1000);
  else {
    if (process.env.FAKE_LOG) fs.appendFileSync(process.env.FAKE_LOG, JSON.stringify({ bin: 'soffice', version: true }) + '\\n');
    const text = process.env.FAKE_SOFFICE_VERSION === undefined ? 'LibreOffice 26.8.0.3 bce0998afefdbc355585ca324285661a2170ba77' : process.env.FAKE_SOFFICE_VERSION;
    if (process.env.FAKE_SOFFICE_VERSION_STREAM === 'stderr') process.stderr.write(text + '\\n'); else if (text) process.stdout.write(text + '\\n');
    process.exit(Number(process.env.FAKE_SOFFICE_VERSION_EXIT || 0));
  }
} else {
const outdir = args[args.indexOf('--outdir') + 1];
const input = args.filter((a) => a.endsWith('.docx')).pop();
if (process.env.FAKE_LOG) fs.appendFileSync(process.env.FAKE_LOG, JSON.stringify({ bin: 'soffice', filter: args[args.indexOf('--convert-to') + 1], input: path.basename(input) }) + '\\n');
if (process.env.FAKE_COPY_INPUT) fs.copyFileSync(input, process.env.FAKE_COPY_INPUT + '.' + (process.env.FAKE_COPY_SEQ || '0'));
if (process.env.FAKE_SOFFICE_FAIL === '1') { process.stderr.write('fake soffice failure'); process.exit(1); }
if (process.env.FAKE_SOFFICE_FAIL_MARKUP === '1' && String(args[args.indexOf('--convert-to') + 1]).includes('ExportNotesInMargin')) { process.stderr.write('fake soffice markup failure'); process.exit(1); }
fs.writeFileSync(path.join(outdir, path.basename(input, '.docx') + '.pdf'), pdf(Number(process.env.FAKE_PDF_PAGES || 1)));
}
`
  );
}

/** A stand-in for pdftoppm: writes <prefix>-N.png|jpg for FAKE_PDF_PAGES pages (zero-padded like Poppler for 10+). */
export function makeFakePdftoppm(dir: string): string {
  return makeFakeBinary(
    dir,
    'pdftoppm',
    `
const args = process.argv.slice(2);
const prefix = args[args.length - 1];
const jpeg = args.includes('-jpeg');
const n = Number(process.env.FAKE_PDF_PAGES || 1);
if (process.env.FAKE_LOG) fs.appendFileSync(process.env.FAKE_LOG, JSON.stringify({ bin: 'pdftoppm', args: args.slice(0, -2) }) + '\\n');
if (process.env.FAKE_PDFTOPPM_FAIL === '1') { process.stderr.write('fake pdftoppm failure'); process.exit(1); }
for (let i = 1; i <= n; i++) {
  const num = n >= 10 ? String(i).padStart(2, '0') : String(i);
  fs.writeFileSync(prefix + '-' + num + (jpeg ? '.jpg' : '.png'), 'fake-poppler-' + i);
}
`
  );
}
