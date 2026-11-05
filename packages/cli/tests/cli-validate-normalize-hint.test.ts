import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';

const CLI_JS = path.resolve(__dirname, '../dist/cli.js');
const HINT = 'Run `astra normalize` with `--out` pointing to a new file to attempt supported repairs, then validate the result again.';

function runCli(args: string[]) {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}
function jsonResult(args: string[]) {
  const res = runCli([...args, '--json']);
  return { ...res, data: JSON.parse(res.stdout.trim()) };
}
function invalidDocument(paragraphs: number): Buffer {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const xml = zip.readAsText('word/document.xml');
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const end = xml.lastIndexOf('<w:sectPr');
  const badP = '<w:p><w:r><w:t>misplaced</w:t></w:r><w:pPr><w:jc w:val="center"/></w:pPr></w:p>';
  zip.updateFile('word/document.xml', Buffer.from(xml.slice(0, start) + badP.repeat(paragraphs) + xml.slice(end), 'utf8'));
  return zip.toBuffer();
}

describe('cli validate normalize hint', () => {
  let dir: string;
  let file: string;
  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-val-hint-'));
    file = path.join(dir, 'report 2026.docx');
    fs.writeFileSync(file, invalidDocument(2));
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  it('suggests a separate output and revalidation without predicting repairs or changing the input', () => {
    const before = fs.readFileSync(file);
    const res = runCli(['validate', file]);
    expect(res.status).toBe(1);
    expect(res.stderr).toContain(`Remedy: ${HINT} Repair by hand what remains, for example in Word.`);
    expect(res.stderr).not.toMatch(/leaves \d+|\d+ changes|then validates/);
    expect(fs.readFileSync(file)).toEqual(before);
  });
  it('retains the additive JSON hint and invalid verdict', () => {
    const res = jsonResult(['validate', file]);
    expect(res.status).toBe(1);
    expect(res.data.ok).toBe(false);
    expect(res.data.normalizeHint).toBe(HINT);
  });
  it('still suggests normalize when the error list is truncated', () => {
    const res = jsonResult(['validate', file, '--max-errors', '1']);
    expect(res.status).toBe(1);
    expect(res.data.truncated).toBe(true);
    expect(res.data.errors).toHaveLength(1);
    expect(res.data.normalizeHint).toBe(HINT);
  });
  it('still suggests normalize above the default 500-error limit', () => {
    const many = path.join(dir, 'many.docx');
    fs.writeFileSync(many, invalidDocument(501));
    const res = jsonResult(['validate', many]);
    expect(res.status).toBe(1);
    expect(res.data.truncated).toBe(true);
    expect(res.data.errors).toHaveLength(500);
    expect(res.data.normalizeHint).toBe(HINT);
  });
  it('does not suggest normalize for valid documents', () => {
    const clean = path.join(dir, 'valid.docx');
    fs.writeFileSync(clean, DocxParser.createNewDocument('blank').rezip());
    expect(runCli(['validate', clean]).stdout).toContain('VALID ');
    const res = jsonResult(['validate', clean]);
    expect(res.status).toBe(0);
    expect(res.data).not.toHaveProperty('normalizeHint');
  });
  it.each(['missing types', 'missing file', 'corrupt zip', 'macro'])('preserves the specific %s error without suggesting normalize', (kind) => {
    const target = path.join(dir, `${kind}.docx`);
    if (kind === 'missing types') {
      const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
      zip.deleteFile('[Content_Types].xml');
      fs.writeFileSync(target, zip.toBuffer());
    } else if (kind === 'corrupt zip') fs.writeFileSync(target, 'not a zip');
    else if (kind === 'macro') {
      const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
      zip.addFile('word/vbaProject.bin', Buffer.from('fake macro bytes'));
      fs.writeFileSync(target, zip.toBuffer());
    }
    const human = runCli(['validate', target]);
    expect(human.status).toBe(1);
    expect(human.stderr).not.toContain(HINT);
    const res = jsonResult(['validate', target]);
    expect(res.status).toBe(1);
    expect(res.data).not.toHaveProperty('normalizeHint');
    if (kind === 'missing types') expect(res.data.errors.join(' ')).toContain('[Content_Types].xml');
    if (kind === 'macro') expect(human.stderr).toContain('macro');
    if (kind === 'missing file') expect(human.stderr).toMatch(/not found|ENOENT|no such file/i);
  });
  it.each(['--allow-macros', '--strip-macros'])('gives non-predictive advice after the explicit %s policy', (flag) => {
    const zip = new AdmZip(invalidDocument(2));
    zip.addFile('word/vbaProject.bin', Buffer.from('fake macro bytes'));
    const input = path.join(dir, 'allowed-macro.docm');
    fs.writeFileSync(input, zip.toBuffer());
    const res = jsonResult(['validate', input, flag]);
    expect(res.status).toBe(1);
    expect(res.data.normalizeHint).toBe(HINT);
  });
  it('gives honest advice for interacting faults and normalize reduces their errors', () => {
    const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
    let xml = zip.readAsText('word/document.xml');
    const ids = [...xml.matchAll(/w14:paraId="([^"]+)"/g)].map(m => m[1]);
    xml = xml.replace(`w14:paraId="${ids[1]}"`, `w14:paraId="${ids[0]}"`);
    xml = xml.replace('<w:pPr>', '<w:pPr><w:jc w:val="left"/></w:pPr><w:pPr>');
    const table = '<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcMar><w:bottom w:w="100" w:type="dxa"/><w:top w:w="100" w:type="dxa"/></w:tcMar></w:tcPr><w:p><w:r><w:t>cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
    xml = xml.replace('</w:body>', `${table}</w:body>`);
    zip.updateFile('word/document.xml', Buffer.from(xml));
    const input = path.join(dir, 'interacting.docx');
    const output = path.join(dir, 'interacting-fixed.docx');
    fs.writeFileSync(input, zip.toBuffer());
    const before = fs.readFileSync(input);
    const res = jsonResult(['validate', input]);
    expect(res.data.ok).toBe(false);
    expect(res.data.normalizeHint).toBe(HINT);
    const normalized = runCli(['normalize', input, '--out', output]);
    expect(normalized.status, normalized.stderr).toBe(0);
    const after = jsonResult(['validate', output]);
    expect(after.data.errors.length).toBeLessThan(res.data.errors.length);
    expect(fs.readFileSync(input)).toEqual(before);
  });
});
