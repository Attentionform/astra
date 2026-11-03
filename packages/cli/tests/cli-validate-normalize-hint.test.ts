import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';

const CLI_JS = path.resolve(__dirname, '../dist/cli.js');

function runCli(args: string[]): { status: number; stdout: string; stderr: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

describe('cli validate normalize hint', () => {
  let dir: string;
  let file: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-val-hint-'));
    file = path.join(dir, 'misplaced.docx');
    const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
    const xml = zip.readAsText('word/document.xml');
    const start = xml.indexOf('<w:body>') + '<w:body>'.length;
    const end = xml.lastIndexOf('<w:sectPr');
    const badP = '<w:p><w:r><w:t>misplaced</w:t></w:r><w:pPr><w:jc w:val="center"/></w:pPr></w:p>';
    zip.updateFile('word/document.xml', Buffer.from(xml.slice(0, start) + badP + xml.slice(end), 'utf8'));
    zip.writeZip(file);
  });

  afterAll(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('prints normalize hint in stderr remedy on failed validation', () => {
    const res = runCli(['validate', file]);
    expect(res.status).toBe(1);
    expect(res.stderr).toContain('failed OOXML validation');
    expect(res.stderr).toContain('Remedy: repair the document in Word and retry.');
    expect(res.stderr).toContain(`Run \`astra normalize ${file}\`: it can repair some of these errors`);
  });

  it('includes normalizeHint in JSON output', () => {
    const res = runCli(['validate', file, '--json']);
    expect(res.status).toBe(1);
    const parsed = JSON.parse(res.stdout.trim().split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(parsed.ok).toBe(false);
    expect(parsed.normalizeHint).toBeDefined();
    expect(parsed.normalizeHint).toContain(`Run \`astra normalize ${file}\`: it can repair some of these errors`);
  });
});
