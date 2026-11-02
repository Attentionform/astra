import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';

const CLI_JS = path.resolve(__dirname, '../dist/cli.js');

function runCli(args: string[]): { status: number; stdout: string; stderr: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

describe('cli validate reports clear error when [Content_Types].xml is missing', () => {
  let dir: string;
  let file: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-no-ct-'));
    file = path.join(dir, 'no-content-types.docx');
    const zip = new AdmZip();
    zip.addFile('word/document.xml', Buffer.from('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>hello</w:t></w:r></w:p></w:body></w:document>', 'utf8'));
    zip.writeZip(file);
  });

  afterAll(() => {
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('prints clear message naming missing [Content_Types].xml on human output', () => {
    const res = runCli(['validate', file]);
    expect(res.status).toBe(1);
    expect(res.stdout).toContain('INVALID');
    expect(res.stdout).toContain('[Content_Types].xml is missing from the package root, so Word cannot find any part of this file.');
  });

  it('reports clear message in JSON output', () => {
    const res = runCli(['validate', file, '--json']);
    expect(res.status).toBe(1);
    const parsed = JSON.parse(res.stdout.trim().split('\n').filter((l) => l.startsWith('{')).pop() as string);
    expect(parsed.ok).toBe(false);
    expect(parsed.errors[0]).toContain('[Content_Types].xml is missing from the package root, so Word cannot find any part of this file.');
  });
});
