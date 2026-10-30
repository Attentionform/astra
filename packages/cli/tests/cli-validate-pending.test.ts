/**
 * `astra validate` checks the file as ASTra would save it (children put in schema
 * order first), so a file whose property children are out of order comes back
 * VALID. It now also lists the repairs a save would apply, as `inspect` and
 * `normalize --dry-run` do, so the verdict cannot be mistaken for "Word will
 * open this exactly as it is".
 */
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

const jsonLine = (stdout: string): Record<string, unknown> =>
  JSON.parse(stdout.trim().split('\n').filter((l) => l.startsWith('{')).pop() as string) as Record<string, unknown>;

describe('cli validate reports pending repairs', () => {
  let dir: string;
  let clean: string;
  let misordered: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-validate-'));
    const md = path.join(dir, 'in.md');
    clean = path.join(dir, 'clean.docx');
    misordered = path.join(dir, 'misordered.docx');
    fs.writeFileSync(md, '| A | B |\n| --- | --- |\n| 1 | 2 |\n\nBody text.\n');
    expect(runCli(['create', clean, '--from-md', md]).status).toBe(0);
    const zip = new AdmZip(clean);
    const xml = zip.readAsText('word/document.xml');
    const edited = xml.replace('</w:tblBorders><w:tblCellMar>', '</w:tblBorders><w:tblLook w:val="04A0"/><w:tblCellMar>');
    expect(edited).not.toBe(xml);
    zip.updateFile('word/document.xml', Buffer.from(edited, 'utf8'));
    zip.writeZip(misordered);
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  it('a clean file prints VALID alone and its JSON has no pendingNormalizations', () => {
    const human = runCli(['validate', clean]);
    expect(human.status).toBe(0);
    expect(human.stdout).toContain('VALID');
    expect(human.stdout).not.toContain('would-be normalize');
    const json = jsonLine(runCli(['--json', 'validate', clean]).stdout);
    expect(json).toEqual({ ok: true, file: clean, errors: [] });
  });

  it('a file with children out of order is VALID and names the repair a save would make', () => {
    const human = runCli(['validate', misordered]);
    expect(human.status).toBe(0);
    expect(human.stdout).toContain(`VALID ${misordered}`);
    expect(human.stdout).toMatch(/would-be normalize: Reordered w:tblPr into schema order/);
    const json = jsonLine(runCli(['--json', 'validate', misordered]).stdout);
    expect(json.ok).toBe(true);
    expect(json.errors).toEqual([]);
    expect(json.pendingNormalizations).toEqual([expect.stringMatching(/Reordered w:tblPr into schema order .*\[w:tblStyle,w:tblBorders,w:tblLook,w:tblCellMar\]/)]);
  });
});
