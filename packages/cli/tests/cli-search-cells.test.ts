/**
 * `astra search` lists the cell addresses that hold a match on a table hit,
 * and the same address is accepted by `astra batch` as an edit target.
 */
import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';

const CLI_JS = path.resolve(__dirname, '../dist/cli.js');

function runCli(args: string[]): { status: number; stdout: string; stderr: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

const SOURCE = '# Figures\n\nOpening line.\n\n| Region | Revenue |\n| --- | --- |\n| North | 120 |\n| South | 98 |\n';

describe('cli search cells', () => {
  let dir: string;
  let docx: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-cells-'));
    const md = path.join(dir, 'figures.md');
    docx = path.join(dir, 'figures.docx');
    fs.writeFileSync(md, SOURCE);
    expect(runCli(['create', docx, '--from-md', md]).status).toBe(0);
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  it('prints [cells: ...] on a table hit and nothing extra on a paragraph hit', () => {
    const table = runCli(['search', docx, 'South']);
    expect(table.status).toBe(0);
    expect(table.stdout).toMatch(/\[cells: \S+:r3c1\]/);
    const prose = runCli(['search', docx, 'Opening']);
    expect(prose.status).toBe(0);
    expect(prose.stdout).not.toContain('[cells:');
  });

  it('puts cells in the JSON hit', () => {
    const res = runCli(['--json', 'search', docx, 'North']);
    expect(res.status).toBe(0);
    const payload = JSON.parse(res.stdout.trim().split('\n').filter((l) => l.startsWith('{')).pop() as string) as {
      hits?: Array<{ id: string; type: string; cells?: string[] }>;
    };
    const hit = (payload.hits ?? []).find((h) => h.type === 'table');
    expect(hit?.cells).toEqual([`${hit?.id}:r2c1`]);
  });

  it('batch accepts the address from search as a target', () => {
    const found = runCli(['search', docx, 'South']);
    const address = /\[cells: (\S+?)\]/.exec(found.stdout)?.[1];
    expect(address).toBeDefined();
    const ops = path.join(dir, 'ops.json');
    fs.writeFileSync(ops, JSON.stringify([{ action: 'replace', target: address, text: 'South-East' }]));
    const applied = runCli(['batch', docx, ops]);
    expect(applied.status).toBe(0);
    expect(runCli(['search', docx, 'South-East']).stdout).toMatch(/1 match/);
  });
});
