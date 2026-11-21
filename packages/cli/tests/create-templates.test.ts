/**
 * The built CLI makes every built-in template, on Letter, A4 and a custom page, from a different working directory,
 * and composes Markdown onto an untouched template. The templates are built at run time, so this guards
 * the page flags and that the installed engine package builds them.
 */
import { describe, expect, it } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { AdmZip, TEMPLATE_NAMES } from '@attentionform/astra-engine';

const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');
const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-cli-templates-'));

function cli(args: string[]): { status: number; stdout: string; stderr: string } {
  const r = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8', cwd: dir });
  return { status: r.status ?? 99, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
}

describe('astra create --template', () => {
  for (const name of TEMPLATE_NAMES) {
    it(`${name}: Letter by default, A4 with --page-size A4`, () => {
      const letter = path.join(dir, `${name}-letter.docx`);
      const a4 = path.join(dir, `${name}-a4.docx`);
      expect(cli(['create', letter, '--template', name]).status).toBe(0);
      expect(cli(['create', a4, '--template', name, '--page-size', 'A4']).status).toBe(0);
      expect(new AdmZip(fs.readFileSync(letter)).readAsText('word/document.xml')).toContain('w:w="12240"');
      expect(new AdmZip(fs.readFileSync(a4)).readAsText('word/document.xml')).toContain('w:w="11906"');
    });
  }

  it('an unknown template exits 2 and names the valid values', () => {
    const t = cli(['create', path.join(dir, 'x.docx'), '--template', 'corporate']);
    expect(t.status).toBe(2);
    expect(t.stderr).toMatch(/blank\|report\|memo/);
  });

  it('--page-size takes any named paper or a size, --landscape turns it, --margins sets them', () => {
    const out = path.join(dir, 'custom.docx');
    const r = cli(['create', out, '--template', 'memo', '--page-size', '148x210mm', '--landscape', '--margins', '15mm']);
    expect(r.status, r.stderr).toBe(0);
    const xml = new AdmZip(fs.readFileSync(out)).readAsText('word/document.xml');
    expect(xml).toContain('w:w="11906" w:h="8391"');
    expect(xml).toContain('w:orient="landscape"');
    expect(xml).toContain('w:top="850" w:right="850" w:bottom="850" w:left="850"');
    const legal = path.join(dir, 'legal.docx');
    expect(cli(['create', legal, '--page-size', 'Legal']).status).toBe(0);
    expect(new AdmZip(fs.readFileSync(legal)).readAsText('word/document.xml')).toContain('w:h="20160"');
  });

  it('a size or margins that cannot be laid out exit with a message that says what to use', () => {
    const big = cli(['create', path.join(dir, 'big.docx'), '--page-size', '30x40in']);
    expect(big.status).not.toBe(0);
    expect(`${big.stderr}${big.stdout}`).toMatch(/outside 1 in to 22 in/);
    const narrow = cli(['create', path.join(dir, 'narrow.docx'), '--page-size', 'A5', '--margins', '70mm']);
    expect(narrow.status).not.toBe(0);
    expect(`${narrow.stderr}${narrow.stdout}`).toMatch(/text width/);
    const both = cli(['create', path.join(dir, 'both.docx'), '--landscape', '--portrait']);
    expect(both.status).toBe(2);
  });

  it('--from-md replaces the untouched placeholder body and fills the title', () => {
    const md = path.join(dir, 'in.md');
    fs.writeFileSync(md, '# Quarterly Review\n\nFirst paragraph.\n');
    const out = path.join(dir, 'composed.docx');
    expect(cli(['create', out, '--template', 'report', '--from-md', md]).status).toBe(0);
    const xml = new AdmZip(fs.readFileSync(out)).readAsText('word/document.xml');
    expect(xml).toContain('First paragraph.');
    expect(xml).toContain('Quarterly Review');
    expect(xml).not.toContain('Summarise the purpose');
  });
});
