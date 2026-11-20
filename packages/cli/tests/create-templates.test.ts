/**
 * The built CLI makes every built-in template, in both paper sizes, from a different working directory,
 * and composes Markdown onto an untouched template. The CLI bundles the engine code, so this guards the
 * lookup of the template files in the installed engine package.
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
    it(`${name}: Letter by default, A4 with --page-size a4`, () => {
      const letter = path.join(dir, `${name}-letter.docx`);
      const a4 = path.join(dir, `${name}-a4.docx`);
      expect(cli(['create', letter, '--template', name]).status).toBe(0);
      expect(cli(['create', a4, '--template', name, '--page-size', 'a4']).status).toBe(0);
      expect(new AdmZip(fs.readFileSync(letter)).readAsText('word/document.xml')).toContain('w:w="12240"');
      expect(new AdmZip(fs.readFileSync(a4)).readAsText('word/document.xml')).toContain('w:w="11906"');
    });
  }

  it('an unknown template or page size exits 2 and names the valid values', () => {
    const t = cli(['create', path.join(dir, 'x.docx'), '--template', 'corporate']);
    expect(t.status).toBe(2);
    expect(t.stderr).toMatch(/blank\|report\|memo/);
    const s = cli(['create', path.join(dir, 'y.docx'), '--page-size', 'legal']);
    expect(s.status).toBe(2);
    expect(s.stderr).toMatch(/letter\|a4/);
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
