import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import AdmZip from 'adm-zip';
import { BUILT_IN_GUIDE_PATH, DEFAULT_STYLE_GUIDE, resolveStyleGuide } from '@attentionform/astra-engine';

/**
 * The built-in guide that a document gets when none is named can also be asked for by name: `default` or
 * `Default Corporate`. Before, both were refused as unknown, and the help called the built-in guide "Default Corporate".
 */
describe('resolveStyleGuide with the built-in names', () => {
  it.each(['default', 'Default Corporate', 'DEFAULT', '  default corporate ', 'default-corporate', 'Default_Corporate'])('%j gives the built-in guide', (ref) => {
    const resolved = resolveStyleGuide(ref);
    expect(resolved.config).toEqual(DEFAULT_STYLE_GUIDE);
    expect(resolved.warnings).toEqual([]);
    expect(resolved.filePath).toBe(BUILT_IN_GUIDE_PATH);
  });

  it('gives a copy that a caller may change', () => {
    const first = resolveStyleGuide('default').config;
    first.tables!.headerFill = '000000';
    expect(resolveStyleGuide('default').config.tables!.headerFill).toBe(DEFAULT_STYLE_GUIDE.tables!.headerFill);
    expect(DEFAULT_STYLE_GUIDE.tables!.headerFill).not.toBe('000000');
  });

  it('a file named default.md is still a file, not the built-in guide', () => {
    expect(() => resolveStyleGuide('./default.md', { allowedDir: os.tmpdir() })).toThrow(/not found/);
  });

  it('an unknown name lists the bundled guides and "default"', () => {
    try {
      resolveStyleGuide('nope');
      throw new Error('expected a refusal');
    } catch (err) {
      const message = String((err as Error).message);
      expect(message).toMatch(/unknown style guide 'nope'/);
      expect(message).toMatch(/available: .*report.*default/);
      expect(message).toContain('"default" is the built-in Default Corporate guide');
    }
  });

  it('"legacy" is not a guide name for the engine', () => {
    expect(() => resolveStyleGuide('legacy')).toThrow(/unknown style guide 'legacy'/);
  });
});

describe('the command line', () => {
  const CLI_JS = path.resolve(__dirname, '..', '..', 'packages', 'cli', 'dist', 'cli.js');
  const run = (args: string[]) => {
    const r = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
    return { status: r.status ?? 99, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
  };
  let dir: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-guide-name-'));
    fs.writeFileSync(path.join(dir, 'in.md'), '# Title\n\n| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |\n');
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  const tableXml = (file: string): string => {
    const xml = new AdmZip(file).readAsText('word/document.xml');
    return xml.slice(xml.indexOf('<w:tbl>'), xml.indexOf('</w:tbl>'));
  };

  it('import with --style-guide "Default Corporate" and with none give the same table', () => {
    const none = path.join(dir, 'none.docx');
    const named = path.join(dir, 'named.docx');
    const plain = path.join(dir, 'plain.docx');
    expect(run(['import', path.join(dir, 'in.md'), '-o', none]).status).toBe(0);
    expect(run(['import', path.join(dir, 'in.md'), '-o', named, '--style-guide', 'Default Corporate']).status).toBe(0);
    expect(run(['import', path.join(dir, 'in.md'), '-o', plain, '--style-guide=default']).status).toBe(0);
    const strip = (x: string) => x.replace(/w14:paraId="[0-9A-F]+"/g, '');
    expect(strip(tableXml(named))).toBe(strip(tableXml(none)));
    expect(strip(tableXml(plain))).toBe(strip(tableXml(none)));
  });

  it('validate accepts the built-in name and refuses "legacy" as unknown', () => {
    const doc = path.join(dir, 'v.docx');
    expect(run(['import', path.join(dir, 'in.md'), '-o', doc]).status).toBe(0);
    expect(run(['validate', doc, '--style-guide', 'Default Corporate']).status).toBe(0);
    const legacy = run(['validate', doc, '--style-guide', 'legacy']);
    expect(legacy.status).not.toBe(0);
    expect(legacy.stderr).toContain("unknown style guide 'legacy'");
  });

  it('the help names the built-in guide and says where legacy applies', () => {
    for (const command of ['validate', 'batch']) {
      const help = run([command, '--help']).stdout.replace(/\s+/g, ' ');
      expect(help, command).toContain('`default` or `Default Corporate` is the built-in guide');
      expect(help, command).toContain('`legacy` is not a guide name here; it applies to create, import and compose only.');
    }
    for (const command of ['create', 'import']) {
      const help = run([command, '--help']).stdout.replace(/\s+/g, ' ');
      expect(help, command).toContain('`default` or `Default Corporate` is the built-in guide');
      expect(help, command).toContain("alerts take the guide's slate theme");
    }
  });
});
