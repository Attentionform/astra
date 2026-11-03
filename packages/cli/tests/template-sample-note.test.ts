import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { templateSampleNote } from '../src/cli/create-cmd';

/**
 * Pure style templates: built-in templates (`report`, `memo`, `blank`) provide pure style and layout
 * definitions without hardcoded dummy text. Composed Markdown cleanly replaces the fresh template placeholder.
 */
const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');
const run = (args: string[]) => {
  const r = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: r.status ?? 99, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
};

describe('templateSampleNote', () => {
  it('returns undefined for pure style templates', () => {
    expect(templateSampleNote('report', false)).toBeUndefined();
    expect(templateSampleNote('report', true)).toBeUndefined();
    expect(templateSampleNote('memo', false)).toBeUndefined();
    expect(templateSampleNote('memo', true)).toBeUndefined();
    expect(templateSampleNote('blank', false)).toBeUndefined();
    expect(templateSampleNote('blank', true)).toBeUndefined();
  });
});

describe('create and import with pure style templates', () => {
  let dir: string;
  let md: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-template-clean-'));
    md = path.join(dir, 'own.md');
    fs.writeFileSync(md, '# My own title\n\nMy own text [@c1].\n\n[@c1]: Reference 1\n');
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  const rows = (file: string): string[] =>
    run(['inspect', file]).stdout.split('\n').filter((l) => /^[0-9A-F]{8} \|/.test(l)).map((l) => l.split('|')[5].trim());

  it('create --template report creates a clean document without sample text', () => {
    const r = run(['create', path.join(dir, 'a.docx'), '--template', 'report']);
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/^CREATED .*\(template=report\)/);
    expect(r.stderr).not.toContain('NOTE');
    expect(rows(path.join(dir, 'a.docx'))).not.toContain('Q3 Executive Financial Performance Report');
  });

  it('composed Markdown replaces the template placeholder cleanly across report, memo, and blank', () => {
    const report = path.join(dir, 'b.docx');
    const r = run(['create', report, '--template', 'report', '--from-md', md]);
    expect(r.stderr).not.toContain('NOTE');
    const reportRows = rows(report);
    expect(reportRows[0]).toBe('My own title');
    expect(reportRows).not.toContain('Q3 Executive Financial Performance Report');
    expect(reportRows).toContain('My own text [1].');
    expect(reportRows).toContain('[1] Reference 1');

    const memo = path.join(dir, 'c.docx');
    const rm = run(['create', memo, '--template', 'memo', '--from-md', md]);
    expect(rm.stderr).not.toContain('NOTE');
    const memoRows = rows(memo);
    expect(memoRows[0]).toBe('My own title');
    expect(memoRows).not.toContain('MEMORANDUM');

    const blank = path.join(dir, 'd.docx');
    const rb = run(['create', blank, '--template', 'blank', '--from-md', md]);
    expect(rb.stderr).not.toContain('NOTE');
    const blankRows = rows(blank);
    expect(blankRows[0]).toBe('My own title');
    expect(blankRows).not.toContain('Blank Document');
  });

  it('create --template blank and the default print no note', () => {
    expect(run(['create', path.join(dir, 'e.docx'), '--template', 'blank']).stderr).not.toContain('NOTE');
    expect(run(['create', path.join(dir, 'f.docx')]).stderr).not.toContain('NOTE');
  });

  it('import --template report replaces seed without note and starts at line 1', () => {
    const r = run(['import', md, '-o', path.join(dir, 'g.docx'), '--template', 'report']);
    expect(r.status).toBe(0);
    expect(r.stderr).not.toContain('NOTE');
    const gRows = rows(path.join(dir, 'g.docx'));
    expect(gRows[0]).toBe('My own title');
    expect(gRows).not.toContain('Q3 Executive');
  });

  it('--json carries no templateNote for pure style templates', () => {
    const r = run(['--json', 'create', path.join(dir, 'h.docx'), '--template', 'report']);
    expect(JSON.parse(r.stdout).templateNote).toBeUndefined();
    expect(r.stderr).not.toContain('NOTE');
    const blank = run(['--json', 'create', path.join(dir, 'i.docx')]);
    expect(JSON.parse(blank.stdout).templateNote).toBeUndefined();
  });

  it('the help of create and import describes pure style templates', () => {
    const create = run(['create', '--help']).stdout.replace(/\s+/g, ' ');
    expect(create).toContain('pure style and layout templates');
    const imp = run(['import', '--help']).stdout.replace(/\s+/g, ' ');
    expect(imp).toContain('Fresh template placeholders are cleanly replaced on compose/import.');
  });
});
