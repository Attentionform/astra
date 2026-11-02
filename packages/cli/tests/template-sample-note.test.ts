import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { templateSampleNote } from '../src/cli/create-cmd';

/**
 * `--template report` writes a whole sample report with invented figures, and composed Markdown is added after
 * the sample text of `report` and `memo`. `create` and `import` say so on stderr and in the JSON, and the help does too.
 */
const CLI_JS = path.resolve(__dirname, '..', 'dist', 'cli.js');
const run = (args: string[]) => {
  const r = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8' });
  return { status: r.status ?? 99, stdout: r.stdout ?? '', stderr: r.stderr ?? '' };
};

describe('templateSampleNote', () => {
  it('names the sample content of report and memo, and nothing for blank', () => {
    expect(templateSampleNote('report', false)).toContain('sample content (a Q3 financial report with invented figures)');
    expect(templateSampleNote('memo', false)).toContain('a heading and a placeholder line');
    expect(templateSampleNote('blank', false)).toBeUndefined();
    expect(templateSampleNote('blank', true)).toBeUndefined();
  });

  it('adds that composed Markdown comes after the sample text only when something is composed', () => {
    expect(templateSampleNote('report', false)).not.toContain('added after that text');
    expect(templateSampleNote('report', true)).toContain('added after that text, not in place of it');
    expect(templateSampleNote('memo', true)).toContain('added after that text, not in place of it');
  });
});

describe('create and import with a template', () => {
  let dir: string;
  let md: string;

  beforeAll(() => {
    dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-template-note-'));
    md = path.join(dir, 'own.md');
    fs.writeFileSync(md, '# My own title\n\nMy own text.\n');
  });
  afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

  const rows = (file: string): string[] =>
    run(['inspect', file]).stdout.split('\n').filter((l) => /^[0-9A-F]{8} \|/.test(l)).map((l) => l.split('|')[5].trim());

  it('create --template report prints the note on stderr', () => {
    const r = run(['create', path.join(dir, 'a.docx'), '--template', 'report']);
    expect(r.status).toBe(0);
    expect(r.stdout).toMatch(/^CREATED .*\(template=report\)/);
    expect(r.stderr).toContain('NOTE: --template report includes sample content');
    expect(r.stderr).not.toContain('added after that text');
  });

  it('the note is true: composed Markdown lands after the sample text of report and memo, and replaces the blank prompt', () => {
    const report = path.join(dir, 'b.docx');
    const r = run(['create', report, '--template', 'report', '--from-md', md]);
    expect(r.stderr).toContain('added after that text, not in place of it');
    const reportRows = rows(report);
    expect(reportRows.indexOf('Q3 Executive Financial Performance Report')).toBeLessThan(reportRows.indexOf('My own title'));
    expect(reportRows).toContain('My own text.');

    const memo = path.join(dir, 'c.docx');
    expect(run(['create', memo, '--template', 'memo', '--from-md', md]).stderr).toContain('--template memo includes a heading');
    expect(rows(memo).indexOf('MEMORANDUM')).toBeLessThan(rows(memo).indexOf('My own title'));

    const blank = path.join(dir, 'd.docx');
    const rb = run(['create', blank, '--template', 'blank', '--from-md', md]);
    expect(rb.stderr).not.toContain('NOTE');
    expect(rows(blank)).not.toContain('Blank Document');
  });

  it('create --template blank and the default print no note', () => {
    expect(run(['create', path.join(dir, 'e.docx'), '--template', 'blank']).stderr).not.toContain('NOTE');
    expect(run(['create', path.join(dir, 'f.docx')]).stderr).not.toContain('NOTE');
  });

  it('import --template report prints the note, with the composed sentence', () => {
    const r = run(['import', md, '-o', path.join(dir, 'g.docx'), '--template', 'report']);
    expect(r.status).toBe(0);
    expect(r.stderr).toContain('NOTE: --template report includes sample content');
    expect(r.stderr).toContain('added after that text');
  });

  it('--json carries templateNote and keeps stderr free of the note', () => {
    const r = run(['--json', 'create', path.join(dir, 'h.docx'), '--template', 'report']);
    expect(JSON.parse(r.stdout).templateNote).toContain('sample content');
    expect(r.stderr).not.toContain('NOTE');
    const blank = run(['--json', 'create', path.join(dir, 'i.docx')]);
    expect(JSON.parse(blank.stdout).templateNote).toBeUndefined();
  });

  it('the help of create and import describes the templates', () => {
    const create = run(['create', '--help']).stdout.replace(/\s+/g, ' ');
    expect(create).toContain('report holds a sample report with invented Q3 figures');
    expect(create).toContain('the Markdown is added after the sample text of report and memo, not in place of it');
    const imp = run(['import', '--help']).stdout.replace(/\s+/g, ' ');
    expect(imp).toContain('report holds a sample report with invented Q3 figures');
    expect(imp).toContain('blank is the only template whose prompt lines the manuscript replaces');
  });
});
