import { describe, it, expect, beforeAll, afterAll } from 'vitest';
import { spawnSync } from 'child_process';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { BatchOperationsSchema } from '@attentionform/astra-engine';
import { batchSchemaPayload } from '../src/cli/batch-cmd';

/**
 * `astra batch --schema` shows one example per kind of operation. Every action that the operation schema
 * accepts has at least one example, every example passes the schema gate, and every example that does not
 * depend on an earlier operation runs on a real document.
 */
const REPO_ROOT = path.resolve(__dirname, '..', '..', '..');
const CLI_JS = path.join(REPO_ROOT, 'packages', 'cli', 'dist', 'cli.js');

// A 1x1 PNG.
const PNG = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

/** Examples that name a comment or chart made by an earlier operation, so they are only checked against the schema. */
const NEEDS_EARLIER_OPERATION = new Set(['comment_reply', 'chart_update']);

type Example = Record<string, unknown>;

function examples(): Array<[string, Example]> {
  return Object.entries(batchSchemaPayload()).filter(([k]) => k !== '$comment') as Array<[string, Example]>;
}

function run(args: string[], cwd?: string): { status: number; stdout: string; stderr: string } {
  const res = spawnSync(process.execPath, [CLI_JS, ...args], { encoding: 'utf8', cwd });
  return { status: res.status ?? 99, stdout: res.stdout ?? '', stderr: res.stderr ?? '' };
}

function substitute(value: unknown, ids: Record<string, string>): unknown {
  if (typeof value === 'string') return ids[value] ?? value;
  if (Array.isArray(value)) return value.map((v) => substitute(v, ids));
  if (value && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, substitute(v, ids)]));
  }
  return value;
}

describe('batch --schema examples', () => {
  it('has an example for every action of the operation schema', () => {
    const element = (BatchOperationsSchema as unknown as { shape: { operations: { element: { options: Array<{ shape: { action: { def: { values: string[] } } } }> } } } }).shape.operations.element;
    const actions = element.options.map((o) => o.shape.action.def.values[0]);
    expect(actions.length).toBeGreaterThan(20);
    const shown = new Set(examples().map(([, ex]) => ex.action as string));
    // `clearDoc` is a command-line extension that the engine schema does not list.
    const missing = actions.filter((a) => !shown.has(a));
    expect(missing).toEqual([]);
  });

  it('shows each table operation that people reach for', () => {
    const tableOps = examples()
      .filter(([, ex]) => ex.action === 'table')
      .map(([, ex]) => ex.operation);
    for (const op of ['updateCell', 'addRow', 'deleteRow', 'addColumn', 'mergeCells']) expect(tableOps).toContain(op);
  });

  it('shows the paragraph spacing keys in a style example', () => {
    const styleKeys = new Set(examples().filter(([, ex]) => ex.action === 'style').flatMap(([, ex]) => Object.keys(ex)));
    for (const key of ['spacingBefore', 'spacingAfter', 'lineSpacing', 'indentLeft', 'alignment']) expect(styleKeys).toContain(key);
  });

  it.each(examples().filter(([k]) => !k.endsWith('_cli')))('%s passes the schema gate', (_name, ex) => {
    const parsed = BatchOperationsSchema.safeParse({ operations: [ex] });
    expect(parsed.success, parsed.success ? '' : JSON.stringify(parsed.error.issues)).toBe(true);
  });

  describe('on a document', () => {
    let dir: string;
    let ids: Record<string, string>;

    beforeAll(() => {
      dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-schema-examples-'));
      fs.writeFileSync(path.join(dir, 'logo.png'), PNG);
      fs.writeFileSync(path.join(dir, 'seed.md'), '# Title\n\nFirst paragraph of the document.\n\n| A | B |\n|---|---|\n| 1 | 2 |\n| 3 | 4 |\n');
      const made = run(['create', path.join(dir, 'seed.docx'), '--from-md', path.join(dir, 'seed.md')], dir);
      expect(made.status, made.stderr).toBe(0);
      const listing = run(['inspect', path.join(dir, 'seed.docx')], dir).stdout.split('\n');
      const rowId = (needle: RegExp): string => {
        const row = listing.find((l) => needle.test(l));
        expect(row, `a row matching ${needle}`).toBeDefined();
        return row!.split('|')[0].trim();
      };
      ids = {
        '<nodeId>': rowId(/First paragraph of the document/),
        '<otherNodeId>': rowId(/\| Title \|/),
        '<tableNodeId>': rowId(/\[TBL\]|Table 3x2|Table \d+x\d+/),
      };
    });

    afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));

    it.each(examples().filter(([k]) => !NEEDS_EARLIER_OPERATION.has(k)))('%s runs', (name, ex) => {
      const work = path.join(dir, `${name}.docx`);
      fs.copyFileSync(path.join(dir, 'seed.docx'), work);
      const ops = path.join(dir, `${name}.json`);
      fs.writeFileSync(ops, JSON.stringify({ operations: [substitute(ex, ids)] }));
      const r = run(['batch', work, ops, '--allowed-dir', dir, '--force'], dir);
      expect(r.status, `${r.stdout}\n${r.stderr}`).toBe(0);
    });
  });
});
