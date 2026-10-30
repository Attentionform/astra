import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run, resolveInspectWidth, INSPECT_DEFAULT_WRAP_COLUMNS } from '../src/cli';
import { DocxParser, ToonGenerator, unwrapToonText } from '@attentionform/astra-engine';
import { withLoSlot } from '../../../tests/helpers/lo-slot';

// Slot + budget note: the --visual runs below spawn headless LibreOffice, so
// they hold the suite-wide LO slot (see tests/helpers/lo-slot.ts) and carry
// 300s for queue + own run. Telemetry-only asserts stay on the fast path.

const LONG =
  'Alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike ' +
  'november oscar papa quebec romeo sierra tango uniform victor whiskey xray yankee ' +
  'zulu extra words for wrap testing here';

const tmpFiles: string[] = [];
let outLines: string[];
let errLines: string[];

function tmpPath(prefix: string): string {
  const file = path.join(os.tmpdir(), `${prefix}-${process.pid}-${Date.now()}-${tmpFiles.length}.docx`);
  tmpFiles.push(file);
  return file;
}

async function makeDoc(): Promise<string> {
  const file = tmpPath('astra-inspect-wrap');
  fs.writeFileSync(file, DocxParser.createNewDocument('blank').rezip());
  await expect(run(['batch', file, '--append', LONG])).resolves.toBe(0);
  return file;
}

async function inspect(file: string, extra: string[]): Promise<number> {
  outLines = [];
  errLines = [];
  return run(['inspect', file, ...extra]);
}

/** Flatten console.log entries (one print call can hold many lines) and drop CLI footers. */
function toonLines(): string[] {
  return outLines
    .flatMap((s) => s.split('\n'))
    .filter((l) => {
      const t = l.trim();
      if (t === '') return false;
      if (/^(rows|words)=/.test(t)) return false;
      if (t.startsWith('Flags:') || t.startsWith('Units:')) return false;
      return true;
    });
}

/** Parse the last JSON payload from mocked console.log output. */
function lastJson(): any {
  const candidate = [...outLines].reverse().find((s) => s.trimStart().startsWith('{'));
  return JSON.parse((candidate ?? outLines.join('\n')).trim());
}

beforeEach(() => {
  outLines = [];
  errLines = [];
  vi.spyOn(console, 'log').mockImplementation((...a: unknown[]) => {
    outLines.push(a.map(String).join(' '));
  });
  vi.spyOn(console, 'error').mockImplementation((...a: unknown[]) => {
    errLines.push(a.map(String).join(' '));
  });
});

afterEach(() => {
  vi.restoreAllMocks();
  while (tmpFiles.length > 0) {
    const f = tmpFiles.pop() as string;
    try {
      fs.unlinkSync(f);
    } catch {
      /* best-effort */
    }
  }
});

describe('inspect wrap width resolution', () => {
  it('explicit width wins, wrap detects columns, default is 80', () => {
    expect(INSPECT_DEFAULT_WRAP_COLUMNS).toBe(80);
    expect(resolveInspectWidth({})).toBeUndefined();
    expect(resolveInspectWidth({ width: 40 })).toBe(40);
    expect(resolveInspectWidth({ width: 50, wrap: true, columns: 40 })).toBe(50);
    expect(resolveInspectWidth({ wrap: true, columns: 40 })).toBe(40);
    expect(resolveInspectWidth({ wrap: true, columns: 0 })).toBe(80);
    expect(resolveInspectWidth({ wrap: true, columns: Number.NaN })).toBe(80);
  });
});

describe('inspect --width/--wrap', () => {
  it('default inspect wraps nothing: every Text row is one line', async () => {
    const file = await makeDoc();
    await expect(inspect(file, [])).resolves.toBe(0);
    const flat = outLines.flatMap((s) => s.split('\n'));
    expect(flat.some((l) => l.includes(' | '))).toBe(true);
    for (const l of flat) {
      if (l.includes(' | ')) expect(l.startsWith(' ')).toBe(false);
    }
  });

  it('--width 40 wraps Text with hanging indent and unwraps to the default rows', async () => {
    const file = await makeDoc();
    await expect(inspect(file, [])).resolves.toBe(0);
    const def = toonLines();
    await expect(inspect(file, ['--width', '40'])).resolves.toBe(0);
    const wrapped = toonLines();
    const flat = wrapped;
    expect(flat.some((l) => l.startsWith(' '))).toBe(true);
    const groups: string[][] = [];
    for (const l of flat) {
      if (l.startsWith(' ') && groups.length > 0) groups[groups.length - 1].push(l);
      else groups.push([l]);
    }
    for (const g of groups) {
      for (const c of g.slice(1, -1)) expect(c).not.toContain(' | ');
    }
    expect(unwrapToonText(`${wrapped.join('\n')}\n`)).toBe(`${def.join('\n')}\n`);
  });

  it('--width=40 equals form works', async () => {
    const file = await makeDoc();
    await expect(inspect(file, ['--width=40'])).resolves.toBe(0);
    expect(outLines.flatMap((s) => s.split('\n')).some((l) => l.startsWith(' '))).toBe(true);
  });

  it('--wrap is accepted (piped stdout falls back to 80)', async () => {
    const file = await makeDoc();
    await expect(inspect(file, ['--wrap'])).resolves.toBe(0);
  });

  it('bad --width values are usage errors (exit 2)', async () => {
    const file = await makeDoc();
    await expect(inspect(file, ['--width', 'abc'])).resolves.toBe(2);
    await expect(inspect(file, ['--width', '0'])).resolves.toBe(2);
    await expect(inspect(file, ['--width', '10001'])).resolves.toBe(2);
  });
});

describe('inspect --visual restored', () => {
  it(
    '--visual exits 0 and prints telemetry (not unknown-flag 2)',
    async () => {
      const file = await makeDoc();
      await expect(withLoSlot(() => inspect(file, ['--visual']))).resolves.toBe(0);
      const flat = outLines.flatMap((s) => s.split('\n'));
      expect(flat.some((l) => l.includes('visual:'))).toBe(true);
    },
    300000,
  );

  it(
    '--visual combines with --width',
    async () => {
      const file = await makeDoc();
      await expect(withLoSlot(() => inspect(file, ['--visual', '--width', '40']))).resolves.toBe(0);
      expect(outLines.join('\n')).toContain('visual:');
    },
    300000,
  );

  it(
    '--visual under --json carries the visual payload',
    async () => {
      const file = await makeDoc();
      outLines = [];
      errLines = [];
      await expect(withLoSlot(() => run(['--json', 'inspect', file, '--visual']))).resolves.toBe(0);
      const payload = lastJson();
      expect(payload.ok).toBe(true);
      expect(payload.visual).toBeDefined();
      expect(payload.visual.telemetry).toBeDefined();
    },
    300000,
  );
});

describe('inspect --json purity under wrap', () => {
  it('--json toon stays unwrapped under --width (human stdout wraps)', async () => {
    const file = await makeDoc();
    await expect(inspect(file, ['--width', '40'])).resolves.toBe(0);
    const humanHasIndent = outLines.flatMap((s) => s.split('\n')).some((l) => l.startsWith(' '));
    expect(humanHasIndent).toBe(true);
    outLines = [];
    errLines = [];
    await expect(run(['--json', 'inspect', file, '--width', '40'])).resolves.toBe(0);
    const wrapped = lastJson();
    for (const line of String(wrapped.toon).split('\n')) {
      expect(line.startsWith(' ')).toBe(false);
    }
  });

  it('--json toon identical with and without --width', async () => {
    const file = await makeDoc();
    outLines = [];
    errLines = [];
    await expect(run(['--json', 'inspect', file])).resolves.toBe(0);
    const plain = lastJson().toon;
    outLines = [];
    errLines = [];
    await expect(run(['--json', 'inspect', file, '--width', '40'])).resolves.toBe(0);
    const wrapped = lastJson().toon;
    expect(wrapped).toBe(plain);
  });

  it('help documents human-only wrap and whitespace unwrap limit', async () => {
    outLines = [];
    errLines = [];
    await expect(run(['inspect', '--help'])).resolves.toBe(0);
    const text = outLines.join('\n');
    expect(text).toContain('--json toon stays unwrapped');
    expect(text).toContain('multi-space/tab');
  });
});

describe('inspect wrap round-trip limits', () => {
  function renderWithText(text: string): { gen: ToonGenerator; def: string; wrapped: string } {
    const gen = new ToonGenerator();
    const table = gen.generateToon({
      sections: [
        {
          id: 's',
          version: 1,
          checksum: 'c',
          heading: { id: 'h', text: 'G', level: 1 },
          nodes: [{ id: 'n1', version: 1, checksum: 'c1', type: 'paragraph', metadata: { text } }],
        },
      ],
    } as any);
    const def = gen.renderToonText(table);
    return { gen, def, wrapped: gen.renderToonText(table, { width: 40 }) };
  }

  it('single-space text unwraps byte-identically', () => {
    const { wrapped, def } = renderWithText(LONG);
    expect(wrapped).not.toBe(def);
    expect(unwrapToonText(wrapped)).toBe(def);
  });

  it('multi-space and tab runs collapse to single spaces (lossy by design)', () => {
    const multi = `${LONG.slice(0, 60)}  ${LONG.slice(60, 120)}   ${LONG.slice(120)}`;
    const { wrapped, def } = renderWithText(multi);
    expect(wrapped).toContain(' | ');
    const unwrapped = unwrapToonText(wrapped);
    expect(unwrapped).not.toBe(def);
    expect(def).toContain('  ');
    expect(unwrapped).not.toContain('  ');
  });

  it('tab separators rejoin as single spaces', () => {
    const tabbed = `alpha\tbravo charlie delta echo foxtrot golf hotel india juliet kilo lima mike november oscar papa quebec romeo sierra tango uniform victor whiskey extra words here`;
    const { wrapped, def } = renderWithText(tabbed);
    const unwrapped = unwrapToonText(wrapped);
    expect(unwrapped).not.toBe(def);
    expect(def).toContain('\t');
    expect(unwrapped).not.toContain('\t');
  });
});
