import { describe, it, expect } from 'vitest';
import { ToonGenerator, unwrapToonText } from '@attentionform/astra-engine';

const PROSE =
  'The quick brown fox jumps over the lazy dog near the quiet river bank at dawn ' +
  'with soft light on the still water while birds sing over the green hill side far area';

const LONG_PATH = 'LongSectionHeadingWithoutSpacesExceedingFortyColumnsWide';
const LONG_STYLE = 'CustomStyleNameThatIsLongAndHasNoSpacesAtAllLong';

function buildFixture(): any {
  return {
    sections: [
      {
        id: 'sA',
        version: 1,
        checksum: 'cA',
        heading: { id: 'hA', text: 'G', level: 1 },
        nodes: [
          { id: 'n1', version: 1, checksum: 'c1', type: 'paragraph', metadata: { text: PROSE } },
          { id: 'n2', version: 1, checksum: 'c2', type: 'paragraph', metadata: { text: 'Hi there', hasBookmark: true } },
        ],
      },
      {
        id: 'sB',
        version: 1,
        checksum: 'cB',
        heading: { id: 'hB', text: LONG_PATH, level: 1 },
        nodes: [
          {
            id: 'n3',
            version: 2,
            checksum: 'c3',
            type: 'paragraph',
            metadata: { text: PROSE, namedStyle: LONG_STYLE },
          },
          {
            id: 'n4',
            version: 1,
            checksum: 'c4',
            type: 'paragraph',
            metadata: {
              text: PROSE,
              hasImage: true,
              hasEquation: true,
              hasComment: true,
              hasTrackedChanges: true,
              hasHyperlink: true,
              hasBookmark: true,
            },
          },
        ],
      },
    ],
  };
}

function renderDefault(): { gen: ToonGenerator; table: any; text: string } {
  const gen = new ToonGenerator();
  const table = gen.generateToon(buildFixture());
  return { gen, table, text: gen.renderToonText(table) };
}

/** Column-count-aware split of one standard/compact logical row (Text may hold the separator). */
function parseRow(line: string, sep: string): {
  id: string;
  version: string;
  path: string;
  header: string;
  style: string;
  text: string;
  flags: string;
} {
  const parts = line.split(sep);
  expect(parts.length).toBeGreaterThanOrEqual(7);
  return {
    id: parts[0],
    version: parts[1],
    path: parts[2],
    header: parts[3],
    style: parts[4],
    text: parts.slice(5, parts.length - 1).join(sep),
    flags: parts[parts.length - 1],
  };
}

describe('text-column wrap: machine default untouched', () => {
  it('default render stays strict 1-line-per-node', () => {
    const { table, text } = renderDefault();
    expect(table.rows.length).toBe(4);
    const lines = text.split('\n').filter((l) => l.length > 0);
    expect(lines[0]).toBe('> fidelity: lossy tier: standard');
    expect(lines.length).toBe(table.rows.length + 3);
    for (const l of lines) expect(l.startsWith(' ')).toBe(false);
  });

  it('default output is repeat-stable, huge widths are identical, bad widths are ignored', () => {
    const { gen, table, text } = renderDefault();
    expect(gen.renderToonText(table)).toBe(text);
    expect(gen.renderToonText(table, { width: 100000 })).toBe(text);
    for (const bad of [0, -5, Number.NaN]) {
      expect(gen.renderToonText(table, { width: bad })).toBe(text);
    }
  });
});

describe('text-column wrap: only Text wraps', () => {
  it('width 80 wraps long Text with hanging indent; only the last line carries flags', () => {
    const { gen, table } = renderDefault();
    const wrapped = gen.renderToonText(table, { width: 80 });
    expect(wrapped.split('\n')[0]).toBe('> fidelity: lossy tier: standard');
    const lines = wrapped.split('\n').filter((l) => l.length > 0);
    const group: string[] = [];
    let open = false;
    for (const l of lines) {
      if (l.startsWith('n1 | ')) {
        open = true;
        group.push(l);
      } else if (open && l.startsWith(' ')) {
        group.push(l);
      } else {
        open = false;
      }
    }
    expect(group.length).toBeGreaterThan(1);
    const prefix = 'n1 | 1 | G | G | body:L-1 | ';
    expect(group[0].startsWith(prefix)).toBe(true);
    const firstSeg = group[0].slice(prefix.length);
    expect(firstSeg.length).toBeGreaterThan(0);
    expect(PROSE.startsWith(firstSeg)).toBe(true);
    expect(group[0].includes(' | ', prefix.length)).toBe(false);
    for (const c of group.slice(1, -1)) expect(c).not.toContain(' | ');
    expect(group[group.length - 1].endsWith(' | ')).toBe(true);
  });

  it('wrapped output unwraps to byte-identical default (80 and 40)', () => {
    const { gen, table, text } = renderDefault();
    expect(unwrapToonText(gen.renderToonText(table, { width: 80 }))).toBe(text);
    expect(unwrapToonText(gen.renderToonText(table, { width: 40 }))).toBe(text);
  });

  it('40-col smoke: short-prefix row fits 40, indent aligns under Text, cells re-parse identically', () => {
    const { gen, table, text } = renderDefault();
    const wrapped = gen.renderToonText(table, { width: 40 });
    const lines = wrapped.split('\n').filter((l) => l.length > 0);
    const group: string[] = [];
    let open = false;
    for (const l of lines) {
      if (l.startsWith('n1 | ')) {
        open = true;
        group.push(l);
      } else if (open && l.startsWith(' ')) {
        group.push(l);
      } else {
        open = false;
      }
    }
    expect(group.length).toBeGreaterThan(1);
    for (const l of group) expect(l.length).toBeLessThanOrEqual(40);
    const prefix = 'n1 | 1 | G | G | body:L-1 | ';
    const indent = ' '.repeat(prefix.length);
    expect(group[0].startsWith(prefix)).toBe(true);
    for (const c of group.slice(1, -1)) {
      expect(c.startsWith(indent)).toBe(true);
      expect(c).not.toContain(' | ');
    }
    const last = group[group.length - 1];
    expect(last.startsWith(indent)).toBe(true);
    expect(last.endsWith(' | ')).toBe(true);
    const unwrapped = unwrapToonText(wrapped);
    expect(unwrapped).toBe(text);
    const row = table.rows[0];
    const cells = parseRow(
      unwrapped.split('\n').find((l) => l.startsWith('n1 | ')) as string,
      ' | '
    );
    expect(cells).toEqual({
      id: row.id,
      version: String(row.version),
      path: row.path,
      header: row.header,
      style: row.style,
      text: row.text,
      flags: row.flags,
    });
  });

  it('other columns never wrap: overlong path/style/flags stay whole on the first line', () => {
    const { gen, table } = renderDefault();
    const wrapped = gen.renderToonText(table, { width: 40 });
    const lines = wrapped.split('\n').filter((l) => l.length > 0);
    const n3 = lines.find((l) => l.startsWith('n3 | ')) as string;
    expect(n3).toContain(LONG_PATH);
    expect(n3).toContain(LONG_STYLE);
    const afterN3 = lines[lines.indexOf(n3) + 1];
    expect(afterN3.startsWith(' ')).toBe(false);
    const n4 = lines.find((l) => l.startsWith('n4 | ')) as string;
    expect(n4).toContain('[IMG] [EQ] [CMT] [TRK] [LNK] [BKMK]');
  });

  it('an overlong single word stays whole and still round-trips', () => {
    const gen = new ToonGenerator();
    const text = `short ${'x'.repeat(50)} tail words here`;
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
    const wrapped = gen.renderToonText(table, { width: 40 });
    expect(wrapped).toContain('x'.repeat(50));
    expect(unwrapToonText(wrapped)).toBe(def);
  });

  it('compact separator wraps the same way and round-trips', () => {
    const gen = new ToonGenerator();
    const table = gen.toon(buildFixture(), { fidelity: 'compact' });
    const def = gen.renderToonText(table, { compact: true });
    const wrapped = gen.renderToonText(table, { compact: true, width: 80 });
    const lines = wrapped.split('\n').filter((l) => l.length > 0);
    const groups: string[][] = [];
    for (const l of lines) {
      if (l.startsWith(' ') && groups.length > 0) groups[groups.length - 1].push(l);
      else groups.push([l]);
    }
    expect(groups.some((g) => g.length > 1)).toBe(true);
    for (const g of groups) {
      for (const c of g.slice(1, -1)) expect(c).not.toContain('|');
    }
    expect(unwrapToonText(wrapped)).toBe(def);
  });
});
