import fs from 'node:fs';
import { describe, expect, it } from 'vitest';
import { validatorBinary } from '../helpers.js';
import { buildCorpus, runCorpus, type GoldenEntry } from './corpus.js';

const golden: { entries: Record<string, GoldenEntry> } = JSON.parse(
  fs.readFileSync(new URL('./upstream-golden.json', import.meta.url), 'utf8')
);

describe('one-shot output matches the reference validator it was adapted from', () => {
  const docs = buildCorpus();

  it('covers valid and invalid documents, including unreadable ones', () => {
    const values = Object.values(golden.entries);
    expect(values.filter((e) => e.json && e.ok).length).toBeGreaterThan(10);
    expect(values.filter((e) => e.json && e.ok === false).length).toBeGreaterThan(30);
    expect(golden.entries['garbage.docx']?.ok).toBe(false);
    expect(golden.entries['empty.docx']?.ok).toBe(false);
  });

  it('has a recorded result for every corpus document and no stale ones', () => {
    expect(Object.keys(golden.entries).sort()).toEqual(docs.map((d) => d.name).sort());
  });

  it('gives the same exit code, ok flag and errors for every document', async () => {
    const actual = await runCorpus(validatorBinary(), docs);
    const differing = Object.keys(golden.entries).filter(
      (name) => JSON.stringify(actual[name]) !== JSON.stringify(golden.entries[name])
    );
    expect(differing).toEqual([]);
  }, 180_000);
});
