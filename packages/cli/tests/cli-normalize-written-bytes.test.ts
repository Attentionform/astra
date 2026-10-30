import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { DocxParser, validateDocxBuffer } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

const GALLERY = path.resolve(__dirname, '../../../test-gallery');
const FIXTURES = fs.readdirSync(GALLERY).filter((f) => f.endsWith('.docx'));

let dir: string;

beforeEach(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-norm-bytes-'));
  vi.spyOn(console, 'log').mockImplementation(() => undefined);
  vi.spyOn(console, 'error').mockImplementation(() => undefined);
});

afterEach(() => {
  vi.restoreAllMocks();
  fs.rmSync(dir, { recursive: true, force: true });
});

describe('normalize writes the bytes its validation gate checked', () => {
  it('has fixtures to run against', () => {
    expect(FIXTURES.length).toBeGreaterThan(0);
  });

  for (const name of FIXTURES) {
    it(`${name}: the written file is what the gate rebuilds from it, and it validates`, async () => {
      const out = path.join(dir, `out-${name}`);
      expect(await run(['normalize', path.join(GALLERY, name), '--out', out])).toBe(0);
      const written = fs.readFileSync(out);
      // The gate parses the normalized buffer and validates the parser's rebuild
      // of it; the file on disk is the buffer itself. They must be the same bytes.
      expect(Buffer.compare(new DocxParser(written).rezip(), written)).toBe(0);
      expect((await validateDocxBuffer(written)).ok).toBe(true);
    }, 30000);
  }
});
