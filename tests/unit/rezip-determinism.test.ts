/**
 * The CLI and MCP save paths validate the bytes of one `rezip()` call and
 * write the bytes of another. That is only safe while `rezip()` is
 * deterministic and idempotent; this test is the reason it is safe.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser } from '../../packages/engine/src';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

const GALLERY = path.resolve(__dirname, '../../test-gallery');
const files = fs.readdirSync(GALLERY).filter((f) => f.endsWith('.docx'));

describe('rezip is deterministic and idempotent', () => {
  it('finds the gallery fixtures', () => {
    expect(files.length).toBeGreaterThan(0);
  });

  for (const name of files) {
    it(`${name}: same bytes from a fresh parser, a second call, and a rezip of the rezip`, () => {
      const input = fs.readFileSync(path.join(GALLERY, name));
      const a = new DocxParser(input).rezip();
      expect(Buffer.compare(a, new DocxParser(input).rezip())).toBe(0);
      const p = new DocxParser(input);
      const first = p.rezip();
      expect(Buffer.compare(first, p.rezip())).toBe(0);
      expect(Buffer.compare(a, new DocxParser(a).rezip())).toBe(0);
    });
  }
});
