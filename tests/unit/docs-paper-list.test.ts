import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { PAPER_NAMES } from '../../packages/engine/src/page/papers';

const root = path.resolve(__dirname, '..', '..');
const read = (p: string): string => fs.readFileSync(path.join(root, p), 'utf8');

describe('paper names in docs and CLI usage', () => {
  // Each place that lists the papers names every one of them, and no name that does not exist.
  for (const file of ['docs/page-layout.md', 'docs/cli-guide.md', 'docs/mcp-specification.md', 'packages/cli/src/cli/usage.ts']) {
    it(`${file} lists every named paper`, () => {
      const text = read(file);
      const missing = PAPER_NAMES.filter((n) => !new RegExp(`\\b${n}\\b`).test(text));
      expect(missing).toEqual([]);
    });
  }
});
