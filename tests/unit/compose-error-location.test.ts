import { describe, it, expect } from 'vitest';
import { DocxParser } from '@attentionform/astra-engine';
import { compileMarkdown, composeDocument } from '../../packages/engine/src/compose/markdown';

/**
 * Every E_VALIDATION from compose says where the problem is ("line N") and how to fix it
 * ("Remedy:"). A reader of a long Markdown file must never get a bare "validation failed".
 */
const REJECTED: Array<{ name: string; markdown: string; line: number }> = [
  { name: 'remote image', markdown: '# Title\n\nText.\n\n![x](https://example.com/a.png)\n', line: 5 },
  { name: 'raw HTML block', markdown: '# Title\n\n<div>hello</div>\n', line: 3 },
  { name: 'script link', markdown: 'Intro.\n\nSee [x](javascript:alert(1)) here.\n', line: 3 },
  { name: 'missing footnote definition', markdown: '# A\n\nClaim[^missing] here.\n', line: 3 },
];

describe('compose errors name the line', () => {
  for (const c of REJECTED) {
    it(`${c.name}: message holds "line ${c.line}" and a remedy`, async () => {
      const parser = DocxParser.createNewDocument('blank');
      let message = '';
      try {
        compileMarkdown(c.markdown);
        await composeDocument(parser, c.markdown);
      } catch (err: unknown) {
        message = err instanceof Error ? err.message : String(err);
      }
      expect(message, 'the input must be rejected').not.toBe('');
      expect(message).toContain('E_VALIDATION');
      expect(message).toMatch(/line \d+/);
      expect(message).toContain(`line ${c.line}`);
      expect(message).toMatch(/Remedy:/);
    });
  }
});
