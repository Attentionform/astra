import { describe, expect, it, vi } from 'vitest';
import AdmZip from 'adm-zip';

let failAt: 'none' | 'mark' | 'fill' = 'none';

vi.mock('../src/compose/template-seed', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/compose/template-seed')>();
  return {
    ...actual,
    markTemplateComposed: (...args: Parameters<typeof actual.markTemplateComposed>) => {
      if (failAt === 'mark') throw new Error('injected marker failure');
      return actual.markTemplateComposed(...args);
    },
    fillSlot: (...args: Parameters<typeof actual.fillSlot>) => {
      if (failAt === 'fill') throw new Error('injected slot failure');
      return actual.fillSlot(...args);
    },
  };
});

import { DocxParser, composeDocument } from '../src/internal';

const entries = (buf: Buffer): Record<string, string> => {
  const zip = new AdmZip(buf);
  return Object.fromEntries(zip.getEntries().filter((e) => !e.isDirectory).map((e) => [e.entryName, e.getData().toString('utf8')]));
};

describe('a failure while the template is being replaced rewinds the whole compose', () => {
  for (const point of ['mark', 'fill'] as const) {
    it(`${point}: every part of the package is exactly as it was, and the next compose still works`, async () => {
      const parser = DocxParser.createNewDocument('report');
      const before = entries(parser.rezip());
      failAt = point;
      const res = await composeDocument(parser, '# Title\n\nBody text.\n');
      expect(res.success).toBe(false);
      expect(res.rollbackTriggered).toBe(true);
      expect(res.errors.some((e) => /Template cleanup failed/.test(e.error))).toBe(true);
      expect(entries(parser.rezip())).toEqual(before);

      failAt = 'none';
      const again = await composeDocument(parser, '# Title\n\nBody text.\n');
      expect(again.success, JSON.stringify(again.errors)).toBe(true);
      expect(again.warnings.filter((w) => w.startsWith('TEMPLATE_'))).toEqual([]);
      expect(entries(parser.rezip())['word/document.xml']).toContain('Body text.');
    });
  }
});
