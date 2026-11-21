/**
 * Random page requests: each one is either refused with E_VALIDATION or produces a file that the OOXML validator accepts and whose text area
 * is at least the minimum. Applying the same request twice gives the same file. The generator is seeded, so a failure repeats.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, validateDocxBuffer, PAPER_NAMES } from '../src/internal';

let seed = 20261121;
const rnd = (): number => {
  seed = (seed * 1664525 + 1013904223) >>> 0;
  return seed / 2 ** 32;
};
const pick = <T,>(items: readonly T[]): T => items[Math.floor(rnd() * items.length)];
const length = (): number | string => {
  const unit = pick(['mm', 'cm', 'in', 'pt', 'tw', '']);
  const value = pick([0, 0.25, 0.5, 1, 2.5, 7.5, 12, 20, 25.4, 72, 150, 720, 720.5, 1440, 3000, 31680]) * (rnd() < 0.5 ? 1 : 1 + rnd());
  return unit === '' ? Math.round(value) : `${Math.round(value * 100) / 100}${unit}`;
};
const request = (): Record<string, unknown> => {
  const op: Record<string, unknown> = { action: 'pageSetup' };
  if (rnd() < 0.7) op.paperSize = rnd() < 0.5 ? pick(PAPER_NAMES) : { width: length(), height: length() };
  if (rnd() < 0.4) op.orientation = pick(['portrait', 'landscape']);
  if (rnd() < 0.7) {
    op.margins = rnd() < 0.3 ? length() : Object.fromEntries((['top', 'bottom', 'left', 'right', 'header', 'footer', 'gutter'] as const).filter(() => rnd() < 0.5).map((k) => [k, length()]));
  }
  if (rnd() < 0.2) op.columns = 1 + Math.floor(rnd() * 3);
  return op;
};

const files = (parser: DocxParser): string => {
  const zip = new AdmZip(parser.rezip());
  return `${zip.readAsText('word/document.xml')}\n${zip.readAsText('word/settings.xml')}`;
};

describe('random page requests', () => {
  for (let i = 0; i < 60; i++) {
    const op = request();
    it(`#${i} ${JSON.stringify(op)}`, async () => {
      const parser = DocxParser.createNewDocument('memo');
      const engine = new OperationEngine(await parser.getAST(), parser);
      const first = await engine.executeBatch({ operations: [op] as never });
      if (!first.success) {
        expect(JSON.stringify(first.errors)).toMatch(/E_VALIDATION|outside 1 in to 22 in|text (width|height)|must be|zero or more|not a known paper|is not a/);
        return;
      }
      const valid = await validateDocxBuffer(parser.rezip());
      expect(valid.ok, JSON.stringify(valid.errors)).toBe(true);
      const sect = [...parser.readDocumentXml().matchAll(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g)].map((m) => m[0]).pop() as string;
      const n = (re: RegExp): number => Number(re.exec(sect)?.[1] ?? 0);
      expect(n(/<w:pgSz[^>]*w:w="(\d+)"/) - n(/w:left="(-?\d+)"/) - n(/w:right="(-?\d+)"/) - n(/w:gutter="(\d+)"/)).toBeGreaterThanOrEqual(1080);
      expect(n(/<w:pgSz[^>]*w:h="(\d+)"/) - n(/w:top="(-?\d+)"/) - n(/w:bottom="(-?\d+)"/)).toBeGreaterThanOrEqual(720);
      expect(sect).not.toMatch(/="\d+\.\d+"/); // whole twips only
      // Applying the same request again changes nothing.
      const once = files(parser);
      const again = await new OperationEngine(await parser.getAST(), parser).executeBatch({ operations: [op] as never });
      expect(again.success, JSON.stringify(again.errors)).toBe(true);
      expect(files(parser)).toBe(once);
    });
  }
});
