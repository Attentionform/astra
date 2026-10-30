import { DocxParser } from '@attentionform/astra-engine';
import { EXIT_OK, type GlobalOpts, print, printJson } from './output';
import { failHuman } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { computeDocStats, readDocx } from './io';

export async function cmdStats(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'stats');
  if (args.length !== 1 || args[0].startsWith('--')) {
    return usage(out, 'stats requires exactly one <docx> argument');
  }
  const file = args[0];
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  try {
    const parser = new DocxParser(buf);
    const ast = await parser.getAST();
    const stats = computeDocStats(ast);
    print(out, `STATS ${file}: words=${stats.words} paragraphs=${stats.paragraphs} pages~${stats.pages} (est, 500 wpp)`);
    print(out, `recipe: body paragraphs + table cells; excludes headers/footers/footnotes/comments; pages~P rough estimate = max(1, ceil(words/500))`);
    printJson(out, { ok: true, file, words: stats.words, paragraphs: stats.paragraphs, pages: stats.pages });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
