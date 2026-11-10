import { DocxParser } from '@attentionform/astra-engine';
import { EXIT_OK, type GlobalOpts, print, printJson } from './output';
import { failHuman } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { computeDocStats, macroFlagsUsageError, macroOptsFromFlags, readDocx } from './io';

export async function cmdStats(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'stats');
  let allowMacros = false;
  let stripMacros = false;
  const positional: string[] = [];
  for (const a of args) {
    if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a.startsWith('--')) return usage(out, `stats: unknown flag "${a}"`);
    else positional.push(a);
  }
  if (positional.length !== 1) {
    return usage(out, 'stats requires exactly one <docx> argument');
  }
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  const file = positional[0];
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  try {
    const parser = new DocxParser(buf, macroOptsFromFlags({ allowMacros, stripMacros }));
    const ast = await parser.getAST();
    const stats = computeDocStats(ast);
    print(out, `STATS ${file}: words=${stats.words} paragraphs=${stats.paragraphs} pages~${stats.pages} (est, 500 wpp)`);
    print(out, `tables=${stats.tables} callouts=${stats.callouts} images=${stats.images} (a callout is a one-cell filled table and is not counted in tables)`);
    print(out, `recipe: body paragraphs + table cells; excludes headers/footers/footnotes/comments; pages~P rough estimate = max(1, ceil(words/500))`);
    printJson(out, { ok: true, file, words: stats.words, paragraphs: stats.paragraphs, pages: stats.pages, tables: stats.tables, callouts: stats.callouts, images: stats.images });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
