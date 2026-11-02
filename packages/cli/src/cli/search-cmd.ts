import { DocxParser, cellAddressesMatching } from '@attentionform/astra-engine';
import { EXIT_OK, type GlobalOpts, print, printJson } from './output';
import { failHuman } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { buildSnippet, macroFlagsUsageError, macroOptsFromFlags, readDocx } from './io';

export async function cmdSearch(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'search');
  let file: string | undefined;
  let query: string | undefined;
  let limitRaw: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--limit') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'search --limit requires a value (integer 1..100)');
      }
      limitRaw = args[++i];
    }
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a.startsWith('--')) return usage(out, `search: unknown flag "${a}"`);
    else if (!file) file = a;
    else if (!query) query = a;
    else return usage(out, 'search takes exactly <docx> <query>');
  }
  if (!file || !query) return usage(out, 'search requires <docx> <query>');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  let limit = 20;
  if (limitRaw !== undefined) {
    const n = Number(limitRaw);
    if (!Number.isInteger(n) || n < 1 || n > 100) {
      return usage(out, 'search --limit must be an integer 1..100');
    }
    limit = n;
  }
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  try {
    // Engine-side search is the single source of truth for node identity:
    // DocxParser.searchNodesInAST matches the same AST the batch engine acts
    // on, so hit ids are directly usable as batch targets. The CLI only adds
    // snippet windowing + output shaping (no new cross-package edges).
    const parser = new DocxParser(buf, macroOptsFromFlags({ allowMacros, stripMacros }));
    const ast = await parser.getAST();
    const allHits = parser.searchNodesInAST(ast, query);
    const total = allHits.length;
    const hits = allHits.slice(0, limit);
    const needle = query.toLowerCase();
    const shaped = hits.map((h) => {
      const cells = h.type === 'table' ? cellAddressesMatching(h, (cell) => cell.toLowerCase().includes(needle)) : [];
      return {
        id: h.id,
        type: h.type,
        snippet: buildSnippet(h.metadata?.text ?? '', query),
        ...(cells.length > 0 ? { cells } : {}),
      };
    });
    for (const h of shaped) print(out, `${h.id} ${h.snippet}${h.cells ? ` [cells: ${h.cells.join(', ')}]` : ''}`);
    // `--limit` cuts the list, not the count: say how many there are, so a cut list is not read as the whole.
    if (total > shaped.length) {
      const advice = limit < 100 ? 'raise --limit (1..100) to see more' : 'the limit is 100';
      print(out, `${shaped.length} of ${total} matches; ${advice}`);
    } else {
      print(out, shaped.length === 1 ? '1 match' : `${shaped.length} matches`);
    }
    printJson(out, { ok: true, file, query, count: shaped.length, total, truncated: total > shaped.length, hits: shaped });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
