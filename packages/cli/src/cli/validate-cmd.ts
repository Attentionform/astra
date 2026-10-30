import { DocxParser, chartXmlGetterForAst, lintDocumentStyle, resolveStyleGuide, summarizeStyleGuide, type StyleGuideConfig } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_FAIL, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman, failWithCode, inferBatchCode } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { macroFlagsUsageError, macroOptsFromFlags, parseMaxRatioFlag, parseMaxUncompressedBytesFlag, readDocx, withZipCeilingEnv, type ZipCeilingOverrides } from './io';

export async function cmdValidate(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'validate');
  let file: string | undefined;
  let maxRatioRaw: string | undefined;
  let maxBytesRaw: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  let styleGuideRaw: string | undefined;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--max-ratio') maxRatioRaw = args[++i];
    else if (a.startsWith('--max-ratio=')) maxRatioRaw = a.slice('--max-ratio='.length);
    else if (a === '--max-uncompressed-bytes') maxBytesRaw = args[++i];
    else if (a.startsWith('--max-uncompressed-bytes=')) maxBytesRaw = a.slice('--max-uncompressed-bytes='.length);
    else if (a === '--allow-macros') allowMacros = true;
    else if (a === '--strip-macros') stripMacros = true;
    else if (a === '--style-guide') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'validate --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
      }
      styleGuideRaw = args[++i];
    }
    else if (a.startsWith('--style-guide=')) {
      const value = a.slice('--style-guide='.length);
      if (value.length === 0) {
        return usage(out, 'validate --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
      }
      styleGuideRaw = value;
    }
    else if (a.startsWith('--')) return usage(out, `validate: unknown flag "${a}"`);
    else if (!file) file = a;
    else return usage(out, 'validate takes exactly one <docx> argument');
  }
  if (!file) return usage(out, 'validate requires exactly one <docx> argument');
  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  const ceilings: ZipCeilingOverrides = {};
  try {
    if (maxRatioRaw !== undefined) ceilings.maxRatio = parseMaxRatioFlag(maxRatioRaw);
    if (maxBytesRaw !== undefined) ceilings.maxUncompressedBytes = parseMaxUncompressedBytesFlag(maxBytesRaw);
  } catch (err: unknown) {
    return usage(out, err instanceof Error ? err.message : String(err));
  }
  let buf: Buffer;
  try {
    buf = await readDocx(file);
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_IO' });
  }
  try {
    // Guide attach (fail-closed before reading): unknown names, jail
    // escapes, and remote URLs refuse with E_VALIDATION/E_SANDBOX.
    // Validate carries no --allowed-dir, so file refs are cwd-anchored.
    let styleGuide: StyleGuideConfig | undefined;
    let guideParseNotes: string[] = [];
    if (styleGuideRaw !== undefined) {
      try {
        const resolved = resolveStyleGuide(styleGuideRaw, { allowedDir: process.cwd() });
        styleGuide = resolved.config;
        guideParseNotes = [...resolved.warnings];
        for (const w of resolved.warnings) print(out, `STYLE-NOTE ${w.slice(0, 500)}`);
      } catch (err: unknown) {
        const rawMsg = err instanceof Error ? err.message : String(err);
        return failWithCode(out, rawMsg, inferBatchCode(err, 'E_VALIDATION'), undefined, undefined, err);
      }
    }
    // R-8: flag ceilings scope the DocxParser zip-bomb gate for this load.
    const macroOpts = macroOptsFromFlags({ allowMacros, stripMacros });
    // The verdict is on the file as a save would write it (children put in
    // schema order first), so list the repairs that save would make: a file
    // Word may reject as it stands still reads VALID here.
    let pendingNormalizations: string[] = [];
    const result = await withZipCeilingEnv(ceilings, async () => {
      const p = new DocxParser(buf, macroOpts);
      try {
        pendingNormalizations = p.getPendingNormalizations();
      } catch {
        pendingNormalizations = [];
      }
      return p.validateOOXML();
    });
    const pending = pendingNormalizations.length > 0 ? { pendingNormalizations } : {};
    if (result.ok) {
      print(out, `VALID ${file}`);
      for (const n of pendingNormalizations.slice(0, 10)) print(out, `would-be normalize: ${n}`);
      if (pendingNormalizations.length > 10) {
        print(out, `would-be normalize: +${pendingNormalizations.length - 10} more (total ${pendingNormalizations.length}; save applies them)`);
      }
      if (styleGuide !== undefined) {
        // Advisory style pass (read-only, best-effort): findings ride the
        // envelope but never flip the verdict — style is advisory, schema
        // is law.
        let styleWarnings: Array<{ nodeId: string; rule: string; message: string }> = [];
        try {
          const guideParser = new DocxParser(buf, macroOpts);
          const ast = await guideParser.getAST();
          const zip = guideParser.getZipAdapter();
          styleWarnings = lintDocumentStyle(
            ast,
            styleGuide,
            chartXmlGetterForAst(ast, (part) => zip.readTextOrNull(part)),
          );
        } catch {
          styleWarnings = [];
        }
        for (const w of styleWarnings) {
          print(out, `STYLE ${w.rule} ${w.nodeId}: ${w.message.slice(0, 500)}`);
        }
        // Guide parse notes ride the same machine channel as lint findings
        // (human STYLE-NOTE lines stay too): one entry per note so --json
        // consumers never have to scrape stderr.
        for (const note of guideParseNotes) {
          styleWarnings.push({ nodeId: '(guide)', rule: 'guide-parse', message: note.slice(0, 500) });
        }
        printJson(out, {
          ok: true,
          file,
          errors: [],
          ...pending,
          styleGuide: summarizeStyleGuide(styleGuide),
          styleWarnings,
        });
        return EXIT_OK;
      }
      printJson(out, { ok: true, file, errors: [], ...pending });
      return EXIT_OK;
    }
    const errors = result.errors.slice(0, 10).map((e) => {
      const err = e as { description?: unknown; message?: unknown; path?: unknown; xpath?: unknown; xPath?: unknown };
      const desc =
        (typeof err.description === 'string' && err.description) ||
        (typeof err.message === 'string' && err.message) ||
        JSON.stringify(e);
      const where =
        (typeof err.xpath === 'string' && err.xpath) ||
        (typeof err.xPath === 'string' && err.xPath) ||
        (typeof err.path === 'string' && err.path);
      return where ? `${desc} (at ${where})` : desc;
    });
    print(out, `INVALID ${file}: ${result.errors.length} error(s)`);
    for (const e of errors) print(out, `  - ${e.slice(0, 500)}`);
    printJson(out, { ok: false, file, error: `[E_VALIDATION] "${file}" failed OOXML validation`, code: 'E_VALIDATION', errorCount: result.errors.length, errors });
    // Validation failure is operational (fail-closed), not usage.
    if (!out.json) printStderr(`${CLI_NAME}: error: [E_VALIDATION] "${file}" failed OOXML validation. Remedy: repair the document in Word and retry`);
    return EXIT_FAIL;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
