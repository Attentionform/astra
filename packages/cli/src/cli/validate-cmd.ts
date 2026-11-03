import { DocxParser, chartXmlGetterForAst, lintDocumentStyle, resolveStyleGuide, summarizeStyleGuide, validateDocxBuffer, forecastNormalize, normalizeHintSentence, type StyleGuideConfig } from '@attentionform/astra-engine';
import { CLI_NAME, EXIT_FAIL, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman, failWithCode, inferBatchCode } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { macroFlagsUsageError, macroOptsFromFlags, parseMaxRatioFlag, parseMaxUncompressedBytesFlag, readDocx, withZipCeilingEnv, type ZipCeilingOverrides } from './io';

/** `--json` lists up to this many errors unless `--max-errors` says otherwise. */
const DEFAULT_MAX_ERRORS = 500;
const MAX_ERRORS_LIMIT = 5000;
/** The human output shows this many errors and counts the rest. */
const HUMAN_ERRORS_SHOWN = 10;

export async function cmdValidate(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'validate');
  let file: string | undefined;
  let maxRatioRaw: string | undefined;
  let maxBytesRaw: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  let styleGuideRaw: string | undefined;
  let maxErrorsRaw: string | undefined;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--max-ratio') maxRatioRaw = args[++i];
    else if (a.startsWith('--max-ratio=')) maxRatioRaw = a.slice('--max-ratio='.length);
    else if (a === '--max-uncompressed-bytes') maxBytesRaw = args[++i];
    else if (a.startsWith('--max-uncompressed-bytes=')) maxBytesRaw = a.slice('--max-uncompressed-bytes='.length);
    else if (a === '--max-errors') {
      const value = args[++i];
      if (value === undefined) return usage(out, `validate --max-errors requires a value (integer from 1 to ${MAX_ERRORS_LIMIT})`);
      maxErrorsRaw = value;
    } else if (a.startsWith('--max-errors=')) maxErrorsRaw = a.slice('--max-errors='.length);
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
  let maxErrors = DEFAULT_MAX_ERRORS;
  if (maxErrorsRaw !== undefined) {
    const n = Number(maxErrorsRaw);
    if (!/^\d+$/.test(maxErrorsRaw) || !Number.isSafeInteger(n) || n < 1 || n > MAX_ERRORS_LIMIT) {
      return usage(out, `validate --max-errors must be an integer from 1 to ${MAX_ERRORS_LIMIT} (got "${maxErrorsRaw}")`);
    }
    maxErrors = n;
  }
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
      let p: DocxParser;
      try {
        p = new DocxParser(buf, macroOpts);
      } catch (err: unknown) {
        // A package without [Content_Types].xml is refused at open, so it cannot be edited. A check
        // reports it as invalid with a clear message stating the root cause.
        if (err instanceof Error && /missing \[Content_Types\]\.xml/.test(err.message)) {
          return {
            file: file ?? '',
            ok: false,
            errors: [
              {
                description: '[Content_Types].xml is missing from the package root, so Word cannot find any part of this file.',
                path: '/[Content_Types].xml',
                errorType: 'Package',
                id: 'Pkg_MissingContentTypes',
                class: 'package',
              },
            ],
            truncated: false,
          };
        }
        throw err;
      }
      try {
        pendingNormalizations = p.getPendingNormalizations();
      } catch {
        pendingNormalizations = [];
      }
      return p.validateOOXML({ maxErrors });
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
    const errors = result.errors.slice(0, maxErrors).map((e) => {
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
    // The validator marks a list it cut at the limit; `errors` then holds the first ones, not all.
    const truncated = result.truncated === true;
    print(out, `INVALID ${file}: ${truncated ? `at least ${errors.length}` : errors.length} error(s)`);
    for (const e of errors.slice(0, HUMAN_ERRORS_SHOWN)) print(out, `  - ${e.slice(0, 500)}`);
    if (errors.length > HUMAN_ERRORS_SHOWN) {
      print(out, `  ... and ${errors.length - HUMAN_ERRORS_SHOWN} more${truncated ? ' (the list is cut at the limit)' : ''}; --json lists them (up to ${maxErrors}, --max-errors <n> changes the limit)`);
    }
    const forecast = truncated ? undefined : await forecastNormalize(buf);
    const hint = forecast ? normalizeHintSentence(forecast, errors.length, `Run \`${CLI_NAME} normalize ${file}\`: it`) : undefined;
    printJson(out, {
      ok: false,
      file,
      error: `[E_VALIDATION] "${file}" failed OOXML validation`,
      code: 'E_VALIDATION',
      errorCount: errors.length,
      ...(truncated ? { truncated: true, maxErrors } : {}),
      ...(hint ? { normalizeHint: hint } : {}),
      errors,
    });
    // Validation failure is operational (fail-closed), not usage.
    if (!out.json) {
      const hintSuffix = hint ? ` ${hint}` : '';
      printStderr(`${CLI_NAME}: error: [E_VALIDATION] "${file}" failed OOXML validation. Remedy: repair the document in Word and retry.${hintSuffix}`);
    }
    return EXIT_FAIL;
  } catch (err: unknown) {
    return failHuman(out, err, { file, fallbackCode: 'E_VALIDATION' });
  }
}
