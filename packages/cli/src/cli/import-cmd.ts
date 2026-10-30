import * as fsp from 'fs/promises';
import * as path from 'path';
import { DocxParser, resolveStyleGuide, summarizeStyleGuide, type StyleGuideConfig } from '@attentionform/astra-engine';
import { composeDocument } from '@attentionform/astra-engine/compose/markdown';
import { CLI_NAME, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman, failWithCode, inferBatchCode, withGatedEngineOutput } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { writeDocx } from './io';
import { CREATE_TEMPLATES, type CreateTemplate } from './create-cmd';

export async function cmdImport(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'import');
  let src: string | undefined;
  let outRaw: string | undefined;
  let templateRaw: string | undefined;
  let allowedDirRaw: string | undefined;
  let styleGuideRaw: string | undefined;
  let dollarMath = false;
  let force = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--template') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'import --template requires a value (report|memo|blank)');
      }
      templateRaw = args[++i];
    }
    else if (a.startsWith('--template=')) templateRaw = a.slice('--template='.length);
    else if (a === '--allowed-dir') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'import --allowed-dir requires a value (directory holding image/chart files)');
      }
      allowedDirRaw = args[++i];
    }
    else if (a.startsWith('--allowed-dir=')) allowedDirRaw = a.slice('--allowed-dir='.length);
    else if (a === '--style-guide') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'import --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
      }
      styleGuideRaw = args[++i];
    }
    else if (a.startsWith('--style-guide=')) styleGuideRaw = a.slice('--style-guide='.length);
    else if (a === '-o' || a === '--out' || a === '--output') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'import -o requires a value (output .docx path)');
      }
      outRaw = args[++i];
    }
    else if (a.startsWith('-o=')) outRaw = a.slice('-o='.length);
    else if (a.startsWith('--out=')) outRaw = a.slice('--out='.length);
    else if (a.startsWith('--output=')) outRaw = a.slice('--output='.length);
    else if (a === '--dollar-math') dollarMath = true;
    else if (a === '--force') force = true;
    else if (a.startsWith('--')) return usage(out, `import: unknown flag "${a}"`);
    else if (!src) src = a;
    else return usage(out, 'import takes exactly one <file.md> argument');
  }
  if (!src) return usage(out, 'import requires <file.md>');
  if (outRaw !== undefined && outRaw.length === 0) {
    return usage(out, 'import -o requires a value (output .docx path)');
  }
  if (allowedDirRaw !== undefined && allowedDirRaw.length === 0) {
    return usage(out, 'import --allowed-dir requires a value (directory holding image/chart files)');
  }
  if (styleGuideRaw !== undefined && styleGuideRaw.length === 0) {
    return usage(out, 'import --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
  }
  const template: string = templateRaw ?? 'blank';
  if (!(CREATE_TEMPLATES as readonly string[]).includes(template)) {
    return usage(out, `import --template must be one of ${CREATE_TEMPLATES.join('|')} (got "${template}")`);
  }
  const srcExt = path.extname(src).toLowerCase();
  if (srcExt === '.html' || srcExt === '.htm') {
    return failWithCode(
      out,
      `HTML import is not supported directly ("${path.basename(src)}")`,
      'E_VALIDATION',
      undefined,
      'ASTra imports Markdown (.md). Convert the HTML file to Markdown or supply a .md document'
    );
  }
  const dest =
    outRaw ??
    path.join(path.dirname(path.resolve(src)), `${path.basename(src, path.extname(src))}.docx`);
  if (!force) {
    try {
      await fsp.access(dest);
      return usage(out, `import: destination "${path.basename(dest)}" already exists (refusing to overwrite without --force)`);
    } catch {
      /* destination is free — proceed */
    }
  }
  let mdSource: string;
  try {
    mdSource = await fsp.readFile(src, 'utf8');
  } catch (err: unknown) {
    return failHuman(out, err, {
      file: src,
      fallbackCode: 'E_IO',
      fallbackRemedy: 'check the Markdown path exists and is readable, then retry',
    });
  }
  try {
    // Same compose path as create --from-md: fresh template seed plus the
    // strict compile gate, quotas, structural checks, per-chunk atomic
    // rollback, and inline-span run splitting. Input errors throw before any
    // write; the line number in the message names the failing block.
    // Guide attach (fail-closed before any write, same resolve/jail
    // semantics as create --from-md): unknown names, jail escapes, and
    // remote URLs refuse with E_VALIDATION/E_SANDBOX. File refs jail to
    // --allowed-dir when one is in play, else cwd-anchored.
    let styleGuide: StyleGuideConfig | 'legacy' | undefined;
    let styleWarnings: string[] = [];
    if (styleGuideRaw === 'legacy') {
      // The pre-default look: grey table header, no header text colour.
      styleGuide = 'legacy';
    } else if (styleGuideRaw !== undefined) {
      try {
        const resolved = resolveStyleGuide(styleGuideRaw, {
          allowedDir: allowedDirRaw ?? process.cwd(),
        });
        styleGuide = resolved.config;
        styleWarnings = [...resolved.warnings];
        for (const w of resolved.warnings) printStderr(`${CLI_NAME}: WARNING: ${w}`);
      } catch (err: unknown) {
        const rawMsg = err instanceof Error ? err.message : String(err);
        return failWithCode(out, rawMsg, inferBatchCode(err, 'E_VALIDATION'), undefined, undefined, err);
      }
    }
    const parser = DocxParser.createNewDocument(template as CreateTemplate);
    const composed = await withGatedEngineOutput(out, () =>
      composeDocument(parser, mdSource, {
        ...(allowedDirRaw !== undefined ? { allowedDir: allowedDirRaw } : {}),
        template: template as CreateTemplate,
        ...(styleGuide !== undefined ? { styleGuide } : {}),
        ...(dollarMath ? { dollarMath: true } : {}),
      })
    );
    if (!composed.success) {
      const first = composed.errors[0];
      const rawMessage = first ? String(first.error) : 'Markdown compose failed.';
      const hint = first && typeof first.hint === 'string' && first.hint ? first.hint : undefined;
      return failWithCode(
        out,
        rawMessage,
        inferBatchCode(rawMessage, 'E_VALIDATION'),
        undefined,
        hint ?? 'fix the flagged Markdown block and retry import',
        rawMessage
      );
    }
    const composeWarnings = [...composed.warnings];
    for (const w of composeWarnings) printStderr(`${CLI_NAME}: WARNING: ${w}`);
    const validation = await parser.validateOOXML();
    if (!validation.ok) {
      return failHuman(out, `composed "${template}" document failed validation`, {
        fallbackCode: 'E_VALIDATION',
        fallbackRemedy: 'retry with --template blank and report the template name',
      });
    }
    await writeDocx(dest, parser.rezip());
    print(out, `IMPORTED ${src} -> ${dest} (template=${template}, applied=${composed.appliedCount})`);
    printJson(out, {
      ok: true,
      command: 'import',
      src,
      file: dest,
      template,
      appliedCount: composed.appliedCount,
      valid: true,
      ...(allowedDirRaw !== undefined ? { allowedDir: allowedDirRaw } : {}),
      ...(styleGuide === 'legacy' ? { styleGuide: 'legacy' } : styleGuide !== undefined ? { styleGuide: summarizeStyleGuide(styleGuide) } : {}),
      ...(styleWarnings.length > 0 ? { styleWarnings } : {}),
      ...(composeWarnings.length > 0 ? { composeWarnings } : {}),
    });
    return EXIT_OK;
  } catch (err: unknown) {
    return failHuman(out, err, { file: dest, fallbackCode: 'E_IO' });
  }
}
