import * as fsp from 'fs/promises';
import * as path from 'path';
import { DocxParser, resolveStyleGuide, summarizeStyleGuide, type StyleGuideConfig } from '@attentionform/astra-engine';
import { composeDocument } from '@attentionform/astra-engine/compose/markdown';
import { CLI_NAME, EXIT_OK, type GlobalOpts, print, printJson, printStderr } from './output';
import { failHuman, failWithCode, inferBatchCode, withGatedEngineOutput } from './errors';
import { hasHelpFlag, printCommandHelp, usage } from './usage';
import { writeDocx } from './io';

export const CREATE_TEMPLATES = ['report', 'memo', 'blank'] as const;
export type CreateTemplate = (typeof CREATE_TEMPLATES)[number];

export async function cmdCreate(out: GlobalOpts, args: string[]): Promise<number> {
  if (hasHelpFlag(args)) return printCommandHelp(out, 'create');
  let dest: string | undefined;
  let templateRaw: string | undefined;
  let fromMdRaw: string | undefined;
  let allowedDirRaw: string | undefined;
  let styleGuideRaw: string | undefined;
  let dollarMath = false;
  let force = false;
  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    if (a === '--template') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'create --template requires a value (report|memo|blank)');
      }
      templateRaw = args[++i];
    }
    else if (a.startsWith('--template=')) templateRaw = a.slice('--template='.length);
    else if (a === '--from-md') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'create --from-md requires a value (path to a Markdown file)');
      }
      fromMdRaw = args[++i];
    }
    else if (a.startsWith('--from-md=')) fromMdRaw = a.slice('--from-md='.length);
    else if (a === '--allowed-dir') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'create --allowed-dir requires a value (directory holding image/chart files)');
      }
      allowedDirRaw = args[++i];
    }
    else if (a.startsWith('--allowed-dir=')) allowedDirRaw = a.slice('--allowed-dir='.length);
    else if (a === '--style-guide') {
      const next = args[i + 1];
      if (next === undefined || next.length === 0 || next.startsWith('-')) {
        return usage(out, 'create --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
      }
      styleGuideRaw = args[++i];
    }
    else if (a.startsWith('--style-guide=')) styleGuideRaw = a.slice('--style-guide='.length);
    else if (a === '--dollar-math') dollarMath = true;
    else if (a === '--force') force = true;
    else if (a.startsWith('--')) return usage(out, `create: unknown flag "${a}"`);
    else if (!dest) dest = a;
    else return usage(out, 'create takes exactly one <out.docx> argument');
  }
  if (!dest) return usage(out, 'create requires <out.docx>');
  if (fromMdRaw !== undefined && fromMdRaw.length === 0) {
    return usage(out, 'create --from-md requires a value (path to a Markdown file)');
  }
  if (allowedDirRaw !== undefined && allowedDirRaw.length === 0) {
    return usage(out, 'create --allowed-dir requires a value (directory holding image/chart files)');
  }
  if (styleGuideRaw !== undefined && styleGuideRaw.length === 0) {
    return usage(out, 'create --style-guide requires a value (bundled guide name or path to a style-guide.md file)');
  }
  const template: string = templateRaw ?? 'blank';
  if (!(CREATE_TEMPLATES as readonly string[]).includes(template)) {
    return usage(out, `create --template must be one of ${CREATE_TEMPLATES.join('|')} (got "${template}")`);
  }
  if (!force) {
    try {
      await fsp.access(dest);
      return usage(out, `create: destination "${path.basename(dest)}" already exists (refusing to overwrite without --force)`);
    } catch {
      /* destination is free — proceed */
    }
  }
  let mdSource: string | undefined;
  if (fromMdRaw !== undefined) {
    try {
      mdSource = await fsp.readFile(fromMdRaw, 'utf8');
    } catch (err: unknown) {
      return failHuman(out, err, {
        file: fromMdRaw,
        fallbackCode: 'E_IO',
        fallbackRemedy: 'check the Markdown path exists and is readable, then retry',
      });
    }
  }
  try {
    // Guide attach (fail-closed before any write): unknown names, jail
    // escapes, and remote URLs refuse with E_VALIDATION/E_SANDBOX. File
    // refs jail to --allowed-dir when one is in play, else cwd-anchored.
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
    // Reuse the existing blank-doc builder (no new package edges): templates
    // are engine-owned, the CLI only gates + writes.
    const parser = DocxParser.createNewDocument(template as CreateTemplate);
    let appliedCount = 0;
    const composeWarnings: string[] = [];
    if (mdSource !== undefined) {
      // Same compose path as compose_document (strict schema gate, quotas,
      // structural checks, per-chunk atomic rollback, inline-span run
      // splitting). The fresh document must still hold the template seed.
      // Input errors throw before any write; runtime failures roll back per
      // chunk. Block image/chart refs resolve against --allowed-dir.
      const composed = await withGatedEngineOutput(out, () =>
        composeDocument(parser, mdSource as string, {
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
          hint ?? 'fix the flagged Markdown block and retry create --from-md',
          rawMessage
        );
      }
      appliedCount = composed.appliedCount;
      composeWarnings.push(...composed.warnings);
      for (const w of composeWarnings) printStderr(`${CLI_NAME}: WARNING: ${w}`);
    }
    const validation = await parser.validateOOXML();
    if (!validation.ok) {
      return failHuman(out, `fresh "${template}" template failed validation`, {
        fallbackCode: 'E_VALIDATION',
        fallbackRemedy: 'retry with --template blank and report the template name',
      });
    }
    await writeDocx(dest, parser.rezip());
    const fromMdSuffix = mdSource !== undefined ? `, fromMd=${path.basename(fromMdRaw as string)}, applied=${appliedCount}` : '';
    print(out, `CREATED ${dest} (template=${template}${fromMdSuffix})`);
    printJson(out, {
      ok: true,
      file: dest,
      template,
      ...(mdSource !== undefined ? { fromMd: fromMdRaw, appliedCount, valid: true } : {}),
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
