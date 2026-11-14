import * as fsp from 'fs/promises';
import * as path from 'path';
import { executeExport } from '../export';
import { CLI_NAME, EXIT_OK, type GlobalOpts, formatCliOutput, print, printJson, stringifyJsonPayload } from './output';
import { failHuman } from './errors';
import { usage } from './usage';
import { macroFlagsUsageError, macroOptsFromFlags } from './io';
import { printRenderNotes, readRenderFlag, renderOptionsFromFlags, type RenderFlags } from './render-flags';

export async function cmdExport(out: GlobalOpts, args: string[]): Promise<number> {
  if (args.length === 0) return usage(out, 'export requires a <docx> argument');
  if (args.includes('--help') || args.includes('-h')) {
    const text = [
      `Usage: ${CLI_NAME} export <docx> [--format|-f <html|markdown|md|pdf>] [--output|--out|-o <path>] [--force] [--image-mode <embed|file>] [--asset-prefix <dir>] [--allow-macros] [--strip-macros] [--renderer <auto|word|libreoffice>] [--allow-network-references] [--show-markup]`,
      '',
      'Exports a .docx document to HTML, Markdown, or PDF.',
      '  --force                                 allow --out to overwrite an existing file (without it, an existing --out target refuses with exit 2)',
      '  --format <v> | -f <v>              format: html|markdown|md|pdf (default html)',
      '  --format=<v> | -f=<v>              same, equals form',
      '  --output <p> | --out <p> | -o <p>  write file (default: <input-basename>.<ext> next to input)',
      '  --output=<p> | -o=<p>              same, equals form',
      '  --image-mode <embed|file>           Markdown only: embed data URLs (default) or write folder references',
      '  --image-mode=<embed|file>           same, equals form',
      '  --asset-prefix <dir>                Markdown file mode: relative image folder next to the .md file (default media)',
      '  --asset-prefix=<dir>                same, equals form (--asset-folder is an alias)',
      '  --untagged | --untagged-pdf            PDF only: untagged render (default is tagged)',
      '  --strict                            PDF only: fail closed on preflight blockers (default warns and continues)',
      '  --allow-macros                      HTML and Markdown: preserve macro/active-content parts (default refuses E_VALIDATION naming the parts); a PDF export never opens them',
      '  --strip-macros                      HTML and Markdown: remove macro/active-content parts and continue (never with --allow-macros)',
      '  --renderer <auto|word|libreoffice>   PDF only: which program renders (default auto = LibreOffice; word is Windows only and must be asked for; env ASTRA_RENDERER)',
      '  --renderer=<v>                      same, equals form',
      '  --allow-network-references          PDF only: let the renderer load images and files the document names at network locations (default: not loaded; env ASTRA_ALLOW_NETWORK_REFERENCES=1)',
      '  --no-allow-network-references       PDF only: keep them unloaded even when the environment variable allows them',
      '  --show-markup                      PDF only: draw tracked changes and comments: Word shows inline marks and balloons, LibreOffice shows comments in the margin; pages can be scaled and page numbers refer to this render; the output shows reviewer names, deleted text and comments (env ASTRA_SHOW_MARKUP=1)',
      '  --no-show-markup                   keep the default render even when the environment variable turns markup on',
      'Headless PDF export uses LibreOffice (in PATH) or, with --renderer word on Windows, Microsoft Word.',
      'A ~11s first-render LibreOffice cold start is expected behavior, not a bug (later renders reuse the warm backend).',
    ].join('\n');
    if (!out.json) console.log(formatCliOutput(text, 'stdout'));
    if (out.json) console.log(stringifyJsonPayload({ ok: true, help: true, command: 'export' }));
    return EXIT_OK;
  }

  let docxPath: string | undefined;
  let format = 'html';
  let outPath: string | undefined;
  let force = false;
  let untagged = false;
  let strict = false;
  let imageMode: string | undefined;
  let assetPrefix: string | undefined;
  let allowMacros = false;
  let stripMacros = false;
  const renderFlags: RenderFlags = {};

  for (let i = 0; i < args.length; i++) {
    const a = args[i];
    const rf = readRenderFlag('export', args, i, renderFlags, { pdfToImage: false });
    if (rf.matched) {
      if ('error' in rf) return usage(out, rf.error);
      i = rf.next;
      continue;
    }
    if (a === '--format' || a === '-f') {
      const next = args[++i];
      if (!next) return usage(out, '--format requires an argument (html, markdown, or pdf)');
      format = next;
    } else if (a === '--image-mode') {
      const next = args[++i];
      if (!next) return usage(out, '--image-mode requires an argument (embed or file)');
      imageMode = next;
    } else if (a === '--asset-prefix' || a === '--asset-folder') {
      const next = args[++i];
      if (!next) return usage(out, '--asset-prefix requires a value (relative folder such as media)');
      assetPrefix = next;
    } else if (a === '--untagged' || a === '--untagged-pdf') {
      untagged = true;
    } else if (a === '--force') {
      force = true;
    } else if (a === '--strict') {
      strict = true;
    } else if (a === '--allow-macros') {
      allowMacros = true;
    } else if (a === '--strip-macros') {
      stripMacros = true;
    } else if (a === '--output' || a === '-o' || a === '--out') {
      const next = args[++i];
      if (!next) return usage(out, '--output requires a path argument');
      outPath = next;
    } else if (a.startsWith('--format=')) {
      format = a.slice('--format='.length);
    } else if (a.startsWith('--image-mode=')) {
      imageMode = a.slice('--image-mode='.length);
    } else if (a.startsWith('--asset-prefix=')) {
      assetPrefix = a.slice('--asset-prefix='.length);
    } else if (a.startsWith('--asset-folder=')) {
      assetPrefix = a.slice('--asset-folder='.length);
    } else if (a.startsWith('-f=')) {
      format = a.slice('-f='.length);
    } else if (a.startsWith('--output=')) {
      outPath = a.slice('--output='.length);
    } else if (a.startsWith('-o=')) {
      outPath = a.slice('-o='.length);
    } else if (a.startsWith('-')) {
      return usage(out, `unknown export flag "${a}"`);
    } else {
      if (docxPath === undefined) {
        docxPath = a;
      } else {
        return usage(out, `unexpected positional argument "${a}"`);
      }
    }
  }

  if (!docxPath) {
    return usage(out, 'export requires a <docx> argument');
  }

  if (allowMacros && stripMacros) return usage(out, macroFlagsUsageError());
  const macroOpts = macroOptsFromFlags({ allowMacros, stripMacros });

  const normalizedFormat = (format || '').toLowerCase();
  if (!['html', 'markdown', 'md', 'pdf'].includes(normalizedFormat)) {
    return usage(out, `unsupported format "${format}". Choose from: html, markdown, md, pdf`);
  }
  if (normalizedFormat !== 'pdf' && (renderFlags.renderer !== undefined || renderFlags.allowNetworkReferences !== undefined || renderFlags.showMarkup !== undefined)) {
    return usage(out, 'export: --renderer, --allow-network-references and --show-markup apply to PDF export only (use --format pdf)');
  }
  if (imageMode !== undefined) {
    const normMode = imageMode.toLowerCase();
    if (normMode !== 'embed' && normMode !== 'file') {
      return usage(out, `unsupported image mode "${imageMode}". Choose from: embed, file`);
    }
    imageMode = normMode;
  }
  if (outPath !== undefined && !force) {
    try {
      await fsp.access(outPath);
      return usage(out, `export: --out target "${path.basename(outPath)}" already exists (refusing to overwrite without --force)`);
    } catch {
      /* destination is free — proceed */
    }
  }

  try {
    const result = await executeExport(docxPath, format, outPath, {
      untagged,
      strict,
      ...(macroOpts !== undefined ? { macroOpts } : {}),
      ...(imageMode !== undefined ? { imageMode: imageMode as 'embed' | 'file' } : {}),
      ...(assetPrefix !== undefined ? { assetPrefix } : {}),
      ...renderOptionsFromFlags(renderFlags),
    });
    print(out, `EXPORT OK ${docxPath} -> ${result.path} (${result.format.toUpperCase()}, ${result.bytes} bytes)`);
    if (result.note) print(out, `note: ${result.note}`);
    printRenderNotes(result);
    printJson(out, { command: 'export', ...result });
    return EXIT_OK;
  } catch (err: unknown) {
    const msg = err instanceof Error ? err.message : String(err);
    if (msg.includes('E_DEPENDENCY')) {
      return failHuman(out, err, { file: docxPath, fallbackCode: 'E_DEPENDENCY' });
    }
    return failHuman(out, err, { file: docxPath, fallbackCode: 'E_IO' });
  }
}
