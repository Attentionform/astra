import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import {
  convertAstToHtml,
  htmlDocumentTitle,
  convertAstToMarkdown,
  collectMarkdownImageAssets,
  prepareDocxForPdfExport,
  convertDocxToPdfDetailed,
  DocxParser,
  type MarkupReport,
  type ReferenceReport,
  type RendererReport,
} from '@attentionform/astra-engine';
import { canonicalizeForAsset, isPathInsideDir } from '@attentionform/astra-engine';

export interface ExportResult {
  ok: boolean;
  path: string;
  format: 'html' | 'markdown' | 'pdf';
  bytes: number;
  assets?: string[];
  /** Set when the PDF was made from a corrected temporary copy of the source. */
  note?: string;
  /** PDF only: which program made the PDF, and whether `auto` had to fall back. */
  renderer?: RendererReport;
  /** PDF only: what happened to references the document makes to network locations. */
  references?: ReferenceReport;
  /** PDF only, and only when markup was requested: what was asked and what the renderer was told to do. */
  markup?: MarkupReport;
}

export interface ExportOptions {
  /** Render untagged PDF (default false = tagged render). */
  untagged?: boolean;
  /** Fail closed on preflight blockers (default false = warn-only). */
  strict?: boolean;
  /** Markdown only: embed data URLs (default) or write folder references. */
  imageMode?: 'embed' | 'file';
  /** Markdown file mode: relative folder for image files (default media). */
  assetPrefix?: string;
  /** Macro choice for the HTML and Markdown paths (the PDF path never opens the package for editing). */
  macroOpts?: { allowMacros: true } | { stripMacros: true };
  /** PDF only: which program renders (`auto`, `word` on Windows, `libreoffice`). The environment's ASTRA_RENDERER applies when omitted. */
  renderer?: 'auto' | 'word' | 'libreoffice';
  /** PDF only: let the renderer contact network locations the document references (default off). */
  allowNetworkReferences?: boolean;
  /** PDF only: draw tracked changes and comments (default off; env ASTRA_SHOW_MARKUP). */
  showMarkup?: boolean;
}

// Q0a: jail primitives live in @attentionform/astra-engine sandbox/paths (imported above).

function assertDestinationJailed(destination: string): { destDir: string; canonicalDest: string } {
  const destDir = path.dirname(destination);
  const canonicalDest = canonicalizeForAsset(path.normalize(destDir));
  try {
    const st = fs.lstatSync(destination);
    if (st.isSymbolicLink()) {
      throw Object.assign(
        new Error(`E_SANDBOX: output path "${path.basename(destination)}" is a symlink. Remedy: remove the symlink at the output path and write inside the export folder (paths are canonicalized with realpath).`),
        { code: 'E_SANDBOX' }
      );
    }
  } catch (err) {
    if ((err as { code?: string })?.code === 'E_SANDBOX') throw err;
  }
  const canonicalDestination = canonicalizeForAsset(path.normalize(destination));
  if (!isPathInsideDir(canonicalDestination, canonicalDest) && canonicalDestination !== canonicalDest) {
    throw Object.assign(
      new Error(`E_SANDBOX: output path "${path.basename(destination)}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and write inside the export folder (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
  return { destDir, canonicalDest };
}

function assertAssetJailed(assetName: string, destDir: string, canonicalDest: string): void {
  const resolved = path.resolve(destDir, assetName);
  if (!isPathInsideDir(resolved, path.normalize(destDir))) {
    throw Object.assign(
      new Error(`E_SANDBOX: asset path "${assetName}" escapes the export folder. Remedy: keep --asset-prefix inside the export folder such as media.`),
      { code: 'E_SANDBOX' }
    );
  }
  const canonicalCandidate = canonicalizeForAsset(resolved);
  if (!isPathInsideDir(canonicalCandidate, canonicalDest) && canonicalCandidate !== canonicalDest) {
    throw Object.assign(
      new Error(`E_SANDBOX: asset path "${assetName}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and keep --asset-prefix inside the export folder such as media (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
}

async function assertDestinationPostWrite(destination: string, canonicalDest: string): Promise<void> {
  let canonicalWritten: string;
  try {
    canonicalWritten = fs.realpathSync(destination);
  } catch {
    canonicalWritten = canonicalizeForAsset(destination);
  }
  if (!isPathInsideDir(canonicalWritten, canonicalDest) && canonicalWritten !== canonicalDest) {
    try {
      await fsp.unlink(destination);
    } catch { /* best-effort cleanup of the just-written bytes */ }
    throw Object.assign(
      new Error(`E_SANDBOX: output path "${path.basename(destination)}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and write inside the export folder (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
}

function refuseSymlinkNode(target: string, label: string): void {
  let st: fs.Stats;
  try {
    st = fs.lstatSync(target);
  } catch {
    return;
  }
  if (st.isSymbolicLink()) {
    throw Object.assign(
      new Error(`E_SANDBOX: ${label} "${path.basename(target)}" is a symlink and was refused. Remedy: remove the symlink inside the export folder and write inside the export folder (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
}

function assertPostMkdirAssetContainment(resolvedAsset: string, canonicalDest: string): void {
  const parent = path.dirname(resolvedAsset);
  refuseSymlinkNode(parent, 'asset folder');
  let canonicalParent: string;
  try {
    canonicalParent = fs.realpathSync(parent);
  } catch {
    canonicalParent = canonicalizeForAsset(parent);
  }
  if (!isPathInsideDir(canonicalParent, canonicalDest) && canonicalParent !== canonicalDest) {
    throw Object.assign(
      new Error(`E_SANDBOX: asset path "${path.basename(resolvedAsset)}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and keep --asset-prefix inside the export folder such as media (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
  const canonicalAsset = canonicalizeForAsset(resolvedAsset);
  if (!isPathInsideDir(canonicalAsset, canonicalDest) && canonicalAsset !== canonicalDest) {
    throw Object.assign(
      new Error(`E_SANDBOX: asset path "${path.basename(resolvedAsset)}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and keep --asset-prefix inside the export folder such as media (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
}

function assertPreWriteDestination(destination: string, destDir: string, canonicalDest: string): void {
  refuseSymlinkNode(destination, 'output path');
  refuseSymlinkNode(destDir, 'export folder');
  let canonicalParent: string;
  try {
    canonicalParent = fs.realpathSync(destDir);
  } catch {
    canonicalParent = canonicalizeForAsset(path.normalize(destDir));
  }
  if (!isPathInsideDir(canonicalParent, canonicalDest) && canonicalParent !== canonicalDest) {
    throw Object.assign(
      new Error(`E_SANDBOX: output path "${path.basename(destination)}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and write inside the export folder (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
  const canonicalDestination = canonicalizeForAsset(path.normalize(destination));
  if (!isPathInsideDir(canonicalDestination, canonicalDest) && canonicalDestination !== canonicalDest) {
    throw Object.assign(
      new Error(`E_SANDBOX: output path "${path.basename(destination)}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and write inside the export folder (paths are canonicalized with realpath).`),
      { code: 'E_SANDBOX' }
    );
  }
}

export async function executeExport(
  docxPath: string,
  formatInput: string = 'html',
  outPath?: string,
  options?: ExportOptions
): Promise<ExportResult> {
  const normFormat = (formatInput || 'html').toLowerCase();
  let format: 'html' | 'markdown' | 'pdf';
  if (normFormat === 'html') format = 'html';
  else if (normFormat === 'md' || normFormat === 'markdown') format = 'markdown';
  else if (normFormat === 'pdf') format = 'pdf';
  else {
    throw new Error(`Unsupported export format "${formatInput}". Expected html, markdown, or pdf.`);
  }

  const resolvedDocx = path.resolve(docxPath);
  if (!fs.existsSync(resolvedDocx)) {
    throw new Error(`ENOENT: no such file or directory, open '${docxPath}'`);
  }

  const extMap: Record<string, string> = {
    html: '.html',
    markdown: '.md',
    pdf: '.pdf',
  };

  const destination = outPath
    ? path.resolve(outPath)
    : path.join(path.dirname(resolvedDocx), `${path.basename(resolvedDocx, path.extname(resolvedDocx))}${extMap[format]}`);

  assertDestinationJailed(destination);

  if (format === 'pdf') {
    if (options?.imageMode !== undefined && options.imageMode !== 'embed') {
      throw new Error('E_VALIDATION: image mode applies to markdown only. Remedy: omit --image-mode or export with --format markdown.');
    }
    const buf = await fsp.readFile(resolvedDocx);
    // A valid file is rendered as it is; a file whose only errors are element order is rendered from a
    // verified temporary copy. The source file is never written.
    const prepared = await prepareDocxForPdfExport(buf, 'PDF export source');
    await fsp.mkdir(path.dirname(destination), { recursive: true });
    const { canonicalDest: pdfDestDir } = assertDestinationJailed(destination);
    assertPreWriteDestination(destination, path.dirname(destination), pdfDestDir);
    const converted = await convertDocxToPdfDetailed(prepared.buffer, destination, {
      tagged: !(options?.untagged ?? false),
      strict: options?.strict ?? false,
      ...(options?.renderer !== undefined ? { renderer: options.renderer } : {}),
      ...(options?.allowNetworkReferences !== undefined ? { allowNetworkReferences: options.allowNetworkReferences } : {}),
      ...(options?.showMarkup !== undefined ? { showMarkup: options.showMarkup } : {}),
    });
    await assertDestinationPostWrite(destination, pdfDestDir);
    const stat = await fsp.stat(destination);
    return {
      ok: true,
      path: destination,
      format,
      bytes: stat.size,
      renderer: converted.renderer,
      references: converted.references,
      ...(converted.markup ? { markup: converted.markup } : {}),
      ...(prepared.repaired
        ? {
            note:
              `the source has ${prepared.reorderedContainers} element-order error(s) in ${prepared.repairedParts.join(', ')}; ` +
              `the PDF was made from a corrected temporary copy, and the file was not changed`,
          }
        : {}),
    };
  }

  const buf = await fsp.readFile(resolvedDocx);
  const parser = new DocxParser(buf, options?.macroOpts);
  const ast = await parser.getAST();

  if (format === 'html') {
    if (options?.imageMode !== undefined && options.imageMode !== 'embed') {
      throw new Error('E_VALIDATION: image mode applies to markdown only. Remedy: omit --image-mode or export with --format markdown.');
    }
    const title = htmlDocumentTitle(ast, path.basename(resolvedDocx, path.extname(resolvedDocx)));
    const html = convertAstToHtml(ast, title, parser);
    await fsp.mkdir(path.dirname(destination), { recursive: true });
    const { canonicalDest: htmlDestDir } = assertDestinationJailed(destination);
    assertPreWriteDestination(destination, path.dirname(destination), htmlDestDir);
    await fsp.writeFile(destination, html, 'utf-8');
    await assertDestinationPostWrite(destination, htmlDestDir);
    const stat = await fsp.stat(destination);
    return { ok: true, path: destination, format, bytes: stat.size };
  } else {
    const imageMode = options?.imageMode ?? 'embed';
    if (imageMode !== 'embed' && imageMode !== 'file') {
      throw new Error(`E_VALIDATION: image mode "${String(options?.imageMode)}" is not supported. Remedy: pass --image-mode embed or file.`);
    }
    if (imageMode === 'embed') {
      const md = convertAstToMarkdown(ast, parser, { imageMode: 'embed' });
      await fsp.mkdir(path.dirname(destination), { recursive: true });
      const { canonicalDest: mdDestDir } = assertDestinationJailed(destination);
      assertPreWriteDestination(destination, path.dirname(destination), mdDestDir);
      await fsp.writeFile(destination, md, 'utf-8');
      await assertDestinationPostWrite(destination, mdDestDir);
      const stat = await fsp.stat(destination);
      return { ok: true, path: destination, format, bytes: stat.size };
    }
    const md = convertAstToMarkdown(ast, parser, { imageMode: 'file', assetPrefix: options?.assetPrefix });
    const assets = collectMarkdownImageAssets(ast, options?.assetPrefix);
    const destDir = path.dirname(destination);
    const canonicalDest = canonicalizeForAsset(path.normalize(destDir));
    assertDestinationJailed(destination);
    for (const asset of assets) {
      assertAssetJailed(asset.name, destDir, canonicalDest);
    }
    await fsp.mkdir(destDir, { recursive: true });
    const canonicalDestPost = canonicalizeForAsset(path.normalize(destDir));
    assertDestinationJailed(destination);
    for (const asset of assets) {
      assertAssetJailed(asset.name, destDir, canonicalDestPost);
    }
    assertPreWriteDestination(destination, destDir, canonicalDestPost);
    await fsp.writeFile(destination, md, 'utf-8');
    await assertDestinationPostWrite(destination, canonicalDestPost);
    const written: string[] = [];
    for (const asset of assets) {
      const resolved = path.resolve(destDir, asset.name);
      await fsp.mkdir(path.dirname(resolved), { recursive: true });
      assertPostMkdirAssetContainment(resolved, canonicalDestPost);
      await fsp.writeFile(resolved, asset.data);
      let canonicalWritten: string;
      try {
        canonicalWritten = fs.realpathSync(resolved);
      } catch {
        canonicalWritten = canonicalizeForAsset(resolved);
      }
      if (!isPathInsideDir(canonicalWritten, canonicalDestPost) && canonicalWritten !== canonicalDestPost) {
        try {
          await fsp.unlink(resolved);
        } catch { /* best-effort cleanup of the just-written bytes */ }
        throw Object.assign(
          new Error(`E_SANDBOX: asset path "${asset.name}" resolves outside the export folder via symlink. Remedy: remove the symlink inside the export folder and keep --asset-prefix inside the export folder such as media (paths are canonicalized with realpath).`),
          { code: 'E_SANDBOX' }
        );
      }
      written.push(resolved);
    }
    const stat = await fsp.stat(destination);
    return { ok: true, path: destination, format, bytes: stat.size, assets: written };
  }
}
