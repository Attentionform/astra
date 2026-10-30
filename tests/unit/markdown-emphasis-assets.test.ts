import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as fsp from 'fs/promises';
import * as path from 'path';
import * as os from 'os';
import { DOMParser } from '@xmldom/xmldom';
import {
  convertAstToMarkdown,
  collectMarkdownImageAssets,
} from '../../packages/engine/src/export/markdown';
import type { DocumentAST } from '../../packages/engine/src/ast';
import { DocxParser } from '../../packages/engine/src';
import { composeDocument } from '../../packages/engine/src/compose/markdown';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';
import { executeExport } from '../../packages/cli/src/export';
import { useReproducibleOutput } from '../helpers/reproducible';

useReproducibleOutput();

function parseXml(xmlStr: string): Element {
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  return doc.documentElement;
}

function buildAst(nodes: any[], extra: any = {}): DocumentAST {
  return {
    metadata: {},
    settings: { trackChanges: true, defaultAuthor: 'Author' },
    nextChangeId: 1,
    styles: {},
    numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {},
    headers: {},
    footers: {},
    mediaFiles: {},
    sections: [
      {
        id: 'sec_1',
        heading: null,
        pageSetup: null,
        nodes,
        version: 1,
        checksum: 'sec_chk_1',
      },
    ],
    ...extra,
  } as unknown as DocumentAST;
}

function buildMeta(text: string, extra: any = {}): any {
  return {
    text,
    hasCitation: false,
    hasComment: false,
    hasBookmark: false,
    hasTrackedChanges: false,
    hasEquation: false,
    hasField: false,
    hasImage: false,
    hasHyperlink: false,
    ...extra,
  };
}

function paraNode(xml: Element, id = 'p_1', metaExtra: any = {}): any {
  return {
    id,
    type: 'paragraph',
    xml,
    metadata: buildMeta('', metaExtra),
    version: 1,
    checksum: `chk_${id}`,
  };
}

const PNG_B64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

function imageAst(extraNodes: any[] = [], altDescr = 'Sample pic'): DocumentAST {
  const imgXml = parseXml(
    '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      '<w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="' +
      altDescr +
      '"/>' +
      '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
      '</wp:inline></w:drawing></w:r></w:p>'
  );
  const pngBytes = Buffer.from(PNG_B64, 'base64');
  return buildAst(
    [
      {
        id: 'img_1',
        type: 'image',
        xml: imgXml,
        metadata: buildMeta('', { hasImage: true }),
        version: 1,
        checksum: 'chk_img_1',
      },
      ...extraNodes,
    ],
    {
      relationships: { rId5: { target: 'media/image1.png' } },
      mediaFiles: { 'image1.png': pngBytes },
    }
  );
}

describe('emphasis runs export as markdown spans', () => {
  it('renders bold runs with double markers', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t>Hello</w:t></w:r></w:p>'
    );
    expect(convertAstToMarkdown(buildAst([paraNode(xml)]))).toContain('**Hello**');
  });

  it('renders italic runs with single markers', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:i/></w:rPr><w:t>World</w:t></w:r></w:p>'
    );
    expect(convertAstToMarkdown(buildAst([paraNode(xml)]))).toContain('*World*');
  });

  it('nests bold and italic markers together', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>Both</w:t></w:r></w:p>'
    );
    expect(convertAstToMarkdown(buildAst([paraNode(xml)]))).toContain('***Both***');
  });

  it('renders monospace runs as code spans', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr><w:t>code1</w:t></w:r></w:p>'
    );
    expect(convertAstToMarkdown(buildAst([paraNode(xml)]))).toContain('`code1`');
  });

  it('keeps combined code with emphasis markers outside the span', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:b/><w:i/><w:rFonts w:ascii="Courier New"/></w:rPr><w:t>combo</w:t></w:r></w:p>'
    );
    expect(convertAstToMarkdown(buildAst([paraNode(xml)]))).toContain('***`combo`***');
  });

  it('moves edge spaces outside emphasis markers', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve"> hello</w:t></w:r>' +
        '<w:r><w:t xml:space="preserve"> world</w:t></w:r></w:p>'
    );
    const md = convertAstToMarkdown(buildAst([paraNode(xml)]));
    expect(md).toBe('**hello** world\n');
    expect(md).not.toContain('** hello');
  });

  it('leaves whitespace-only styled runs unmarked', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">   </w:t></w:r></w:p>'
    );
    expect(convertAstToMarkdown(buildAst([paraNode(xml)]))).not.toContain('****');
  });

  it('escapes markdown metacharacters in plain runs', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:t>a*b_c[d]e</w:t></w:r></w:p>'
    );
    const md = convertAstToMarkdown(buildAst([paraNode(xml)]));
    expect(md).toContain('a\\*b\\_c\\[d\\]e');
  });

  it('renders single-line code style once for monospace content', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:pPr><w:pStyle w:val="Code"/></w:pPr>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr><w:t>hello</w:t></w:r></w:p>'
    );
    const md = convertAstToMarkdown(buildAst([paraNode(xml, 'p_1', { namedStyle: 'code' })]));
    expect(md.trim()).toBe('`hello`');
  });
});

describe('image export modes', () => {
  it('embeds data URLs by default and matches explicit embed output', () => {
    const ast = imageAst();
    const implicit = convertAstToMarkdown(ast);
    const explicit = convertAstToMarkdown(ast, { imageMode: 'embed' });
    expect(implicit).toBe(explicit);
    expect(implicit).toContain('data:image/png;base64,');
  });

  it('emits folder references plus matching assets in file mode', () => {
    const ast = imageAst();
    const md = convertAstToMarkdown(ast, { imageMode: 'file' });
    expect(md.trim()).toBe('![Sample pic](media/image-0001.png)');
    const assets = collectMarkdownImageAssets(ast);
    expect(assets).toHaveLength(1);
    expect(assets[0].name).toBe('media/image-0001.png');
    expect(assets[0].alt).toBe('Sample pic');
    expect(assets[0].data.equals(Buffer.from(PNG_B64, 'base64'))).toBe(true);
  });

  it('honors a custom asset prefix in both text and assets', () => {
    const ast = imageAst();
    const md = convertAstToMarkdown(ast, { imageMode: 'file', assetPrefix: 'assets' });
    expect(md.trim()).toBe('![Sample pic](assets/image-0001.png)');
    const assets = collectMarkdownImageAssets(ast, 'assets');
    expect(assets[0].name).toBe('assets/image-0001.png');
  });

  it('numbers multiple images deterministically with matching extensions', () => {
    const secondXml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<w:r><w:drawing><wp:inline><wp:docPr id="2" name="Picture 2" descr="Second"/>' +
        '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId6"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
        '</wp:inline></w:drawing></w:r></w:p>'
    );
    const gifBytes = Buffer.from('R0lGODdhAQABAIAAAP///////ywAAAAAAQABAAACAkQBADs=', 'base64');
    const ast = buildAst(
      [
        {
          id: 'img_1',
          type: 'image',
          xml: parseXml(
            '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
              '<w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="First"/>' +
              '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
              '</wp:inline></w:drawing></w:r></w:p>'
          ),
          metadata: buildMeta('', { hasImage: true }),
          version: 1,
          checksum: 'chk_img_1',
        },
        {
          id: 'img_2',
          type: 'image',
          xml: secondXml,
          metadata: buildMeta('', { hasImage: true }),
          version: 1,
          checksum: 'chk_img_2',
        },
      ],
      {
        relationships: { rId5: { target: 'media/a.png' }, rId6: { target: 'media/b.gif' } },
        mediaFiles: {
          'a.png': Buffer.from(PNG_B64, 'base64'),
          'b.gif': gifBytes,
        },
      }
    );
    const md = convertAstToMarkdown(ast, { imageMode: 'file' });
    expect(md).toContain('![First](media/image-0001.png)');
    expect(md).toContain('![Second](media/image-0002.gif)');
    const assets = collectMarkdownImageAssets(ast);
    expect(assets.map((entry) => entry.name)).toEqual([
      'media/image-0001.png',
      'media/image-0002.gif',
    ]);
  });

  it('keeps fallback text for images without bytes in both modes', () => {
    const xml = parseXml(
      '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
        '<w:r><w:t>caption only</w:t></w:r></w:p>'
    );
    const ast = buildAst([
      {
        id: 'img_9',
        type: 'image',
        xml,
        metadata: buildMeta('caption only', { hasImage: true }),
        version: 1,
        checksum: 'chk_img_9',
      },
    ]);
    expect(convertAstToMarkdown(ast).trim()).toBe('[Image: caption only]');
    expect(convertAstToMarkdown(ast, { imageMode: 'file' }).trim()).toBe('[Image: caption only]');
    expect(collectMarkdownImageAssets(ast)).toHaveLength(0);
  });

  it('rejects unknown image modes with a remedy', () => {
    const ast = imageAst();
    expect(() => convertAstToMarkdown(ast, { imageMode: 'blob' as never })).toThrow(
      /E_VALIDATION.*Remedy/
    );
  });

  it('rejects escaping prefixes with a remedy', () => {
    const ast = imageAst();
    expect(() => convertAstToMarkdown(ast, { imageMode: 'file', assetPrefix: '../evil' })).toThrow(
      /E_VALIDATION.*Remedy/
    );
    expect(() => collectMarkdownImageAssets(ast, '/abs')).toThrow(/E_VALIDATION.*Remedy/);
    expect(() => collectMarkdownImageAssets(ast, '.hidden')).toThrow(/E_VALIDATION.*Remedy/);
  });

  it('strips brackets from file alt text', () => {
    const ast = imageAst([], 'Pic [bracket] shot');
    const md = convertAstToMarkdown(ast, { imageMode: 'file' });
    expect(md).not.toContain('['.repeat(2));
    expect(md).toContain('![Pic bracket shot](media/image-0001.png)');
  });
});

const IMAGE_DOC = path.resolve(__dirname, '../../test-gallery/T5_image_doc.docx');

async function makeTempDir(prefix: string): Promise<string> {
  return fsp.mkdtemp(path.join(os.tmpdir(), prefix));
}

describe('cli file-mode export writes markdown plus asset files', () => {
  it('writes folder references and matching asset bytes', async () => {
    const dir = await makeTempDir('astra-export-file-');
    try {
      const outMd = path.join(dir, 'out.md');
      const result = await executeExport(IMAGE_DOC, 'markdown', outMd, { imageMode: 'file' });
      expect(result.ok).toBe(true);
      expect(result.format).toBe('markdown');
      expect(result.assets).toHaveLength(1);
      const md = await fsp.readFile(outMd, 'utf-8');
      expect(md).toContain('media/image-0001.png');
      const assetPath = path.join(dir, 'media', 'image-0001.png');
      expect(fs.existsSync(assetPath)).toBe(true);
      const written = await fsp.readFile(assetPath);
      expect(written.length).toBeGreaterThan(0);
      expect(result.assets?.[0]).toBe(assetPath);
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });

  it('refuses folder image mode for html with a remedy', async () => {
    const dir = await makeTempDir('astra-export-html-');
    try {
      const outHtml = path.join(dir, 'out.html');
      await expect(executeExport(IMAGE_DOC, 'html', outHtml, { imageMode: 'file' })).rejects.toThrow(
        /E_VALIDATION.*Remedy/
      );
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });

  it('refuses folder image mode for pdf with a remedy', async () => {
    const dir = await makeTempDir('astra-export-pdf-');
    try {
      const outPdf = path.join(dir, 'out.pdf');
      await expect(executeExport(IMAGE_DOC, 'pdf', outPdf, { imageMode: 'file' })).rejects.toThrow(
        /E_VALIDATION.*Remedy/
      );
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });

  it('refuses a pre-planted symlink folder with a canonicalization remedy and writes nothing outside', async () => {
    const dir = await makeTempDir('astra-export-jail-');
    const outside = await makeTempDir('astra-export-outside-');
    try {
      const outMd = path.join(dir, 'out.md');
      await fsp.mkdir(path.join(dir, 'media'), { recursive: true });
      await fsp.rm(path.join(dir, 'media'), { recursive: true, force: true });
      fs.symlinkSync(outside, path.join(dir, 'media'), 'dir');
      await expect(executeExport(IMAGE_DOC, 'markdown', outMd, { imageMode: 'file' })).rejects.toThrow(
        /E_SANDBOX.*realpath/
      );
      expect(fs.existsSync(outMd)).toBe(false);
      const outsideEntries = await fsp.readdir(outside);
      expect(outsideEntries).toEqual([]);
    } finally {
      try {
        await fsp.rm(path.join(dir, 'media'), { force: true });
      } catch { /* symlink cleanup best effort */ }
      await fsp.rm(dir, { recursive: true, force: true });
      await fsp.rm(outside, { recursive: true, force: true });
    }
  });

  it('treats a trailing slash prefix as the bare folder name', async () => {
    const dir = await makeTempDir('astra-export-slash-');
    try {
      const plainMd = path.join(dir, 'plain.md');
      const slashedMd = path.join(dir, 'slashed.md');
      await executeExport(IMAGE_DOC, 'markdown', plainMd, { imageMode: 'file', assetPrefix: 'media' });
      await executeExport(IMAGE_DOC, 'markdown', slashedMd, {
        imageMode: 'file',
        assetPrefix: 'media/',
      });
      const plain = await fsp.readFile(plainMd, 'utf-8');
      const slashed = await fsp.readFile(slashedMd, 'utf-8');
      expect(slashed).toBe(plain);
      expect(slashed).toContain('media/image-0001.png');
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });

  it('treats a backslash prefix as the forward-slash folder name', async () => {
    const dir = await makeTempDir('astra-export-backslash-');
    try {
      const plainMd = path.join(dir, 'plain.md');
      const backMd = path.join(dir, 'back.md');
      await executeExport(IMAGE_DOC, 'markdown', plainMd, { imageMode: 'file', assetPrefix: 'media' });
      await executeExport(IMAGE_DOC, 'markdown', backMd, {
        imageMode: 'file',
        assetPrefix: 'media\\',
      });
      const plain = await fsp.readFile(plainMd, 'utf-8');
      const back = await fsp.readFile(backMd, 'utf-8');
      expect(back).toBe(plain);
      expect(back).toContain('media/image-0001.png');
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });

  it('keeps a markdown to docx to markdown cycle byte-identical', async () => {
    const dir = await makeTempDir('astra-export-stable-');
    try {
      const manuscript = [
        '# Field Notes on Endurance Training',
        '',
        'Opening paragraph sets the scene for the training block.',
        '',
        '## Weekly Structure',
        '',
        '- warm up easy',
        '- main set steady',
        '',
        '1. first step',
        '2. second step',
        '',
        '| Metric | Value |',
        '| --- | --- |',
        '| Endurance | 42 |',
        '',
        '---',
        '',
        '```text',
        'log line one',
        '```',
        '',
        'A note reference[^fn] in running text.',
        '',
        '[^fn]: The footnote body.',
        '',
      ].join('\n');
      const firstParser = DocxParser.createNewDocument('blank');
      const firstCompose = await composeDocument(firstParser, manuscript);
      expect(firstCompose.success).toBe(true);
      const firstDocx = path.join(dir, 'pass1.docx');
      await fsp.writeFile(firstDocx, firstParser.rezip());
      const firstMdPath = path.join(dir, 'pass1.md');
      await executeExport(firstDocx, 'markdown', firstMdPath);
      const firstMd = await fsp.readFile(firstMdPath, 'utf-8');
      const secondParser = DocxParser.createNewDocument('blank');
      const secondCompose = await composeDocument(secondParser, firstMd);
      expect(secondCompose.success).toBe(true);
      const secondDocx = path.join(dir, 'pass2.docx');
      await fsp.writeFile(secondDocx, secondParser.rezip());
      const secondMdPath = path.join(dir, 'pass2.md');
      await executeExport(secondDocx, 'markdown', secondMdPath);
      const secondMd = await fsp.readFile(secondMdPath, 'utf-8');
      expect(secondMd).toBe(firstMd);
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });

  it('reimports file-mode markdown through the export folder and stays valid', async () => {
    const dir = await makeTempDir('astra-export-reimport-');
    try {
      const outMd = path.join(dir, 'out.md');
      const exported = await executeExport(IMAGE_DOC, 'markdown', outMd, { imageMode: 'file' });
      expect(exported.assets).toHaveLength(1);
      const md = await fsp.readFile(outMd, 'utf-8');
      expect(md).toContain('media/image-0001.png');
      const parser = DocxParser.createNewDocument('blank');
      const composed = await composeDocument(parser, md, { allowedDir: dir });
      expect(composed.success).toBe(true);
      const validated = await validateDocxBuffer(parser.rezip());
      expect(validated.ok).toBe(true);
      expect(validated.errors).toEqual([]);
    } finally {
      await fsp.rm(dir, { recursive: true, force: true });
    }
  });
});
