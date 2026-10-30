import { describe, it, expect, afterEach } from 'vitest';
import fs from 'fs';
import os from 'os';
import path from 'path';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser } from '../src/parser';
import { compileMarkdown, composeDocument } from '../src/compose/markdown';

describe('markdown post-table typographic separation', () => {
  let tempDir: string | null = null;

  afterEach(() => {
    if (tempDir && fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true, force: true });
      tempDir = null;
    }
  });

  const png1x1 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

  function createAssetDir(): string {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-table-sep-'));
    tempDir = dir;
    fs.writeFileSync(path.join(dir, 'sample.png'), Buffer.from(png1x1, 'base64'));
    fs.writeFileSync(
      path.join(dir, 'sample.chart.json'),
      JSON.stringify({ chartType: 'bar', title: 'Metrics', categories: ['A'], series: [{ name: 'S1', values: [10] }] })
    );
    return dir;
  }

  it('injects spacingBefore 240 when paragraph follows a table', () => {
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\nParagraph after table.';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('does NOT inject spacingBefore when H1 or H2 follows a table', () => {
    const mdH1 = '| A | B |\n|---|---|\n| 1 | 2 |\n\n# Heading One';
    const opsH1 = compileMarkdown(mdH1);
    const styleOpsH1 = opsH1.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOpsH1.length).toBe(0);

    const mdH2 = '| A | B |\n|---|---|\n| 1 | 2 |\n\n## Heading Two';
    const opsH2 = compileMarkdown(mdH2);
    const styleOpsH2 = opsH2.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOpsH2.length).toBe(0);
  });

  it('injects spacingBefore 240 when H3 follows a table', () => {
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\n### Heading Three';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('injects spacingBefore 240 on first item when list follows a table', () => {
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\n- First item\n- Second item';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('injects spacingBefore 240 when code block follows a table', () => {
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\n```python\nx = 42\n```';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('injects spacingBefore 240 when image follows a table', () => {
    const dir = createAssetDir();
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\n![Image](sample.png)';
    const ops = compileMarkdown(md, { allowedDir: dir });

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('injects spacingBefore 240 when chart follows a table', () => {
    const dir = createAssetDir();
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\n![Chart](sample.chart.json)';
    const ops = compileMarkdown(md, { allowedDir: dir });

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('injects spacingBefore 240 when divider follows a table', () => {
    const md = '| A | B |\n|---|---|\n| 1 | 2 |\n\n---';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(1);
    expect(styleOps[0]).toMatchObject({
      action: 'style',
      target: 'md_0002',
      spacingBefore: 240,
    });
  });

  it('does NOT inject spacingBefore when table follows a table', () => {
    const md = '| A |\n|---|\n| 1 |\n\n| B |\n|---|\n| 2 |';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(0);
  });

  it('does NOT inject spacingBefore when preceding block is not a table', () => {
    const md = '# Title\n\nFirst para\n\nSecond para\n\n- List item\n\n```js\nconsole.log(1);\n```\n\n---';
    const ops = compileMarkdown(md);

    const styleOps = ops.filter((op) => op.action === 'style' && (op as { spacingBefore?: number }).spacingBefore !== undefined);
    expect(styleOps.length).toBe(0);
  });

  it('applies spacingBefore 240 in rendered document.xml when composed', async () => {
    const parser = DocxParser.createNewDocument('blank');
    const md = '| Header |\n| --- |\n| Cell |\n\nPost-table paragraph';
    const result = await composeDocument(parser, md);
    expect(result.success).toBe(true);

    const zip = new AdmZip(parser.rezip());
    const docXml = zip.getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
    const doc = new DOMParser().parseFromString(docXml, 'text/xml');

    const paras = Array.from(doc.getElementsByTagName('w:p'));
    const postPara = paras.find((p) => {
      const textNodes = Array.from(p.getElementsByTagName('w:t'));
      return textNodes.some((t) => t.textContent?.includes('Post-table paragraph'));
    });
    expect(postPara).toBeDefined();

    const spacing = postPara!.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('240');
  });
});
