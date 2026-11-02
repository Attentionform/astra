import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, convertAstToHtml, convertAstToMarkdown } from '@attentionform/astra-engine';

/**
 * A paragraph style such as Subtitle or Quote is italic through the style; no run carries `w:i`. The HTML and
 * Markdown exports read run properties only, so that text came out plain. Italic is now read from the run,
 * then from the paragraph style and the character style (which toggle each other, as in Word).
 */
interface Run {
  t: string;
  /** direct italic on the run: `on` is `<w:i/>`, `off` is `<w:i w:val="0"/>` */
  i?: 'on' | 'off';
  b?: boolean;
  rStyle?: string;
}

const run = (r: Run): string => {
  const props = `${r.rStyle ? `<w:rStyle w:val="${r.rStyle}"/>` : ''}${r.b ? '<w:b/>' : ''}${
    r.i === 'on' ? '<w:i/>' : r.i === 'off' ? '<w:i w:val="0"/>' : ''
  }`;
  return `<w:r>${props ? `<w:rPr>${props}</w:rPr>` : ''}<w:t xml:space="preserve">${r.t}</w:t></w:r>`;
};
const para = (style: string | undefined, runs: Run[]): string =>
  `<w:p>${style ? `<w:pPr><w:pStyle w:val="${style}"/></w:pPr>` : ''}${runs.map(run).join('')}</w:p>`;

const EXTRA_STYLES = `
<w:style w:type="paragraph" w:customStyle="1" w:styleId="Quote"><w:name w:val="Quote"/><w:basedOn w:val="Normal"/><w:rPr><w:i/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="QuoteChild"><w:name w:val="Quote Child"/><w:basedOn w:val="Quote"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="QuoteOff"><w:name w:val="Quote Off"/><w:basedOn w:val="Quote"/><w:rPr><w:i w:val="0"/></w:rPr></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="LoopA"><w:name w:val="Loop A"/><w:basedOn w:val="LoopB"/></w:style>
<w:style w:type="paragraph" w:customStyle="1" w:styleId="LoopB"><w:name w:val="Loop B"/><w:basedOn w:val="LoopA"/></w:style>
<w:style w:type="character" w:customStyle="1" w:styleId="Emph"><w:name w:val="Emph"/><w:rPr><w:i/></w:rPr></w:style>`;

function docx(paragraphs: string[]): Buffer {
  const zip = new AdmZip(DocxParser.createNewDocument('blank').rezip());
  const styles = zip.readAsText('word/styles.xml');
  zip.updateFile('word/styles.xml', Buffer.from(styles.replace('</w:styles>', `${EXTRA_STYLES}</w:styles>`), 'utf8'));
  const xml = zip.readAsText('word/document.xml');
  const start = xml.indexOf('<w:body>') + '<w:body>'.length;
  const end = xml.lastIndexOf('<w:sectPr');
  zip.updateFile('word/document.xml', Buffer.from(xml.slice(0, start) + paragraphs.join('') + xml.slice(end), 'utf8'));
  return zip.toBuffer();
}

async function exportBoth(paragraphs: string[]): Promise<{ html: string; md: string }> {
  const parser = new DocxParser(docx(paragraphs));
  const ast = await parser.getAST();
  const page = convertAstToHtml(ast, 'T', parser);
  // The page's own CSS mentions italic (table captions), so the checks read the body only.
  return { html: page.slice(page.indexOf('<body'), page.indexOf('</body>')), md: convertAstToMarkdown(ast, parser) };
}

const ITALIC = 'font-style: italic;';

describe('italic from the paragraph style', () => {
  it('a Subtitle paragraph is italic in HTML and in Markdown', async () => {
    const { html, md } = await exportBoth([para('Subtitle', [{ t: 'A quiet start' }])]);
    expect(html).toContain(`<span style="${ITALIC}">A quiet start</span>`);
    expect(md).toContain('*A quiet start*');
  });

  it('follows the basedOn chain', async () => {
    const { html, md } = await exportBoth([para('QuoteChild', [{ t: 'inherited' }])]);
    expect(html).toContain(`<span style="${ITALIC}">inherited</span>`);
    expect(md).toContain('*inherited*');
  });

  it('a style that turns italic off in its chain is plain', async () => {
    const { html, md } = await exportBoth([para('QuoteOff', [{ t: 'switched off' }])]);
    expect(html).not.toContain(ITALIC);
    expect(md).toContain('switched off');
    expect(md).not.toContain('*switched off*');
  });

  it('a loop in basedOn ends and gives plain text', async () => {
    const { html } = await exportBoth([para('LoopA', [{ t: 'loop' }])]);
    expect(html).not.toContain(ITALIC);
  });

  it('a paragraph without an italic style stays plain, as before', async () => {
    const { html, md } = await exportBoth([para(undefined, [{ t: 'ordinary text' }])]);
    expect(html).toContain('<p class="paragraph">ordinary text</p>');
    expect(md.trim()).toBe('ordinary text');
  });
});

describe('direct formatting and the character style', () => {
  it('direct italic off on a run beats an italic paragraph style', async () => {
    const { html, md } = await exportBoth([para('Quote', [{ t: 'lead ' }, { t: 'upright', i: 'off' }])]);
    expect(html).toContain(`<span style="${ITALIC}">lead </span>upright`);
    expect(md).toContain('*lead* upright');
  });

  it('direct italic on in a plain paragraph is italic, as before', async () => {
    const { html, md } = await exportBoth([para(undefined, [{ t: 'slanted', i: 'on' }])]);
    expect(html).toContain(`<span style="${ITALIC}">slanted</span>`);
    expect(md).toContain('*slanted*');
  });

  it('direct italic on in an italic paragraph stays italic (it does not toggle)', async () => {
    const { html } = await exportBoth([para('Quote', [{ t: 'both', i: 'on' }])]);
    expect(html).toContain(`<span style="${ITALIC}">both</span>`);
  });

  it('an italic character style in a plain paragraph is italic', async () => {
    const { html, md } = await exportBoth([para(undefined, [{ t: 'stress', rStyle: 'Emph' }])]);
    expect(html).toContain(`<span style="${ITALIC}">stress</span>`);
    expect(md).toContain('*stress*');
  });

  it('an italic character style in an italic paragraph toggles back to plain, as Word shows it', async () => {
    const { html, md } = await exportBoth([para('Quote', [{ t: 'lead ' }, { t: 'upright', rStyle: 'Emph' }])]);
    expect(html).toContain(`<span style="${ITALIC}">lead </span>upright`);
    expect(md).toContain('*lead* upright');
  });
});

describe('Markdown spans', () => {
  it('writes one span for neighbouring runs with the same emphasis', async () => {
    const { md } = await exportBoth([para('Subtitle', [{ t: 'Start ' }, { t: 'typing ' }, { t: 'here' }])]);
    expect(md).toContain('*Start typing here*');
    expect(md).not.toContain('**');
  });

  it('merges direct italic runs too, which used to give *a**b*', async () => {
    const { md } = await exportBoth([para(undefined, [{ t: 'one', i: 'on' }, { t: 'two', i: 'on' }])]);
    expect(md).toContain('*onetwo*');
    expect(md).not.toContain('*one**two*');
  });

  it('keeps bold and italic runs apart', async () => {
    const { md } = await exportBoth([para('Subtitle', [{ t: 'soft ' }, { t: 'loud', b: true }, { t: ' soft' }])]);
    expect(md).toContain('*soft*');
    expect(md).toContain('***loud***');
    expect(md).toContain('*soft*');
  });

  it('a Subtitle after a heading reads back as Subtitle text in italic', async () => {
    const { md } = await exportBoth([para('Heading1', [{ t: 'Chapter' }]), para('Subtitle', [{ t: 'a line below' }])]);
    expect(md).toMatch(/# Chapter\n+\*a line below\*/);
  });
});

describe('exports without a parser', () => {
  it('read direct formatting only, as before, and do not fail', async () => {
    const parser = new DocxParser(docx([para('Subtitle', [{ t: 'by style' }]), para(undefined, [{ t: 'by hand', i: 'on' }])]));
    const ast = await parser.getAST();
    const page = convertAstToHtml(ast, 'T');
    const body = page.slice(page.indexOf('<body'), page.indexOf('</body>'));
    expect(body).toContain('<p class="paragraph">by style</p>');
    expect(body).toContain(`<span style="${ITALIC}">by hand</span>`);
    expect(convertAstToMarkdown(ast)).toContain('*by hand*');
  });

  it('forget the styles of an earlier export of the same AST', async () => {
    const parser = new DocxParser(docx([para('Subtitle', [{ t: 'by style' }])]));
    const ast = await parser.getAST();
    expect(convertAstToMarkdown(ast, parser)).toContain('*by style*');
    expect(convertAstToMarkdown(ast)).not.toContain('*by style*');
  });
});
