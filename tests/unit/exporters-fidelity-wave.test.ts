import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  convertAstToMarkdown,
  convertAstToHtml,
  createBlankDocxBuffer,
  executeInsert,
  executeNamedStyle,
  executeTable,
  DocumentAST,
} from '../../packages/engine/src';

function parseXml(xmlStr: string): Element {
  const doc = new DOMParser().parseFromString(xmlStr, 'text/xml');
  return doc.documentElement;
}

function createDocWithFootnote(fnText = 'Real Footnote text'): Buffer {
  const zip = new AdmZip();
  const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Section text before note</w:t></w:r>
      <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>`;
  const fnXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:type="separator" w:id="-1">
    <w:p><w:r><w:separator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:type="continuationSeparator" w:id="0">
    <w:p><w:r><w:continuationSeparator/></w:r></w:p>
  </w:footnote>
  <w:footnote w:id="1">
    <w:p>
      <w:r><w:rPr><w:rStyle w:val="FootnoteReference"/></w:rPr><w:footnoteRef/></w:r>
      <w:r><w:t>${fnText}</w:t></w:r>
    </w:p>
  </w:footnote>
</w:footnotes>`;
  zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  zip.addFile('word/footnotes.xml', Buffer.from(fnXml, 'utf8'));
  zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
  return zip.toBuffer();
}

function buildTestAst(nodes: any[], extra: any = {}): DocumentAST {
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

describe('Wave 1D Fidelity Tests: Footnotes, Redline HTML, GFM Alerts, Dividers, Monospace Code', () => {
  describe('P0-5: Real Footnote Bodies in Markdown Export', () => {
    it('emits real footnote body [^1]: Real Footnote text when parser is provided', async () => {
      const buf = createDocWithFootnote('Real Footnote text');
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const md = convertAstToMarkdown(ast, parser);

      expect(md).toBe('Section text before note[^1]\n\n[^1]: Real Footnote text\n');
      expect(md.split('[^1]').length - 1).toBe(2);
      expect(md).not.toContain('[Footnote 1: see document]');
    });

    it('falls back to placeholder when parser is omitted or footnotes part is missing', async () => {
      const buf = createDocWithFootnote('Real Footnote text');
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      // Omitting parser should preserve backward compatibility
      const md = convertAstToMarkdown(ast);

      expect(md).toBe('Section text before note[^1]\n\n[^1]: [Footnote 1: see document]\n');
    });

    it('extracts footnote body from ast.footnotes if provided on AST', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Sentence referencing note</w:t></w:r>
          <w:r><w:footnoteReference w:id="42"/></w:r>
        </w:p>
      `);
      const ast = buildTestAst(
        [
          {
            id: 'p_1',
            type: 'paragraph',
            xml: pXml,
            metadata: buildMeta('Sentence referencing note'),
            version: 1,
            checksum: 'chk_1',
          },
        ],
        { footnotes: { '42': 'Direct AST footnote body' } }
      );

      const md = convertAstToMarkdown(ast);
      expect(md).toBe('Sentence referencing note[^42]\n\n[^42]: Direct AST footnote body\n');
    });
  });

  describe('P1-15: HTML Redline Export with Tracked Changes', () => {
    it('exports HTML with both <ins class="tracked-ins"> and <del class="tracked-del"> containing deleted text', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Unchanged prefix. </w:t></w:r>
          <w:ins w:id="10" w:author="Reviewer" w:date="2026-09-23T10:00:00Z">
            <w:r><w:t>Newly inserted text.</w:t></w:r>
          </w:ins>
          <w:del w:id="11" w:author="Reviewer" w:date="2026-09-23T10:00:00Z">
            <w:r><w:delText>Previously deleted draft.</w:delText></w:r>
          </w:del>
          <w:r><w:t> Unchanged suffix.</w:t></w:r>
        </w:p>
      `);

      const ast = buildTestAst([
        {
          id: 'p_1',
          type: 'paragraph',
          xml: pXml,
          metadata: buildMeta('Unchanged prefix. Newly inserted text. Unchanged suffix.', {
            hasTrackedChanges: true,
          }),
          version: 1,
          checksum: 'chk_1',
        },
      ]);

      const html = convertAstToHtml(ast, 'Redline Test');
      // Outer redline tags stay exact; revision metadata rides on an inner span.
      const redlineDoc = new DOMParser().parseFromString(html, 'text/html');
      const insEls = redlineDoc.getElementsByTagName('ins');
      const delEls = redlineDoc.getElementsByTagName('del');
      expect(insEls.length).toBe(1);
      expect(delEls.length).toBe(1);
      expect((insEls[0] as Element).getAttribute('class')).toBe('tracked-ins');
      expect((delEls[0] as Element).getAttribute('class')).toBe('tracked-del');
      expect((insEls[0] as Element).textContent).toBe('Newly inserted text.');
      expect((delEls[0] as Element).textContent).toBe('Previously deleted draft.');
      const insSpan = (insEls[0] as Element).getElementsByTagName('span')[0] as unknown as Element;
      const delSpan = (delEls[0] as Element).getElementsByTagName('span')[0] as unknown as Element;
      expect(insSpan.getAttribute('data-author')).toBe('Reviewer');
      expect(insSpan.getAttribute('data-date')).toBe('2026-09-23T10:00:00Z');
      expect(insSpan.getAttribute('data-id')).toBe('10');
      expect(insSpan.getAttribute('title')).toBe('Inserted by Reviewer on 2026-09-23T10:00:00Z');
      expect(delSpan.getAttribute('data-author')).toBe('Reviewer');
      expect(delSpan.getAttribute('data-date')).toBe('2026-09-23T10:00:00Z');
      expect(delSpan.getAttribute('data-id')).toBe('11');
      expect(delSpan.getAttribute('title')).toBe('Deleted by Reviewer on 2026-09-23T10:00:00Z');

      // Verify Markdown accepted view retains insertion and omits deletion
      const md = convertAstToMarkdown(ast);
      expect(md).toContain('Newly inserted text.');
      expect(md).not.toContain('Previously deleted draft.');
    });
  });

  describe('P1-16: Markdown Fidelity — GFM Alerts, Dividers, Monospace Code', () => {
    it('renders callout paragraphs as GFM alerts (> [!NOTE], > [!WARNING])', () => {
      const noteXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr><w:pStyle w:val="Callout"/></w:pPr>
          <w:r><w:t>This is an informative callout note.</w:t></w:r>
        </w:p>
      `);
      const warnXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr><w:pStyle w:val="Warning"/></w:pPr>
          <w:r><w:t>Deprecated v1 authentication will be deactivated.</w:t></w:r>
        </w:p>
      `);
      const quoteXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr><w:pStyle w:val="Quote"/></w:pPr>
          <w:r><w:t>Simplicity is prerequisite for reliability.</w:t></w:r>
        </w:p>
      `);

      const ast = buildTestAst([
        {
          id: 'p_note',
          type: 'paragraph',
          xml: noteXml,
          metadata: buildMeta('This is an informative callout note.', { namedStyle: 'Callout' }),
          version: 1,
          checksum: 'chk_note',
        },
        {
          id: 'p_warn',
          type: 'paragraph',
          xml: warnXml,
          metadata: buildMeta('Deprecated v1 authentication will be deactivated.', { namedStyle: 'Warning' }),
          version: 1,
          checksum: 'chk_warn',
        },
        {
          id: 'p_quote',
          type: 'paragraph',
          xml: quoteXml,
          metadata: buildMeta('Simplicity is prerequisite for reliability.', { namedStyle: 'Quote' }),
          version: 1,
          checksum: 'chk_quote',
        },
      ]);

      const md = convertAstToMarkdown(ast);
      expect(md).toBe(
        '> [!NOTE]\n> This is an informative callout note.\n\n> [!WARNING]\n> Deprecated v1 authentication will be deactivated.\n\n> Simplicity is prerequisite for reliability.\n'
      );
      expect(md).not.toContain('> [!NOTE]\n> Simplicity');
    });

    it('renders callout table boxes with metadata as GFM alerts', () => {
      const ast = buildTestAst([
        {
          id: 'tbl_callout',
          type: 'table',
          xml: parseXml('<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tr><w:tc><w:p><w:r><w:t>Amber alert content</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
          metadata: buildMeta('Amber alert content', {
            callout: { theme: 'amber', title: 'Security Warning' },
          }),
          version: 1,
          checksum: 'chk_tbl_c',
        },
      ]);

      const md = convertAstToMarkdown(ast);
      expect(md).toBe('> [!WARNING]\n> **Security Warning**\n> Amber alert content\n');
      expect(md).not.toContain('| Amber alert content |');
    });

    it('renders horizontal divider as --- from w:pBdr bottom or divider node type', () => {
      const dividerXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:pBdr>
              <w:bottom w:val="single" w:sz="12" w:space="1" w:color="CBD5E1"/>
            </w:pBdr>
          </w:pPr>
        </w:p>
      `);

      const ast = buildTestAst([
        {
          id: 'p_before',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Content before divider</w:t></w:r></w:p>'),
          metadata: buildMeta('Content before divider'),
          version: 1,
          checksum: 'chk_before',
        },
        {
          id: 'div_1',
          type: 'paragraph',
          xml: dividerXml,
          metadata: buildMeta(''),
          version: 1,
          checksum: 'chk_div',
        },
        {
          id: 'div_2',
          type: 'divider',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'),
          metadata: buildMeta(''),
          version: 1,
          checksum: 'chk_div2',
        },
        {
          id: 'p_after',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Content after divider</w:t></w:r></w:p>'),
          metadata: buildMeta('Content after divider'),
          version: 1,
          checksum: 'chk_after',
        },
      ]);

      const md = convertAstToMarkdown(ast);
      expect(md).toContain('Content before divider\n');
      expect(md).toContain('---\n');
      expect(md).toContain('Content after divider\n');
      const dividerMatches = md.match(/^---$/gm);
      expect(dividerMatches?.length).toBe(2);
    });

    it('wraps monospace runs (Consolas, Courier New, Code style) in inline code backticks', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Call </w:t></w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr>
            <w:t>GET /v2/telemetry</w:t>
          </w:r>
          <w:r><w:t> with header </w:t></w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr>
            <w:t>Authorization</w:t>
          </w:r>
          <w:r><w:t> and run </w:t></w:r>
          <w:r>
            <w:rPr><w:rStyle w:val="Code"/></w:rPr>
            <w:t>npm test</w:t>
          </w:r>
          <w:r><w:t> to verify.</w:t></w:r>
        </w:p>
      `);

      const ast = buildTestAst([
        {
          id: 'p_code',
          type: 'paragraph',
          xml: pXml,
          metadata: buildMeta('Call GET /v2/telemetry with header Authorization and run npm test to verify.'),
          version: 1,
          checksum: 'chk_code',
        },
      ]);

      const md = convertAstToMarkdown(ast);
      expect(md).toBe('Call `GET /v2/telemetry` with header `Authorization` and run `npm test` to verify.\n');
    });
  });

  describe('Hostile-review fixes: note collision, del metadata, callout type, divider text, code escapes', () => {
    function createDocWithFootnoteAndEndnote(): Buffer {
      const zip = new AdmZip();
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Body with footnote</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t> and endnote</w:t></w:r>
      <w:r><w:endnoteReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>`;
      const fnXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>Footnote body alpha</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`;
      const enXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:p><w:r><w:t>Endnote body beta</w:t></w:r></w:p>
  </w:endnote>
</w:endnotes>`;
      zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
      zip.addFile('word/footnotes.xml', Buffer.from(fnXml, 'utf8'));
      zip.addFile('word/endnotes.xml', Buffer.from(enXml, 'utf8'));
      zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
      return zip.toBuffer();
    }

    it('keeps footnote id 1 and endnote id 1 as distinct defs with correct bodies', async () => {
      const buf = createDocWithFootnoteAndEndnote();
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const md = convertAstToMarkdown(ast, parser);

      expect(md).toContain('[^1]: Footnote body alpha');
      expect(md).toContain('[^en-1]: Endnote body beta');
      expect(md).toContain('[^1]');
      expect(md).toContain('[^en-1]');
    });

    it('emits both note entries in HTML without id collision', async () => {
      const buf = createDocWithFootnoteAndEndnote();
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const html = convertAstToHtml(ast, 'Notes');

      expect(html).toContain('id="fn-1"');
      expect(html).toContain('id="en-1"');
      expect(html).toContain('[1]');
      expect(html).toContain('[en-1]');
    });

    it('skips orphan note defs that have no in-text ref', () => {
      const ast = buildTestAst(
        [
          {
            id: 'p_1',
            type: 'paragraph',
            xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Plain text.</w:t></w:r></w:p>'),
            metadata: buildMeta('Plain text.'),
            version: 1,
            checksum: 'chk_1',
          },
        ],
        { footnotes: { '9': 'Orphan body never referenced' } }
      );
      const md = convertAstToMarkdown(ast);
      expect(md).toContain('Plain text.');
      expect(md).not.toContain('Orphan body never referenced');
      expect(md).not.toContain('[^9]');
    });

    it('carries w:author/date/id on del and ins metadata spans', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:ins w:id="7" w:author="Alice" w:date="2026-05-01T00:00:00Z">
            <w:r><w:t>Added line.</w:t></w:r>
          </w:ins>
          <w:del w:id="8" w:author="Bob" w:date="2026-05-02T00:00:00Z">
            <w:r><w:delText>Removed line.</w:delText></w:r>
          </w:del>
        </w:p>
      `);
      const ast = buildTestAst([
        {
          id: 'p_1',
          type: 'paragraph',
          xml: pXml,
          metadata: buildMeta('Added line.', { hasTrackedChanges: true }),
          version: 1,
          checksum: 'chk_1',
        },
      ]);
      const html = convertAstToHtml(ast, 'Meta');
      const metaDoc = new DOMParser().parseFromString(html, 'text/html');
      const metaIns = metaDoc.getElementsByTagName('ins');
      const metaDel = metaDoc.getElementsByTagName('del');
      expect(metaIns.length).toBe(1);
      expect(metaDel.length).toBe(1);
      expect((metaIns[0] as Element).getAttribute('class')).toBe('tracked-ins');
      expect((metaDel[0] as Element).getAttribute('class')).toBe('tracked-del');
      expect((metaIns[0] as Element).textContent).toBe('Added line.');
      expect((metaDel[0] as Element).textContent).toBe('Removed line.');
      const metaInsSpan = (metaIns[0] as Element).getElementsByTagName('span')[0] as unknown as Element;
      const metaDelSpan = (metaDel[0] as Element).getElementsByTagName('span')[0] as unknown as Element;
      expect(metaInsSpan.getAttribute('data-author')).toBe('Alice');
      expect(metaInsSpan.getAttribute('data-date')).toBe('2026-05-01T00:00:00Z');
      expect(metaInsSpan.getAttribute('data-id')).toBe('7');
      expect(metaInsSpan.getAttribute('title')).toBe('Inserted by Alice on 2026-05-01T00:00:00Z');
      expect(metaDelSpan.getAttribute('data-author')).toBe('Bob');
      expect(metaDelSpan.getAttribute('data-date')).toBe('2026-05-02T00:00:00Z');
      expect(metaDelSpan.getAttribute('data-id')).toBe('8');
      expect(metaDelSpan.getAttribute('title')).toBe('Deleted by Bob on 2026-05-02T00:00:00Z');
    });

    it('honors an explicit table callout type over theme mapping', () => {
      const ast = buildTestAst([
        {
          id: 'tbl_callout',
          type: 'table',
          xml: parseXml('<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tr><w:tc><w:p><w:r><w:t>Helpful hint body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
          metadata: buildMeta('Helpful hint body', {
            callout: { theme: 'amber', title: 'Hint', type: 'TIP' },
          }),
          version: 1,
          checksum: 'chk_tbl_tip',
        },
      ]);
      const md = convertAstToMarkdown(ast);
      expect(md).toContain('> [!TIP]');
      expect(md).not.toContain('> [!WARNING]');
      expect(md).toContain('> Helpful hint body');
    });

    it('never drops paragraph text on the divider branch', () => {
      const dividerWithText = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:pPr>
            <w:pBdr>
              <w:bottom w:val="single" w:sz="12" w:space="1" w:color="CBD5E1"/>
            </w:pBdr>
          </w:pPr>
          <w:r><w:t>Keep this sentence</w:t></w:r>
        </w:p>
      `);
      const ast = buildTestAst([
        {
          id: 'div_text',
          type: 'divider',
          xml: dividerWithText,
          metadata: buildMeta('Keep this sentence'),
          version: 1,
          checksum: 'chk_div_text',
        },
      ]);
      const md = convertAstToMarkdown(ast);
      expect(md).toBe('Keep this sentence\n\n---\n');
    });

    it('escapes backticks and pipes inside monospace code spans', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Run </w:t></w:r>
          <w:r>
            <w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr>
            <w:t>cmd \`tick\` a|b</w:t>
          </w:r>
          <w:r><w:t> now.</w:t></w:r>
        </w:p>
      `);
      const ast = buildTestAst([
        {
          id: 'p_code',
          type: 'paragraph',
          xml: pXml,
          metadata: buildMeta('Run cmd tick now.'),
          version: 1,
          checksum: 'chk_code_esc',
        },
      ]);
      const md = convertAstToMarkdown(ast);
      // Pipe escaped so GFM tables cannot split; backtick fence lengthened.
      expect(md).toContain('\\|');
      expect(md).toContain('``');
      expect(md).not.toContain('`cmd `tick` a|b`');
    });

    it('emits no empty del/ins wrappers for empty tracked changes', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:del w:id="1" w:author="A" w:date="2026-05-01T00:00:00Z"><w:r><w:t></w:t></w:r></w:del>
          <w:ins w:id="2" w:author="A" w:date="2026-05-01T00:00:00Z"><w:r><w:t></w:t></w:r></w:ins>
          <w:r><w:t>Visible</w:t></w:r>
        </w:p>
      `);
      const ast = buildTestAst([
        {
          id: 'p_empty',
          type: 'paragraph',
          xml: pXml,
          metadata: buildMeta('Visible', { hasTrackedChanges: true }),
          version: 1,
          checksum: 'chk_empty',
        },
      ]);
      const html = convertAstToHtml(ast, 'Empty');
      expect(html).toContain('Visible');
      expect(html).not.toContain('<del');
      expect(html).not.toContain('<ins');
      expect(html).not.toContain('<span></span>');
    });

    it('keeps normal del/ins output free of empty spans', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Before </w:t></w:r>
          <w:ins w:id="3" w:author="Ann" w:date="2026-05-03T00:00:00Z"><w:r><w:t>added</w:t></w:r></w:ins>
          <w:del w:id="4" w:author="Bob" w:date="2026-05-04T00:00:00Z"><w:r><w:delText>removed</w:delText></w:r></w:del>
          <w:r><w:t> after.</w:t></w:r>
        </w:p>
      `);
      const ast = buildTestAst([
        {
          id: 'p_normal',
          type: 'paragraph',
          xml: pXml,
          metadata: buildMeta('Before added after.', { hasTrackedChanges: true }),
          version: 1,
          checksum: 'chk_normal',
        },
      ]);
      const html = convertAstToHtml(ast, 'Normal');
      expect(html).toContain('<ins class="tracked-ins">');
      expect(html).toContain('<del class="tracked-del">');
      expect(html).toContain('added');
      expect(html).toContain('removed');
      expect(html).not.toContain('<span></span>');
    });

    it('propagates list-item footnotes into the endnotes section', () => {
      const paraXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Para text</w:t></w:r>
          <w:r><w:footnoteReference w:id="1"/></w:r>
        </w:p>
      `);
      const listXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>List item text</w:t></w:r>
          <w:r><w:footnoteReference w:id="2"/></w:r>
        </w:p>
      `);
      const ast = buildTestAst([
        {
          id: 'p_fn',
          type: 'paragraph',
          xml: paraXml,
          metadata: buildMeta('Para text'),
          version: 1,
          checksum: 'chk_p_fn',
        },
        {
          id: 'li_fn',
          type: 'list',
          xml: listXml,
          metadata: buildMeta('List item text'),
          version: 1,
          checksum: 'chk_li_fn',
        },
      ]);
      const html = convertAstToHtml(ast, 'ListFn');
      expect(html).toContain('<sup class="footnote-ref">[1]</sup>');
      expect(html).toContain('<sup class="footnote-ref">[2]</sup>');
      expect(html).toContain('id="fn-1"');
      expect(html).toContain('id="fn-2"');
      const sectionStart = html.indexOf('<section class="footnotes">');
      expect(sectionStart).toBeGreaterThan(-1);
      const section = html.slice(sectionStart);
      expect(section).toContain('id="fn-1"');
      expect(section).toContain('id="fn-2"');
    });

    it('trims whitespace before callout type allowlist lookup', () => {
      const ast = buildTestAst([
        {
          id: 'tbl_callout',
          type: 'table',
          xml: parseXml('<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:tr><w:tc><w:p><w:r><w:t>Helpful hint body</w:t></w:r></w:p></w:tc></w:tr></w:tbl>'),
          metadata: buildMeta('Helpful hint body', {
            callout: { theme: 'amber', title: 'Hint', type: ' TIP ' },
          }),
          version: 1,
          checksum: 'chk_tbl_tip_ws',
        },
      ]);
      const md = convertAstToMarkdown(ast);
      expect(md).toContain('> [!TIP]');
      expect(md).not.toContain('> [!WARNING]');
      expect(md).toContain('> Helpful hint body');
    });
  });

  describe('R-39: chart branch precedes image branch in HTML export', () => {
    it('renders chart node carrying hasImage:true as svg, not image placeholder', () => {
      const ast = buildTestAst([
        {
          id: 'chart_1',
          type: 'chart',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"><a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'),
          metadata: buildMeta('', {
            hasImage: true,
            title: 'Quarterly Revenue',
            categories: ['Q1', 'Q2'],
            series: [{ name: 'Revenue', values: [10, 20], color: '2563EB' }],
            chartType: 'bar',
          }),
          version: 1,
          checksum: 'chk_chart_img',
        },
      ]);
      const html = convertAstToHtml(ast, 'ChartImage');
      expect(html).toContain('<svg');
      expect(html).toContain('Quarterly Revenue');
      expect(html).not.toContain('[Image:');
    });

    it('renders chart3d node carrying hasImage:true as svg, not image placeholder', () => {
      const ast = buildTestAst([
        {
          id: 'chart3d_1',
          type: 'chart3d',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:drawing /></w:r></w:p>'),
          metadata: buildMeta('', {
            hasImage: true,
            title: 'Volume 3D',
            categories: ['A', 'B'],
            series: [{ name: 'S1', values: [5, 15], color: '10B981' }],
            chartType: 'bar3D',
          }),
          version: 1,
          checksum: 'chk_chart3d_img',
        },
      ]);
      const html = convertAstToHtml(ast, 'Chart3DImage');
      expect(html).toContain('<svg');
      expect(html).toContain('Volume 3D');
      expect(html).not.toContain('[Image:');
    });

    it('renders true image node with embedded bytes as img', () => {
      const imgXml = parseXml(
        '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
        '<w:r><w:drawing><wp:inline><wp:docPr id="1" name="Picture 1" descr="Sample pic"/>' +
        '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic>' +
        '</wp:inline></w:drawing></w:r></w:p>'
      );
      const pngBytes = Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64');
      const ast = buildTestAst(
        [
          {
            id: 'img_1',
            type: 'image',
            xml: imgXml,
            metadata: buildMeta('', { hasImage: true }),
            version: 1,
            checksum: 'chk_img_true',
          },
        ],
        {
          relationships: { rId5: { target: 'media/image1.png' } },
          mediaFiles: { 'image1.png': pngBytes, 'media/image1.png': pngBytes, 'word/media/image1.png': pngBytes },
        }
      );
      const html = convertAstToHtml(ast, 'TrueImage');
      expect(html).toContain('<img');
      expect(html).toContain('data:image/png;base64,');
      expect(html).not.toContain('[Image:');
    });

    it('keeps paragraph caption emission for image-branch nodes', () => {
      const ast = buildTestAst([
        {
          id: 'fig_1',
          type: 'drawing',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:drawing /></w:r></w:p>'),
          metadata: buildMeta('Figure 1: Sales overview', { hasImage: true }),
          version: 1,
          checksum: 'chk_caption',
        },
      ]);
      const html = convertAstToHtml(ast, 'Caption');
      expect(html).toContain('doc-image');
      expect(html).toContain('<p class="paragraph">Figure 1: Sales overview</p>');
    });
  });

  describe('R-3: HTML renders real footnote/endnote bodies', () => {
    function createFootnoteEndnoteDoc(): Buffer {
      const zip = new AdmZip();
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Body with footnote</w:t></w:r>
      <w:r><w:footnoteReference w:id="1"/></w:r>
      <w:r><w:t> and endnote</w:t></w:r>
      <w:r><w:endnoteReference w:id="1"/></w:r>
    </w:p>
  </w:body>
</w:document>`;
      const fnXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1">
    <w:p><w:r><w:t>Footnote body alpha</w:t></w:r></w:p>
  </w:footnote>
</w:footnotes>`;
      const enXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1">
    <w:p><w:r><w:t>Endnote body beta</w:t></w:r></w:p>
  </w:endnote>
</w:endnotes>`;
      zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
      zip.addFile('word/footnotes.xml', Buffer.from(fnXml, 'utf8'));
      zip.addFile('word/endnotes.xml', Buffer.from(enXml, 'utf8'));
      zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
      return zip.toBuffer();
    }

    it('emits the real footnote body in <li> when the parser is provided', async () => {
      const buf = createDocWithFootnote('Real Footnote text');
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const html = convertAstToHtml(ast, 'FootnoteDoc', parser);

      expect(html).toContain('<li id="fn-1">Real Footnote text</li>');
      expect(html).not.toContain('[Footnote 1: see document]');
    });

    it('keeps footnote id 1 and endnote id 1 distinct with correct bodies', async () => {
      const buf = createFootnoteEndnoteDoc();
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const html = convertAstToHtml(ast, 'NotesDoc', parser);

      expect(html).toContain('<li id="fn-1">Footnote body alpha</li>');
      expect(html).toContain('<li id="en-1">Endnote body beta</li>');
      expect(html).not.toContain('[Footnote 1: see document]');
      expect(html).not.toContain('[Footnote en-1: see document]');
    });

    it('resolves bodies from ast.footnotes/ast.endnotes without a parser', () => {
      const pXml = parseXml(`
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:r><w:t>Sentence with both notes</w:t></w:r>
          <w:r><w:footnoteReference w:id="1"/></w:r>
          <w:r><w:endnoteReference w:id="1"/></w:r>
        </w:p>
      `);
      const ast = buildTestAst(
        [
          {
            id: 'p_1',
            type: 'paragraph',
            xml: pXml,
            metadata: buildMeta('Sentence with both notes'),
            version: 1,
            checksum: 'chk_notes_ast',
          },
        ],
        { footnotes: { '1': 'AST footnote one' }, endnotes: { '1': 'AST endnote one' } }
      );
      const html = convertAstToHtml(ast, 'AstNotes');

      expect(html).toContain('<li id="fn-1">AST footnote one</li>');
      expect(html).toContain('<li id="en-1">AST endnote one</li>');
    });

    it('skips orphan defs and falls back to placeholder only when the body is unreachable', async () => {
      const orphanAst = buildTestAst(
        [
          {
            id: 'p_1',
            type: 'paragraph',
            xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Plain text.</w:t></w:r></w:p>'),
            metadata: buildMeta('Plain text.'),
            version: 1,
            checksum: 'chk_orphan',
          },
        ],
        { footnotes: { '9': 'Orphan body never referenced' } }
      );
      const orphanHtml = convertAstToHtml(orphanAst, 'Orphan');
      expect(orphanHtml).toContain('Plain text.');
      expect(orphanHtml).not.toContain('Orphan body never referenced');
      expect(orphanHtml).not.toContain('id="fn-9"');
      expect(orphanHtml).not.toContain('<section class="footnotes">');

      const buf = createDocWithFootnote('Real Footnote text');
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const fallbackHtml = convertAstToHtml(ast, 'NoParser');
      expect(fallbackHtml).toContain('<li id="fn-1">[Footnote 1: see document]</li>');
    });

    it('accepts options plus parser without losing bodies or print styles', async () => {
      const buf = createDocWithFootnote('Real Footnote text');
      const parser = new DocxParser(buf);
      const ast = await parser.getAST();
      const html = convertAstToHtml(ast, 'PrintNotes', { printMode: true }, parser);

      expect(html).toContain('<li id="fn-1">Real Footnote text</li>');
      expect(html).toContain('@page');
    });
  });

  describe('R-42: table setCaption single-render proof', () => {
    it('renders real setCaption output exactly once through HTML and Markdown', async () => {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST() as any;
      const anchor = ast.sections[0].nodes[0].id;
      await executeInsert({
        action: 'insert', type: 'table', anchor, position: 'after',
        headers: ['A', 'B'], rows: [['1', '2']],
      } as any, ast, parser);
      const tblNode = ast.sections[0].nodes.find((n: any) => n.type === 'table');
      expect(tblNode).toBeDefined();
      await executeTable({
        action: 'table', target: tblNode.id as string,
        params: { op: 'setCaption', captionText: 'Table 7: Single Render Proof', captionPosition: 'below' },
      }, ast);
      const html = convertAstToHtml(ast, 'SingleRender');
      const md = convertAstToMarkdown(ast, parser);
      const text = 'Table 7: Single Render Proof';
      expect(html.split(text).length - 1).toBe(1);
      expect(md.split(text).length - 1).toBe(1);
      expect(html.indexOf(text)).toBeGreaterThan(html.indexOf('<table'));
    }, 120000);
  });

  describe('hyperlink targets in Markdown export', () => {
    const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const R_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

    function createDocWithHyperlink(withRel: boolean): Buffer {
      const zip = new AdmZip();
      const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}">
  <w:body>
    <w:p>
      <w:r><w:t>See </w:t></w:r>
      <w:hyperlink r:id="rId9"><w:r><w:t>Example Site</w:t></w:r></w:hyperlink>
      <w:r><w:t> for details.</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
      const relEntry = withRel
        ? `<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/docs" TargetMode="External"/>`
        : '';
      const relsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relEntry}</Relationships>`;
      zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
      zip.addFile('word/_rels/document.xml.rels', Buffer.from(relsXml, 'utf8'));
      zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
      return zip.toBuffer();
    }

    it('exports real URLs for hyperlinks resolved through document rels', async () => {
      const parser = new DocxParser(createDocWithHyperlink(true));
      const ast = await parser.getAST();
      const md = convertAstToMarkdown(ast, parser);
      expect(md).toBe('See [Example Site](https://example.com/docs) for details.\n');
    });

    it('degrades to plain text when the rel target is missing', async () => {
      const parser = new DocxParser(createDocWithHyperlink(false));
      const ast = await parser.getAST();
      let md = '';
      expect(() => { md = convertAstToMarkdown(ast, parser); }).not.toThrow();
      expect(md).toContain('Example Site');
      expect(md).not.toContain('[Example Site](');
    });

    it('resolves synthetic hyperlink nodes against the relationships map', () => {
      const pXml = parseXml(
        `<w:p xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:r><w:t>Read </w:t></w:r>` +
        `<w:hyperlink r:id="rId5"><w:r><w:t>Manual</w:t></w:r></w:hyperlink></w:p>`
      );
      const ast = buildTestAst(
        [{ id: 'p1', type: 'paragraph', xml: pXml, metadata: buildMeta('Read Manual') }],
        { relationships: { rId5: { type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink', target: 'https://example.com/manual', targetMode: 'External' } } }
      );
      expect(convertAstToMarkdown(ast)).toContain('[Manual](https://example.com/manual)');
    });
  });

  describe('negative chart values and divider rules in HTML export', () => {
    it('renders negative bars with non-negative heights below the zero baseline', () => {
      const ast = buildTestAst([
        {
          id: 'chart_neg',
          type: 'chart',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>chart</w:t></w:r></w:p>'),
          metadata: buildMeta('', {
            title: 'Profit And Loss',
            categories: ['A', 'B', 'C'],
            series: [{ name: 'Net', values: [30, -20, 10], color: '2563EB' }],
            chartType: 'bar',
          }),
          version: 1,
          checksum: 'chk_chart_neg',
        },
      ]);
      const html = convertAstToHtml(ast, 'NegativeChart');
      expect(html).toContain('<svg');
      expect(html).toContain('Net: -20');
      expect(html).not.toContain('height="-');
      const rects = Array.from(html.matchAll(/<rect x="([\d.eE+-]+)" y="([\d.eE+-]+)" width="([\d.eE+-]+)" height="([\d.eE+-]+)"[^>]*>\s*<title>([^<]*)<\/title>/g));
      expect(rects.length).toBe(3);
      for (const r of rects) {
        expect(Number(r[4])).toBeGreaterThanOrEqual(0);
      }
      const neg = rects.find(r => r[5].includes('-20'));
      const pos = rects.find(r => r[5].includes('30'));
      expect(neg).toBeDefined();
      expect(pos).toBeDefined();
      expect(Number(neg![2])).toBeGreaterThan(Number(pos![2]));
    });

    it('keeps negative line points inside the plot area', () => {
      const ast = buildTestAst([
        {
          id: 'chart_neg_line',
          type: 'chart',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>chart</w:t></w:r></w:p>'),
          metadata: buildMeta('', {
            title: 'Trend With Dip',
            categories: ['A', 'B', 'C'],
            series: [{ name: 'Net', values: [10, -15, 5], color: '10B981' }],
            chartType: 'line',
          }),
          version: 1,
          checksum: 'chk_chart_neg_line',
        },
      ]);
      const html = convertAstToHtml(ast, 'NegativeLine');
      expect(html).toContain('<polyline');
      const cys = Array.from(html.matchAll(/<circle cx="[\d.eE+-]+" cy="([\d.eE+-]+)"/g)).map(m => Number(m[1]));
      expect(cys.length).toBe(3);
      for (const cy of cys) {
        expect(cy).toBeGreaterThanOrEqual(40);
        expect(cy).toBeLessThanOrEqual(180);
      }
    });

    it('renders divider nodes as a horizontal rule without dropping nearby text', () => {
      const ast = buildTestAst([
        {
          id: 'p_before',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Before text</w:t></w:r></w:p>'),
          metadata: buildMeta('Before text'),
          version: 1,
          checksum: 'chk_div_before',
        },
        {
          id: 'div_1',
          type: 'divider',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'),
          metadata: buildMeta(''),
          version: 1,
          checksum: 'chk_div_rule',
        },
        {
          id: 'p_after',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>After text</w:t></w:r></w:p>'),
          metadata: buildMeta('After text'),
          version: 1,
          checksum: 'chk_div_after',
        },
      ]);
      const html = convertAstToHtml(ast, 'DividerDoc');
      expect(html).toContain('<hr class="doc-divider"/>');
      expect(html).toContain('Before text');
      expect(html).toContain('After text');
      expect(html).not.toContain('Unsupported node type: divider');
    });

    it('keeps attached text above the rule for divider nodes carrying text', () => {
      const ast = buildTestAst([
        {
          id: 'div_text',
          type: 'horizontalRule',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Section break note</w:t></w:r></w:p>'),
          metadata: buildMeta('Section break note'),
          version: 1,
          checksum: 'chk_div_hr_text',
        },
      ]);
      const html = convertAstToHtml(ast, 'DividerTextDoc');
      expect(html).toContain('<hr class="doc-divider"/>');
      expect(html).toContain('Section break note');
      expect(html.indexOf('Section break note')).toBeLessThan(html.indexOf('<hr class="doc-divider"/>'));
    });

    it('renders paragraph with triple dash as a horizontal rule divider in HTML', () => {
      const ast = buildTestAst([
        {
          id: 'p_before',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Heading intro</w:t></w:r></w:p>'),
          metadata: buildMeta('Heading intro'),
          version: 1,
          checksum: 'chk_p_before',
        },
        {
          id: 'p_divider',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>---</w:t></w:r></w:p>'),
          metadata: buildMeta('---'),
          version: 1,
          checksum: 'chk_p_divider',
        },
        {
          id: 'p_after',
          type: 'paragraph',
          xml: parseXml('<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Following section</w:t></w:r></w:p>'),
          metadata: buildMeta('Following section'),
          version: 1,
          checksum: 'chk_p_after',
        },
      ]);
      const html = convertAstToHtml(ast, 'ParagraphDividerDoc');
      expect(html).toContain('<hr class="doc-divider"/>');
      expect(html).toContain('Heading intro');
      expect(html).toContain('Following section');
      expect(html).not.toContain('<p class="paragraph">---</p>');
    });
  });
});

describe('Multiline code blocks in Markdown and HTML export', () => {
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function codePara(id: string, inner: string, text: string, extraMeta: any = {}): any {
    return {
      id,
      type: 'paragraph',
      xml: parseXml(`<w:p xmlns:w="${W}">${inner}</w:p>`),
      metadata: buildMeta(text, extraMeta),
      version: 1,
      checksum: `chk_${id}`,
    };
  }

  it('exports all-monospace runs joined by breaks as a fenced Markdown block', () => {
    const ast = buildTestAst([
      codePara(
        'p_fence',
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>const a = 1;</w:t></w:r>' +
        '<w:r><w:br/></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>return a;</w:t></w:r>',
        'const a = 1; return a;'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toBe('```\nconst a = 1;\nreturn a;\n```\n');
    expect(md).not.toContain('`const a = 1;`');
  });

  it('exports a code-named-style paragraph spanning breaks as a fenced block', () => {
    const ast = buildTestAst([
      codePara(
        'p_named',
        '<w:pPr><w:pStyle w:val="Code"/></w:pPr>' +
        '<w:r><w:t>first();</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>second();</w:t></w:r>',
        'first(); second();',
        { namedStyle: 'Code' }
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('```\nfirst();\nsecond();\n```');
  });

  it('leaves single-line monospace runs as inline spans without fencing', () => {
    const ast = buildTestAst([
      codePara(
        'p_single',
        '<w:r><w:t>Run </w:t></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>npm test</w:t></w:r>',
        'Run npm test'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('Run `npm test`');
    expect(md).not.toContain('```');
  });

  it('exports a single-line code block with its language as a fence', () => {
    const ast = buildTestAst([
      codePara(
        'p_single_lang',
        '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' +
        '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="D1D5DB"/></w:pBdr></w:pPr>' +
        '<w:bookmarkStart w:id="7" w:name="_AstraCode_js"/>' +
        '<w:r><w:rPr><w:rStyle w:val="Code"/><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr>' +
        '<w:t>console.log("hello");</w:t></w:r>' +
        '<w:bookmarkEnd w:id="7"/>',
        'console.log("hello");'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toBe('```js\nconsole.log("hello");\n```\n');
  });

  it('exports a single-line code block without a language as a bare fence', () => {
    const ast = buildTestAst([
      codePara(
        'p_single_bare',
        '<w:pPr><w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>' +
        '<w:pBdr><w:bottom w:val="single" w:sz="6" w:space="4" w:color="D1D5DB"/></w:pBdr></w:pPr>' +
        '<w:r><w:rPr><w:rStyle w:val="Code"/><w:rFonts w:ascii="Consolas" w:hAnsi="Consolas"/></w:rPr>' +
        '<w:t>just code</w:t></w:r>',
        'just code'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toBe('```\njust code\n```\n');
  });

  it('leaves plain multiline paragraphs without fences', () => {
    const ast = buildTestAst([
      codePara(
        'p_plain',
        '<w:r><w:t>line one</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>line two</w:t></w:r>',
        'line one line two'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('line one\nline two');
    expect(md).not.toContain('```');
  });

  it('leaves mixed monospace and plain multiline runs without fencing', () => {
    const ast = buildTestAst([
      codePara(
        'p_mixed',
        '<w:r><w:t>Run </w:t></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>tool</w:t></w:r>' +
        '<w:r><w:br/></w:r><w:r><w:t>then read the output</w:t></w:r>',
        'Run tool then read the output'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('Run `tool`');
    expect(md).toContain('then read the output');
    expect(md).not.toContain('```');
  });

  it('lengthens the fence when fenced content holds triple backticks', () => {
    const ast = buildTestAst([
      codePara(
        'p_ticks',
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>open ``` here</w:t></w:r>' +
        '<w:r><w:br/></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>close</w:t></w:r>',
        'open ticks close'
      ),
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('````\nopen ``` here\nclose\n````');
  });

  it('exports all-monospace multiline runs as pre code in HTML', () => {
    const ast = buildTestAst([
      codePara(
        'p_pre',
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>if (a &lt; b)</w:t></w:r>' +
        '<w:r><w:br/></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>return a;</w:t></w:r>',
        'if (a < b) return a;'
      ),
    ]);
    const html = convertAstToHtml(ast, 'CodeDoc');
    expect(html).toContain('<pre><code>if (a &lt; b)\nreturn a;</code></pre>');
    expect(html).not.toContain('<p class="paragraph">if');
  });

  it('exports a code-named-style multiline paragraph as pre code in HTML', () => {
    const ast = buildTestAst([
      codePara(
        'p_prenamed',
        '<w:pPr><w:pStyle w:val="Code"/></w:pPr>' +
        '<w:r><w:t>alpha</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>beta</w:t></w:r>',
        'alpha beta',
        { namedStyle: 'Code' }
      ),
    ]);
    const html = convertAstToHtml(ast, 'CodeDoc');
    expect(html).toContain('<pre><code>alpha\nbeta</code></pre>');
  });

  it('leaves single-line monospace and plain multiline paragraphs as HTML paragraphs', () => {
    const ast = buildTestAst([
      codePara(
        'p_hsingle',
        '<w:r><w:t>Call </w:t></w:r>' +
        '<w:r><w:rPr><w:rFonts w:ascii="Consolas"/></w:rPr><w:t>npm test</w:t></w:r>',
        'Call npm test'
      ),
      codePara(
        'p_hplain',
        '<w:r><w:t>line one</w:t></w:r><w:r><w:br/></w:r><w:r><w:t>line two</w:t></w:r>',
        'line one line two'
      ),
    ]);
    const html = convertAstToHtml(ast, 'CodeDoc');
    expect(html).not.toContain('<pre>');
    expect(html).toContain('npm test</span></p>');
    expect(html).toContain('<p class="paragraph">line one<br/>line two</p>');
  });

  it('round-trips an inserted multiline paragraph styled as code through save and reload', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST() as any;
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'const a = 1;\nconst b = 2;\nreturn a + b;', id: 'code-para',
    } as any, ast, parser);
    executeNamedStyle({
      action: 'namedStyle', operation: 'apply', styleName: 'Code', target: 'code-para',
    } as any, ast, parser);

    const mdBefore = convertAstToMarkdown(ast, parser);
    const htmlBefore = convertAstToHtml(ast, 'CodeDoc', parser);
    expect(mdBefore).toContain('```\nconst a = 1;\nconst b = 2;\nreturn a + b;\n```');
    expect(htmlBefore).toContain('<pre><code>const a = 1;\nconst b = 2;\nreturn a + b;</code></pre>');

    const buf = parser.rezip(ast);
    const reloaded = new DocxParser(buf);
    const astAfter = await reloaded.getAST();
    const mdAfter = convertAstToMarkdown(astAfter, reloaded);
    const htmlAfter = convertAstToHtml(astAfter, 'CodeDoc', reloaded);
    expect(mdAfter).toContain('```\nconst a = 1;\nconst b = 2;\nreturn a + b;\n```');
    expect(htmlAfter).toContain('<pre><code>const a = 1;\nconst b = 2;\nreturn a + b;</code></pre>');
    const nodeAfter = astAfter.sections[0].nodes.find((n: any) =>
      ((n.metadata?.text || '').includes('const a = 1;')));
    expect(nodeAfter).toBeDefined();
    expect((nodeAfter!.xml as Element).getElementsByTagName('w:br').length).toBe(2);
  }, 120000);
});

describe('numbered paragraphs export as list items', () => {
  const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

  function numberedParaXml(text: string, numId: string): Element {
    return parseXml(
      `<w:p xmlns:w="${W_NS}"><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="${numId}"/></w:numPr></w:pPr>` +
      `<w:r><w:t>${text}</w:t></w:r></w:p>`
    );
  }

  function numberingMaps(): any {
    const doc = new DOMParser().parseFromString(
      `<w:numbering xmlns:w="${W_NS}">` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
      `<w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
      `<w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/>` +
      `<w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
      `<w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>` +
      `</w:numbering>`,
      'text/xml'
    );
    const abstractNum: Record<string, Element> = {};
    const concreteNum: Record<string, Element> = {};
    const absEls = doc.getElementsByTagName('w:abstractNum');
    for (let i = 0; i < absEls.length; i++) {
      const key = (absEls[i] as Element).getAttribute('w:abstractNumId');
      if (key !== null) abstractNum[key] = absEls[i] as unknown as Element;
    }
    const numEls = doc.getElementsByTagName('w:num');
    for (let i = 0; i < numEls.length; i++) {
      const key = (numEls[i] as Element).getAttribute('w:numId');
      if (key !== null) concreteNum[key] = numEls[i] as unknown as Element;
    }
    return { abstractNum, concreteNum };
  }

  function listDocAst(): DocumentAST {
    return buildTestAst(
      [
        {
          id: 'lb_1',
          type: 'paragraph',
          xml: numberedParaXml('First bullet', '1'),
          metadata: buildMeta('First bullet'),
          version: 1,
          checksum: 'chk_lb_1',
        },
        {
          id: 'lb_2',
          type: 'paragraph',
          xml: numberedParaXml('Second bullet', '1'),
          metadata: buildMeta('Second bullet'),
          version: 1,
          checksum: 'chk_lb_2',
        },
        {
          id: 'lo_1',
          type: 'paragraph',
          xml: numberedParaXml('First step', '2'),
          metadata: buildMeta('First step'),
          version: 1,
          checksum: 'chk_lo_1',
        },
        {
          id: 'lo_2',
          type: 'paragraph',
          xml: numberedParaXml('Second step', '2'),
          metadata: buildMeta('Second step'),
          version: 1,
          checksum: 'chk_lo_2',
        },
        {
          id: 'p_plain',
          type: 'paragraph',
          xml: parseXml(`<w:p xmlns:w="${W_NS}"><w:r><w:t>Plain paragraph</w:t></w:r></w:p>`),
          metadata: buildMeta('Plain paragraph'),
          version: 1,
          checksum: 'chk_plain',
        },
      ],
      { numbering: numberingMaps() }
    );
  }

  function stripLists(html: string): string {
    return html
      .replace(/<ul[\s\S]*?<\/ul>/g, '')
      .replace(/<ol[\s\S]*?<\/ol>/g, '');
  }

  it('renders numbered paragraphs with markers in Markdown and leaves plain text bare', () => {
    const md = convertAstToMarkdown(listDocAst());
    expect(md).toBe('- First bullet\n- Second bullet\n1. First step\n1. Second step\n\nPlain paragraph\n');
    expect(md.split('\n').filter((l) => l.startsWith('- '))).toEqual(['- First bullet', '- Second bullet']);
    expect(md).not.toContain('- Plain paragraph');
    expect(md).not.toMatch(/^1\. Plain paragraph/m);
  });

  it('groups numbered paragraphs into list containers in HTML with no bare items', () => {
    const html = convertAstToHtml(listDocAst(), 'ListDoc');
    const listDoc = new DOMParser().parseFromString(html, 'text/html');
    expect(listDoc.getElementsByTagName('ul').length).toBe(1);
    expect(listDoc.getElementsByTagName('ol').length).toBe(1);
    expect(listDoc.getElementsByTagName('li').length).toBe(4);
    expect(Array.from(listDoc.getElementsByTagName('li')).map((e) => (e as unknown as Element).textContent)).toEqual([
      'First bullet',
      'Second bullet',
      'First step',
      'Second step',
    ]);
    expect(html).toContain('<ul class="doc-list">');
    expect(html).toContain('<ol class="doc-list">');
    expect(stripLists(html)).not.toContain('<li>');
    expect(html.indexOf('<ul class="doc-list">')).toBeLessThan(html.indexOf('<ol class="doc-list">'));
    expect(html).toContain('<p class="paragraph">Plain paragraph</p>');
    expect(html).not.toContain('<p class="paragraph">First bullet</p>');
  });

  it('honors explicit bullet and ordered styles without numbering definitions', () => {
    const ast = buildTestAst([
      {
        id: 'lb_x',
        type: 'paragraph',
        xml: numberedParaXml('Styled bullet', '9'),
        metadata: buildMeta('Styled bullet', { listStyle: 'bullet' }),
        version: 1,
        checksum: 'chk_styled_bullet',
      },
      {
        id: 'lo_x',
        type: 'paragraph',
        xml: numberedParaXml('Styled step', '9'),
        metadata: buildMeta('Styled step', { listStyle: 'decimal' }),
        version: 1,
        checksum: 'chk_styled_step',
      },
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('- Styled bullet');
    expect(md).toContain('1. Styled step');
    const html = convertAstToHtml(ast, 'StyledLists');
    expect(html).toContain('<ul class="doc-list">');
    expect(html).toContain('<ol class="doc-list">');
  });

  it('falls back to bullets when numbering definitions are unreachable', () => {
    const ast = buildTestAst([
      {
        id: 'lb_u',
        type: 'paragraph',
        xml: numberedParaXml('Unknown sequence item', '99'),
        metadata: buildMeta('Unknown sequence item'),
        version: 1,
        checksum: 'chk_unknown_seq',
      },
    ]);
    const md = convertAstToMarkdown(ast);
    expect(md).toContain('- Unknown sequence item');
    const html = convertAstToHtml(ast, 'UnknownSeq');
    expect(html).toContain('<ul class="doc-list">');
    expect(html).toContain('<li>Unknown sequence item</li>');
  });

  it('round-trips a real document with numbering definitions through both exporters', async () => {
    const zip = new AdmZip();
    const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${W_NS}">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>Saved bullet</w:t></w:r></w:p>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr></w:pPr><w:r><w:t>Saved step</w:t></w:r></w:p>
    <w:p><w:r><w:t>Saved plain</w:t></w:r></w:p>
  </w:body>
</w:document>`;
    const numberingXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:numbering xmlns:w="${W_NS}">
  <w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>
  <w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
</w:numbering>`;
    zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
    zip.addFile('word/numbering.xml', Buffer.from(numberingXml, 'utf8'));
    zip.addFile('[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>', 'utf8'));
    const parser = new DocxParser(zip.toBuffer());
    const ast = await parser.getAST();
    expect(ast.sections[0].nodes.every((n: any) => n.type === 'paragraph')).toBe(true);

    const md = convertAstToMarkdown(ast, parser);
    expect(md).toContain('- Saved bullet');
    expect(md).toContain('1. Saved step');
    expect(md).toContain('Saved plain');
    expect(md).not.toContain('- Saved plain');

    const html = convertAstToHtml(ast, 'SavedLists', parser);
    expect(html).toContain('<ul class="doc-list">');
    expect(html).toContain('<ol class="doc-list">');
    expect(html).toContain('<li>Saved bullet</li>');
    expect(html).toContain('<li>Saved step</li>');
    expect(html).toContain('<p class="paragraph">Saved plain</p>');
  });
});

describe('Code blocks coexist with dividers in export', () => {
  it('code exports as a fenced block with its language; real dividers still export ---', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST() as any;
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'line one\nline two', id: 'code_1',
      params: { codeBlock: true, codeLanguage: 'ts' },
    } as any, ast, parser);
    await executeInsert({
      action: 'insert', type: 'divider', anchor: 'code_1', position: 'after', id: 'div_1',
    } as any, ast, parser);

    const md = convertAstToMarkdown(ast, parser);
    expect(md).toContain('```ts\nline one\nline two\n```');
    expect(md).toContain('---\n');

    const html = convertAstToHtml(ast, 'CodeDiv', parser);
    expect(html).toContain('<pre><code class="language-ts">line one\nline two</code></pre>');
    expect(html).toContain('<hr class="doc-divider"/>');
  });

  it('single-line code exports as a fenced block with its language; real dividers still export ---', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST() as any;
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({
      action: 'insert', type: 'paragraph', anchor, position: 'after',
      text: 'console.log("hello");', id: 'code_2',
      params: { codeBlock: true, codeLanguage: 'js' },
    } as any, ast, parser);
    await executeInsert({
      action: 'insert', type: 'divider', anchor: 'code_2', position: 'after', id: 'div_2',
    } as any, ast, parser);

    const md = convertAstToMarkdown(ast, parser);
    expect(md).toContain('```js\nconsole.log("hello");\n```');
    expect(md).not.toContain('`console.log("hello");`');
    expect(md).toContain('---\n');
  });
});
