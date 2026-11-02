import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';
import { DocxParser } from '../src/parser';

/**
 * A new document gets Word's usual paragraph rhythm: 6 pt after each paragraph and 1.15 line spacing, set once
 * in `docDefaults`. Table text keeps single spacing with no gap, through the `TableGrid` style that the tables
 * the engine writes already name. A document that is opened is not given any of this.
 */
function stylesOf(buf: Buffer): Document {
  const xml = new AdmZip(buf).readAsText('word/styles.xml');
  return new DOMParser().parseFromString(xml, 'text/xml');
}

function spacingOf(parent: Element | undefined): Record<string, string | null> | undefined {
  const spacing = parent?.getElementsByTagName('w:spacing')[0];
  if (!spacing) return undefined;
  return {
    after: spacing.getAttribute('w:after'),
    line: spacing.getAttribute('w:line'),
    lineRule: spacing.getAttribute('w:lineRule'),
  };
}

const styleById = (doc: Document, id: string): Element | undefined =>
  Array.from(doc.getElementsByTagName('w:style')).find((s) => s.getAttribute('w:styleId') === id);

describe('new documents: default paragraph spacing', () => {
  it.each(['blank', 'report', 'memo'])('the %s template sets 6 pt after and 1.15 lines in docDefaults', (template) => {
    const doc = stylesOf(DocxParser.createNewDocument(template).rezip());
    const defaults = doc.getElementsByTagName('w:docDefaults')[0];
    const pPrDefault = defaults.getElementsByTagName('w:pPrDefault')[0];
    expect(pPrDefault).toBeDefined();
    expect(spacingOf(pPrDefault)).toEqual({ after: '120', line: '276', lineRule: 'auto' });
    const order = Array.from(defaults.childNodes)
      .filter((n) => n.nodeType === 1)
      .map((e) => (e as Element).tagName);
    expect(order).toEqual(['w:rPrDefault', 'w:pPrDefault']);
  });

  it('the Normal style adds nothing of its own, so body text follows the default', () => {
    const doc = stylesOf(DocxParser.createNewDocument('blank').rezip());
    expect(styleById(doc, 'Normal')?.getElementsByTagName('w:pPr').length).toBe(0);
  });

  it('TableGrid keeps text in tables at single spacing with no gap after', () => {
    const doc = stylesOf(DocxParser.createNewDocument('blank').rezip());
    const grid = styleById(doc, 'TableGrid');
    expect(grid).toBeDefined();
    expect(grid!.getAttribute('w:type')).toBe('table');
    expect(spacingOf(grid)).toEqual({ after: '0', line: '240', lineRule: 'auto' });
    expect(grid!.getElementsByTagName('w:basedOn')[0].getAttribute('w:val')).toBe('TableNormal');
    const normal = styleById(doc, 'TableNormal');
    expect(normal?.getAttribute('w:default')).toBe('1');
  });

  it('every table style the engine names exists in the styles part', async () => {
    const parser = DocxParser.createNewDocument('report');
    const zip = new AdmZip(parser.rezip());
    const used = new Set(Array.from(zip.readAsText('word/document.xml').matchAll(/<w:tblStyle w:val="([^"]+)"/g), (m) => m[1]));
    const doc = stylesOf(parser.rezip());
    for (const id of used) expect(styleById(doc, id), `style ${id}`).toBeDefined();
  });
});

describe('opened documents are left as they are', () => {
  it('a document without pPrDefault does not gain one when it is opened and saved', async () => {
    const source = new AdmZip(DocxParser.createNewDocument('blank').rezip());
    const stripped = source
      .readAsText('word/styles.xml')
      .replace(/<w:pPrDefault>[\s\S]*?<\/w:pPrDefault>/, '');
    expect(stripped).not.toContain('pPrDefault');
    source.updateFile('word/styles.xml', Buffer.from(stripped, 'utf8'));

    const parser = new DocxParser(source.toBuffer());
    await parser.getAST();
    const saved = new AdmZip(parser.rezip()).readAsText('word/styles.xml');
    expect(saved).not.toContain('pPrDefault');
  });
});
