/**
 * 0.5.3: every writer that adds a relationship works when the relationships part binds the package namespace to a
 * prefix (`<p:Relationship>`). The old writers looked for `Relationships` / `Relationship` by plain name, found
 * nothing, and either skipped the write (a hyperlink with no target) or picked an Id that was already taken.
 */
import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, composeDocument, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';

const PKG = 'http://schemas.openxmlformats.org/package/2006/relationships';
const prefixedXml = (xml: string): string => xml
  .replace(/<Relationships\b/, '<p:Relationships').replace(/<\/Relationships>/, '</p:Relationships>')
  .replace(/xmlns=/, 'xmlns:p=').replace(/<Relationship\b/g, '<p:Relationship');

function prefixedPackage(parts: string[] = ['word/_rels/document.xml.rels']): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  for (const part of parts) {
    const text = zip.readAsText(part);
    const next = prefixedXml(text);
    expect(next).toContain('<p:Relationship ');
    zip.updateFile(part, Buffer.from(next, 'utf8'));
  }
  return zip.toBuffer();
}

async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast: any = await parser.getAST();
  const engine = new OperationEngine(ast, parser);
  const run = (operations: unknown[]): Promise<any> => engine.executeBatch({ operations } as any);
  const rels = (part = 'word/_rels/document.xml.rels'): Element[] => {
    const doc = new DOMParser().parseFromString(parser.getZipAdapter().readText(part), 'text/xml') as unknown as Document;
    return Array.from(doc.getElementsByTagNameNS('*', 'Relationship'));
  };
  return { parser, ast, run, rels, anchor: ast.sections[0].nodes[0].id as string };
}
type Session = Awaited<ReturnType<typeof open>>;

/** Every relationship element sits in the package namespace and every Id is unique. */
function expectSound(s: Session, part = 'word/_rels/document.xml.rels'): void {
  const list = s.rels(part);
  const ids = list.map((r) => r.getAttribute('Id'));
  expect(new Set(ids).size).toBe(ids.length);
  for (const r of list) expect(r.namespaceURI).toBe(PKG);
}
const targetsOfType = (s: Session, fragment: string, part?: string): string[] =>
  s.rels(part).filter((r) => (r.getAttribute('Type') ?? '').includes(fragment)).map((r) => r.getAttribute('Target') ?? '');

describe('relationship writers on a prefixed relationships part', () => {
  it('hyperlink operation registers the link target', async () => {
    const s = await open(prefixedPackage());
    const res = await s.run([{ action: 'hyperlink', target: s.anchor, text: 'Portal', url: 'https://astra.example/portal' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(targetsOfType(s, '/hyperlink')).toEqual(['https://astra.example/portal']);
    expectSound(s);
  });

  it('footnote insert registers the footnotes and settings parts', async () => {
    const s = await open(prefixedPackage());
    const res = await s.run([{ action: 'footnote', operation: 'insert', anchor: s.anchor, text: 'A note.' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(targetsOfType(s, '/footnotes')).toEqual(['footnotes.xml']);
    expectSound(s);
  });

  it('comment insert registers the comments part without reusing an Id', async () => {
    const s = await open(prefixedPackage());
    const res = await s.run([
      { action: 'hyperlink', target: s.anchor, text: 'Portal', url: 'https://astra.example/a' },
      { action: 'comment', operation: 'insert', target: s.anchor, text: 'Note', author: 'Dana', date: '2026-10-24T09:30:00Z' },
    ]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(targetsOfType(s, '/comments')).toContain('comments.xml');
    expectSound(s);
  });

  it('header set registers the header part', async () => {
    const s = await open(prefixedPackage());
    const res = await s.run([{ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'Header text' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(targetsOfType(s, '/header').length).toBe(1);
    expectSound(s);
  });

  it('chart insert takes a free Id next to prefixed relationships and writes a prefixed chart relationship', async () => {
    const s = await open(prefixedPackage());
    const before = s.rels().map((r) => r.getAttribute('Id'));
    const res = await s.run([{
      action: 'chart', operation: 'insert', anchor: s.anchor, id: 'c1', chartType: 'column', title: 'T',
      categories: ['a', 'b'], series: [{ name: 'S', values: [1, 2] }],
    }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    const added = s.rels().filter((r) => !before.includes(r.getAttribute('Id')));
    expect(added.some((r) => (r.getAttribute('Type') ?? '').endsWith('/chart'))).toBe(true);
    expectSound(s);
    const chartRels = s.parser.getZipAdapter().entryNames!().filter((n) => /^word\/charts\/_rels\//.test(n));
    for (const part of chartRels) expectSound(s, part);
  });

  it('markdown link registers its target', async () => {
    const s = await open(prefixedPackage());
    const res = await composeDocument(s.parser, 'See [the site](https://astra.example/md) now.');
    expect(res.success, JSON.stringify((res as any).errors)).toBe(true);
    expect(targetsOfType(s, '/hyperlink')).toContain('https://astra.example/md');
    expectSound(s);
  });

  it('setting properties adds the core-properties relationship to a prefixed package root', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    zip.deleteFile('docProps/core.xml');
    const bare = zip.readAsText('_rels/.rels').replace(/<Relationship\b[^>]*core-properties[^>]*\/>/g, '');
    zip.updateFile('_rels/.rels', Buffer.from(prefixedXml(bare), 'utf8'));
    const types = zip.readAsText('[Content_Types].xml').replace(/<Override\b[^>]*\/docProps\/core\.xml[^>]*\/>/g, '');
    zip.updateFile('[Content_Types].xml', Buffer.from(types, 'utf8'));
    const s = await open(zip.toBuffer());
    const res = await s.run([{ action: 'metadata', operation: 'setProperties', title: 'Title' }]);
    expect(res.success, JSON.stringify(res.errors)).toBe(true);
    expect(targetsOfType(s, 'core-properties', '_rels/.rels')).toEqual(['docProps/core.xml']);
    expectSound(s, '_rels/.rels');
  });

  it('the package stays valid after all of them', async () => {
    const s = await open(prefixedPackage());
    await s.run([
      { action: 'hyperlink', target: s.anchor, text: 'Portal', url: 'https://astra.example/a' },
      { action: 'footnote', operation: 'insert', anchor: s.anchor, text: 'A note.' },
      { action: 'chart', operation: 'insert', anchor: s.anchor, id: 'c1', chartType: 'column', title: 'T', categories: ['a'], series: [{ name: 'S', values: [1] }] },
    ]);
    const report = await validateDocxBuffer(s.parser.rezip(s.ast));
    expect(report.errors).toEqual([]);
  });
});
