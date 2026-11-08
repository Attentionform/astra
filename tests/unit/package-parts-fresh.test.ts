import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer } from '@attentionform/astra-engine';

const DATE = '2026-01-01T00:00:00Z';
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function withPart(kind: 'header' | 'footer', text: string, id = 4): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const tag = kind === 'header' ? 'hdr' : 'ftr';
  zip.addFile(`word/${kind}1.xml`, Buffer.from(`<w:${tag} xmlns:w="${W}"><w:p><w:ins w:id="${id}" w:author="A" w:date="${DATE}"><w:r><w:t>${text}</w:t></w:r></w:ins></w:p></w:${tag}>`));
  zip.updateFile('word/_rels/document.xml.rels', Buffer.from(zip.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', `<Relationship Id="rId${kind}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${kind}" Target="${kind}1.xml"/></Relationships>`)));
  zip.updateFile('[Content_Types].xml', Buffer.from(zip.readAsText('[Content_Types].xml').replace('</Types>', `<Override PartName="/word/${kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.${kind}+xml"/></Types>`)));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace('<w:sectPr>', `<w:sectPr><w:${kind}Reference w:type="default" r:id="rId${kind}"/>`)));
  return zip.toBuffer();
}

async function open(buffer: Buffer) {
  const parser = new DocxParser(buffer);
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}

describe('package parts are read from the current package, never from a stale copy (FX-01)', () => {
  for (const kind of ['header', 'footer'] as const) {
    it(`${kind}: an edit made before acceptAll survives acceptAll`, async () => {
      const { parser, engine } = await open(withPart(kind, `OLD ${kind.toUpperCase()}`));
      const set = await engine.executeBatch({ operations: [{ action: 'headerFooter', operation: 'set', targetType: kind, text: `NEW ${kind.toUpperCase()}` }] } as never);
      expect(set.success, JSON.stringify(set.errors)).toBe(true);
      const accept = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
      expect(accept.success, JSON.stringify(accept.errors)).toBe(true);
      const xml = parser.getZipAdapter().readText(`word/${kind}1.xml`);
      expect(xml).toContain(`NEW ${kind.toUpperCase()}`);
      expect(xml).not.toContain(`OLD ${kind.toUpperCase()}`);
    });
  }

  it('a rolled back batch leaves the revision count and the part consistent', async () => {
    const { parser, engine } = await open(withPart('header', 'HDR'));
    const before = parser.countPackageRevisions();
    expect(before).toBe(1);
    const failed = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }, { action: 'replace', target: 'NOSUCHNODE', text: 'x' }] } as never);
    expect(failed.success).toBe(false);
    expect(parser.countPackageRevisions()).toBe(before);
    expect(parser.getZipAdapter().readText('word/header1.xml')).toContain('<w:ins');
    const again = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(again.success, JSON.stringify(again.errors)).toBe(true);
    expect(parser.getZipAdapter().readText('word/header1.xml')).not.toContain('<w:ins');
    expect(parser.countPackageRevisions()).toBe(0);
  });

  it('a part added after the first count is counted', async () => {
    const { parser } = await open(createBlankDocxBuffer('blank'));
    expect(parser.countPackageRevisions()).toBe(0);
    const zip = parser.getZipAdapter();
    zip.writeText('word/footer9.xml', `<w:ftr xmlns:w="${W}"><w:p><w:ins w:id="77" w:author="A" w:date="${DATE}"><w:r><w:t>late</w:t></w:r></w:ins></w:p></w:ftr>`);
    zip.writeText('[Content_Types].xml', zip.readText('[Content_Types].xml').replace('</Types>', '<Override PartName="/word/footer9.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'));
    expect(parser.countPackageRevisions()).toBe(1);
  });

  it('a change written to a part after the first read is seen by the next read and resolved once', async () => {
    const { parser, engine } = await open(withPart('footer', 'FIRST'));
    expect(parser.countPackageRevisions()).toBe(1);
    parser.getZipAdapter().writeText('word/footer1.xml', `<w:ftr xmlns:w="${W}"><w:p><w:ins w:id="4" w:author="A" w:date="${DATE}"><w:r><w:t>SECOND</w:t></w:r></w:ins><w:ins w:id="5" w:author="A" w:date="${DATE}"><w:r><w:t>THIRD</w:t></w:r></w:ins></w:p></w:ftr>`);
    expect(parser.countPackageRevisions()).toBe(2);
    const accept = await engine.executeBatch({ operations: [{ action: 'trackedChange', operation: 'acceptAll' }] } as never);
    expect(accept.success, JSON.stringify(accept.errors)).toBe(true);
    const xml = parser.getZipAdapter().readText('word/footer1.xml');
    expect(xml).toContain('SECOND');
    expect(xml).toContain('THIRD');
    expect(xml).not.toContain('<w:ins');
    expect(parser.countPackageRevisions()).toBe(0);
  });
});
