/**
 * Deleting a section break must not leave the section's header or footer text
 * inside the saved file. The pass that removes those parts may only remove a
 * part nothing refers to: every case where something still refers to it (a
 * shared part, a tracked section change, another relationships file) keeps it,
 * anything it cannot read skips it, and a failure part-way puts everything back.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser, OperationEngine, assertValidDocx, type DocumentAST } from '@attentionform/astra-engine';
import {
  danglingReferences,
  headerFooterRefIds,
  pruneOrphanedHeaderFooters,
  type PackageParts,
} from '../../packages/engine/src/parser/header-footer-prune';
import { docxFromBody, para, sectPr } from '../helpers/section-docs';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PAGE = '<w:pgSz w:w="12240" w:h="15840"/>';
const ref = (id: string) => `<w:headerReference w:type="default" r:id="${id}"/>`;

/** An in-memory package the pure pass can work on; `log` records what it removed. */
function partsOf(docx: Buffer, options: { failRemoving?: string } = {}): { parts: PackageParts; files: Map<string, Buffer> } {
  const files = new Map<string, Buffer>();
  for (const entry of new AdmZip(docx).getEntries()) files.set(entry.entryName, entry.getData());
  const parts: PackageParts = {
    names: () => [...files.keys()],
    read: (name) => files.get(name) ?? null,
    write: (name, data) => void files.set(name, data),
    remove: (name) => {
      if (name === options.failRemoving) throw new Error(`cannot remove ${name}`);
      files.delete(name);
    },
  };
  return { parts, files };
}

const THREE = [
  { name: 'header1.xml', id: 'rId5', text: 'SECRET ONE' },
  { name: 'header2.xml', id: 'rId6', text: 'SECRET TWO' },
  { name: 'header3.xml', id: 'rId7', text: 'KEEP THREE' },
];
const threeSections = () =>
  docxFromBody(
    para('one', sectPr(ref('rId5') + PAGE)) + para('two', sectPr(ref('rId6') + PAGE)) + para('three') + sectPr(ref('rId7') + PAGE),
    THREE,
  );
const documentOf = (files: Map<string, Buffer>) => files.get('word/document.xml')!.toString('utf8');
const noGuard = () => undefined;
/** The edit itself: the sections that used these ids are gone from the document. */
const dropReferences = (files: Map<string, Buffer>, ids: string[]) => {
  let xml = documentOf(files);
  for (const id of ids) xml = xml.replace(new RegExp(`<w:headerReference[^>]*r:id="${id}"[^>]*/>`, 'g'), '');
  files.set('word/document.xml', Buffer.from(xml));
};
const run = (files: Map<string, Buffer>, parts: PackageParts, before: string[], now: string[], guard = noGuard) =>
  pruneOrphanedHeaderFooters({
    parts,
    documentXml: () => documentOf(files),
    referencedBefore: new Set(before),
    referencedNow: new Set(now),
    guard,
  });
const snapshot = (files: Map<string, Buffer>) => [...files.entries()].map(([n, b]) => [n, b.toString('base64')]).sort();

describe('the pass: what it removes', () => {
  it('removes a part that was referenced and no longer is, with its relationship and content-type override', () => {
    const { parts, files } = partsOf(threeSections());
    dropReferences(files, ['rId5']);
    const result = run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId6', 'rId7']);
    expect(result).toEqual({ relationshipIds: ['rId5'], parts: ['word/header1.xml'] });
    expect(files.has('word/header1.xml')).toBe(false);
    expect(files.get('word/_rels/document.xml.rels')!.toString()).not.toContain('rId5');
    expect(files.get('[Content_Types].xml')!.toString()).not.toContain('header1.xml');
    // the others are exactly as they were
    for (const kept of ['header2.xml', 'header3.xml']) {
      expect(files.has(`word/${kept}`)).toBe(true);
      expect(files.get('[Content_Types].xml')!.toString()).toContain(kept);
    }
    expect([...danglingReferences(parts, documentOf(files))]).toEqual([]);
  });

  it('removes a footer part the same way', () => {
    const docx = docxFromBody(para('a', sectPr(PAGE)) + sectPr(PAGE), []);
    const zip = new AdmZip(docx);
    zip.addFile('word/footer1.xml', Buffer.from('<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>SECRET FOOTER</w:t></w:r></w:p></w:ftr>'));
    zip.updateFile(
      'word/_rels/document.xml.rels',
      Buffer.from(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="${R}/footer" Target="footer1.xml"/></Relationships>`),
    );
    const { parts, files } = partsOf(zip.toBuffer());
    const result = run(files, parts, ['rId9'], []);
    expect(result.parts).toEqual(['word/footer1.xml']);
    expect(files.has('word/footer1.xml')).toBe(false);
  });

  it('leaves alone a part that was already unreferenced before the edit', () => {
    const { parts, files } = partsOf(threeSections());
    dropReferences(files, ['rId5']);
    const before = snapshot(files);
    // rId5 was not referenced before this edit either, so it is not ours to remove
    expect(run(files, parts, ['rId6', 'rId7'], ['rId6', 'rId7'])).toEqual({ relationshipIds: [], parts: [] });
    expect(snapshot(files)).toEqual(before);
  });

  it('removes nothing when no reference was lost', () => {
    const { parts, files } = partsOf(threeSections());
    const before = snapshot(files);
    expect(run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId5', 'rId6', 'rId7']).parts).toEqual([]);
    expect(snapshot(files)).toEqual(before);
  });
});

describe('the pass: what keeps a part', () => {
  const keeps = (mutate: (files: Map<string, Buffer>) => void, label: string) =>
    it(`keeps it when ${label}`, () => {
      const { parts, files } = partsOf(threeSections());
      dropReferences(files, ['rId5']);
      mutate(files);
      const before = snapshot(files);
      const result = run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId6', 'rId7']);
      expect(result.parts, label).toEqual([]);
      expect(snapshot(files)).toEqual(before);
    });

  keeps((files) => {
    files.set('word/document.xml', Buffer.from(documentOf(files).replace('<w:body>', '<w:body><w:p><w:r><w:foo r:id="rId5"/></w:r></w:p>')));
  }, 'the document still uses the relationship id somewhere else');

  keeps((files) => {
    files.set(
      'word/_rels/document.xml.rels',
      Buffer.from(files.get('word/_rels/document.xml.rels')!.toString().replace('</Relationships>', `<Relationship Id="rId99" Type="${R}/header" Target="header1.xml"/></Relationships>`)),
    );
  }, 'a second relationship points at the same part');

  keeps((files) => {
    files.set('word/_rels/footnotes.xml.rels', Buffer.from(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/header" Target="header1.xml"/></Relationships>`));
  }, "another part's relationships file mentions it");

  keeps((files) => {
    files.set(
      'word/_rels/document.xml.rels',
      Buffer.from(files.get('word/_rels/document.xml.rels')!.toString().replace('Target="header1.xml"', 'Target="settings.xml"')),
    );
  }, 'the relationship does not point at a headerN.xml part');

  keeps((files) => {
    files.set(
      'word/_rels/document.xml.rels',
      Buffer.from(files.get('word/_rels/document.xml.rels')!.toString().replace('Target="header1.xml"', 'Target="../header1.xml"')),
    );
  }, 'the target climbs out of word/');

  keeps((files) => {
    files.set(
      'word/_rels/document.xml.rels',
      Buffer.from(files.get('word/_rels/document.xml.rels')!.toString().replace('Target="header1.xml"', 'Target="header1.xml" TargetMode="External"')),
    );
  }, 'the relationship is external');

  keeps((files) => {
    files.set(
      'word/_rels/document.xml.rels',
      Buffer.from(files.get('word/_rels/document.xml.rels')!.toString().replace(`${R}/header"`, `${R}/hyperlink"`)),
    );
  }, 'the relationship is not a header or footer relationship');

  it('counts a header reference copied into a tracked section-property change as still in use', () => {
    const doc = new DOMParser().parseFromString(
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="${R}"><w:body><w:p><w:pPr><w:sectPr>${ref('rId6')}<w:sectPrChange w:id="1"><w:sectPr>${ref('rId5')}</w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p></w:body></w:document>`,
      'text/xml',
    );
    expect([...headerFooterRefIds(doc)].sort()).toEqual(['rId5', 'rId6']);
  });
});

describe('the pass: when it cannot be sure', () => {
  it('skips, and changes nothing, when the relationships file is unreadable', () => {
    const { parts, files } = partsOf(threeSections());
    dropReferences(files, ['rId5']);
    files.set('word/_rels/document.xml.rels', Buffer.from('<Relationships><Relationship Id="rId5"'));
    const before = snapshot(files);
    const result = run(files, parts, ['rId5'], []);
    expect(result.parts).toEqual([]);
    expect(result.skipped).toBeTruthy();
    expect(snapshot(files)).toEqual(before);
  });

  it('skips, and changes nothing, when a part is refused by the XML guard', () => {
    const { parts, files } = partsOf(threeSections());
    dropReferences(files, ['rId5']);
    const before = snapshot(files);
    const result = run(files, parts, ['rId5'], [], () => {
      throw new Error('DTD not allowed');
    });
    expect(result.parts).toEqual([]);
    expect(result.skipped).toContain('DTD');
    expect(snapshot(files)).toEqual(before);
  });

  it('puts every entry back when it fails part-way', () => {
    const { parts, files } = partsOf(threeSections(), { failRemoving: 'word/header1.xml' });
    dropReferences(files, ['rId5']);
    const before = snapshot(files);
    const result = run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId6', 'rId7']);
    expect(result.parts).toEqual([]);
    expect(result.skipped).toContain('cannot remove');
    expect(snapshot(files)).toEqual(before);
  });

  it('does not blame itself for references that were already broken', () => {
    const { parts, files } = partsOf(threeSections());
    dropReferences(files, ['rId5']);
    files.set(
      'word/_rels/document.xml.rels',
      Buffer.from(files.get('word/_rels/document.xml.rels')!.toString().replace('</Relationships>', `<Relationship Id="rId50" Type="${R}/image" Target="media/missing.png"/></Relationships>`)),
    );
    const result = run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId6', 'rId7']);
    expect(result.parts).toEqual(['word/header1.xml']);
  });
});

describe('the pass: pictures in a removed header', () => {
  const withPicture = (sharedWithDocument: boolean) => {
    const zip = new AdmZip(threeSections());
    zip.addFile('word/media/logo.png', Buffer.from('PNG-BYTES-1'));
    zip.addFile(
      'word/_rels/header1.xml.rels',
      Buffer.from(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/image" Target="media/logo.png"/></Relationships>`),
    );
    if (sharedWithDocument) {
      const rels = zip.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', `<Relationship Id="rId40" Type="${R}/image" Target="media/logo.png"/></Relationships>`);
      zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
    }
    return partsOf(zip.toBuffer());
  };

  it('removes a picture only the removed header used, and the header\'s relationships file', () => {
    const { parts, files } = withPicture(false);
    dropReferences(files, ['rId5']);
    const result = run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId6', 'rId7']);
    expect(result.parts.sort()).toEqual(['word/header1.xml', 'word/media/logo.png']);
    expect(files.has('word/media/logo.png')).toBe(false);
    expect(files.has('word/_rels/header1.xml.rels')).toBe(false);
  });

  it('keeps a picture the document also uses', () => {
    const { parts, files } = withPicture(true);
    dropReferences(files, ['rId5']);
    const result = run(files, parts, ['rId5', 'rId6', 'rId7'], ['rId6', 'rId7']);
    expect(result.parts).toEqual(['word/header1.xml']);
    expect(files.has('word/media/logo.png')).toBe(true);
    expect(files.has('word/_rels/header1.xml.rels')).toBe(false);
  });
});

interface Session {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}
const open = async (buf: Buffer): Promise<Session> => {
  const parser = new DocxParser(buf);
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
};
const everyText = (docx: Buffer): string =>
  new AdmZip(docx)
    .getEntries()
    .map((e) => e.getData().toString('latin1'))
    .join('\n');
const partNames = (docx: Buffer) => new AdmZip(docx).getEntries().map((e) => e.entryName);
/** The last paragraph of a section is the one that carries its break. */
const breakOf = (s: Session, section: number) => s.ast.sections[section].nodes.slice(-1)[0].id;

describe('deleting a section break in the engine', () => {
  for (const [label, index, gone, kept] of [
    ['the first', 0, 'SECRET ONE', ['SECRET TWO', 'KEEP THREE']],
    ['the second', 1, 'SECRET TWO', ['SECRET ONE', 'KEEP THREE']],
  ] as const) {
    it(`${label} break takes its header text out of the saved file and leaves a valid file`, async () => {
      const s = await open(threeSections());
      const out = await s.engine.executeBatch({ operations: [{ action: 'delete', target: breakOf(s, index) }] as never });
      expect(out.success, JSON.stringify(out.errors)).toBe(true);

      const saved = s.parser.rezip();
      expect(everyText(saved)).not.toContain(gone);
      for (const text of kept) expect(everyText(saved)).toContain(text);
      expect(partNames(saved).filter((n) => /^word\/header\d+\.xml$/.test(n))).toHaveLength(2);
      await assertValidDocx(saved);

      // the model agrees with the file, and saving again changes nothing
      expect(Object.values(s.ast.relationships).some((r) => r.target.startsWith('header'))).toBe(true);
      expect(Object.keys(s.ast.relationships)).toHaveLength(2);
      expect(partNames(s.parser.rezip()).sort()).toEqual(partNames(saved).sort());

      // a fresh process opening the saved file sees the same thing
      const reopened = await open(saved);
      expect(everyText(reopened.parser.rezip())).not.toContain(gone);
      expect(reopened.ast.sections).toHaveLength(2);
    });
  }

  it('a failed batch puts the removed header back, bytes and all', async () => {
    const s = await open(threeSections());
    const before = s.parser.rezip();
    const out = await s.engine.executeBatch({
      operations: [{ action: 'delete', target: breakOf(s, 0) }, { action: 'replace', target: 'no_such_node', text: 'x' }] as never,
      sequential: true,
    });
    expect(out.success).toBe(false);
    const after = s.parser.rezip();
    expect(everyText(after)).toContain('SECRET ONE');
    expect(partNames(after).sort()).toEqual(partNames(before).sort());
    expect(Object.keys(s.ast.relationships).sort()).toEqual(['rId5', 'rId6', 'rId7']);
    await assertValidDocx(after);
  });

  it('a single operation (not a batch) prunes the same way', async () => {
    const s = await open(threeSections());
    await s.engine.execute({ action: 'delete', target: breakOf(s, 0) } as never);
    expect(everyText(s.parser.rezip())).not.toContain('SECRET ONE');
  });

  it('a dry run reports the pruned file but changes nothing', async () => {
    const s = await open(threeSections());
    const out = await s.engine.executeBatch({ operations: [{ action: 'delete', target: breakOf(s, 0) }] as never, dryRun: true });
    expect(out.success).toBe(true);
    expect(everyText(s.parser.rezip())).toContain('SECRET ONE');
  });

  it('an edit that touches no section break leaves a leftover part from another tool alone', async () => {
    const zip = new AdmZip(threeSections());
    zip.addFile('word/header9.xml', Buffer.from('<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>OLD LEFTOVER</w:t></w:r></w:p></w:hdr>'));
    const s = await open(zip.toBuffer());
    const target = s.ast.sections[2].nodes[0].id;
    const out = await s.engine.executeBatch({ operations: [{ action: 'replace', target, text: 'changed' }] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(everyText(s.parser.rezip())).toContain('OLD LEFTOVER');
  });

  it('a header picture goes with its header', async () => {
    const zip = new AdmZip(threeSections());
    zip.addFile('word/media/stamp.png', Buffer.from('STAMP-BYTES'));
    zip.addFile(
      'word/_rels/header1.xml.rels',
      Buffer.from(`<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${R}/image" Target="media/stamp.png"/></Relationships>`),
    );
    const s = await open(zip.toBuffer());
    expect(everyText(s.parser.rezip())).toContain('STAMP-BYTES');
    await s.engine.executeBatch({ operations: [{ action: 'delete', target: breakOf(s, 0) }] as never });
    const saved = s.parser.rezip();
    expect(everyText(saved)).not.toContain('STAMP-BYTES');
    expect(partNames(saved)).not.toContain('word/_rels/header1.xml.rels');
  });
});
