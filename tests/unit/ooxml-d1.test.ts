/**
 * Batch D1 (Phase 3) — OOXML write-path round-trips.
 *
 * Pins save/reload behavior for lists (BX-03), headers/footers (BX-04),
 * sdt inserts (BX-11), schema-order comment preservation (BX-01), TOON cache
 * discipline (BX-08), cloneAST isolation (BX-02), loud coercion errors in
 * parser/manipulator scope (BX-10), save-path `both` normalization (BX-12),
 * and safe bookmark/rId allocation (OX-12 + BX-05).
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import {
  allocateNextBookmarkId,
  allocateSafeHyperlinkRId,
  bootstrapCommentParts,
  cloneAST,
  createBlankDocxBuffer,
  DocumentAST,
  DocxParser,
  executeList,
  HeaderFooterManipulator,
  NumberingManipulator,
  OperationEngine,
  RelationshipsManipulator,
  resetBookmarkHighWaterForTests,
  resetHyperlinkRIdStateForTests,
  RPR_CHILD_ORDER,
  sortChildrenBySchema,
  TOCManipulator,
  ToonGenerator,
} from '@attentionform/astra-engine';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

function stampListNumbering(
  paragraphEl: Element,
  numId: string,
  level = 0
): void {
  const doc = paragraphEl.ownerDocument;
  let pPr = paragraphEl.getElementsByTagName('w:pPr')[0] as Element | undefined;
  if (!pPr) {
    pPr = doc.createElement('w:pPr');
    paragraphEl.insertBefore(pPr, paragraphEl.firstChild);
  }
  let numPr = pPr.getElementsByTagName('w:numPr')[0] as Element | undefined;
  if (!numPr) {
    numPr = doc.createElement('w:numPr');
    pPr.appendChild(numPr);
  }
  let ilvl = numPr.getElementsByTagName('w:ilvl')[0] as Element | undefined;
  if (!ilvl) {
    ilvl = doc.createElement('w:ilvl');
    numPr.appendChild(ilvl);
  }
  ilvl.setAttribute('w:val', String(level));
  let numIdEl = numPr.getElementsByTagName('w:numId')[0] as Element | undefined;
  if (!numIdEl) {
    numIdEl = doc.createElement('w:numId');
    numPr.appendChild(numIdEl);
  }
  numIdEl.setAttribute('w:val', numId);
}

describe('D1 BX-03: numbering serialize-on-rezip + real continuation', () => {
  it('bullet and numbered lists get distinct numIds (no shared hardcoded 1)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    const bulletId = manip.getOrCreateList('bullet');
    const numberedId = manip.getOrCreateList('numbered');
    expect(bulletId).toBeTruthy();
    expect(numberedId).toBeTruthy();
    expect(numberedId).not.toBe(bulletId);
    // Repeat calls reuse (no XML bloat).
    expect(manip.getOrCreateList('bullet')).toBe(bulletId);
    expect(manip.getOrCreateList('numbered')).toBe(numberedId);
  });

  it('list.create survives save/reload with the same numId (numbering.xml written)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    const bulletId = manip.getOrCreateList('bullet');

    const first = ast.sections[0].nodes[0];
    const firstText = first.metadata.text;
    stampListNumbering(first.xml as Element, bulletId, 0);
    first.metadata.listLevel = 0;

    // Real continuation: resolves the anchor's own numId.
    expect(manip.continueList(first.id)).toBe(bulletId);

    const out = parser.rezip(ast);
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')).toBeTruthy();
    const numberingXml = zip.getEntry('word/numbering.xml')!.getData().toString('utf8');
    expect(numberingXml).toContain(`w:numId="${bulletId}"`);
    expect(zip.getEntry('[Content_Types].xml')!.getData().toString('utf8')).toContain('numbering');

    const parser2 = new DocxParser(out);
    const ast2 = await parser2.getAST();
    expect(Object.keys(ast2.numbering.concreteNum)).toContain(bulletId);
    expect(Object.keys(ast2.numbering.abstractNum).length).toBeGreaterThan(0);

    const reloaded = ast2.sections[0].nodes.find(n => n.metadata.text === firstText);
    expect(reloaded).toBeDefined();
    const numIdEl = (reloaded!.xml as Element).getElementsByTagName('w:numId')[0] as Element;
    expect(numIdEl.getAttribute('w:val')).toBe(bulletId);
    // Continuation still resolves after reload (not hardcoded '1').
    expect(new NumberingManipulator(ast2).continueList(reloaded!.id)).toBe(bulletId);

    expect(parser2.validateZipIntegrity().isValid).toBe(true);
  });

  it('continueList throws loudly on unknown anchors and non-list nodes', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    expect(() => manip.continueList('p_999_missing')).toThrow(/not found/);
    expect(() => manip.continueList(ast.sections[0].nodes[0].id)).toThrow(/no list numbering/);
    expect(() => manip.continueList('' as string)).toThrow(/requires a previous list node id/);
  });

  it('numbered list round-trips through the strict OOXML validator clean', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    const numberedId = manip.getOrCreateList('numbered');
    const first = ast.sections[0].nodes[0];
    stampListNumbering(first.xml as Element, numberedId, 0);
    const out = parser.rezip(ast);
    const probe = new DocxParser(out);
    const result = await probe.validateOOXML();
    expect(result.ok).toBe(true);
    expect(result.errors).toHaveLength(0);
  });

  it('bullet levels use a portable glyph with a common font', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    const bulletId = manip.getOrCreateList('bullet');
    const first = ast.sections[0].nodes[0];
    stampListNumbering(first.xml as Element, bulletId, 0);

    const numberingXml = manip.serializeNumbering() ?? '';
    expect(numberingXml).toContain('\u2022');
    expect(numberingXml).toContain('Arial');
    expect(numberingXml).not.toContain(String.fromCharCode(0xF0B7));
    expect(numberingXml).not.toContain('61623');
    expect(numberingXml).not.toContain('Symbol');

    const out = parser.rezip(ast);
    const saved = new AdmZip(out).getEntry('word/numbering.xml')!.getData().toString('utf8');
    expect(saved).toContain('\u2022');
    expect(saved).toContain('Arial');
    expect(saved).not.toContain(String.fromCharCode(0xF0B7));
    expect(saved).not.toContain('61623');
  });
});

describe('List op: create labels truthfully, continue joins the same sequence', () => {
  async function seededList(listType: 'bullet' | 'numbered') {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const eng = new OperationEngine(ast, parser);
    const seed = await eng.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'first item', id: 'li_1' },
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'second item', id: 'li_2' },
      ],
      sequential: true,
    });
    expect(seed.success).toBe(true);
    const created = await eng.executeBatch({
      operations: [{ action: 'list', operation: 'create', target: ['li_1', 'li_2'], listType }],
      sequential: true,
    });
    expect(created.success).toBe(true);
    return { parser, ast, eng };
  }

  function numIdsOf(ast: DocumentAST, ids: string[]): (string | undefined)[] {
    return ids.map((id) => {
      const node = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === id);
      const numIdEl = (node!.xml as Element).getElementsByTagName('w:numId')[0] as Element | undefined;
      return numIdEl?.getAttribute('w:val') ?? undefined;
    });
  }

  it('create records the list style; TOON labels numbered vs bullet truthfully', async () => {
    const { ast } = await seededList('numbered');
    const nodes = ast.sections.flatMap((s) => s.nodes);
    expect(nodes.find((n) => n.id === 'li_1')!.metadata.listStyle).toBe('numbered');
    const gen = new ToonGenerator();
    const rows = gen.toon(ast, { fidelity: 'standard' }).rows;
    expect(rows.find((r) => r.id === 'li_1')!.style).toBe('list:numbered:L1');
    expect(rows.find((r) => r.id === 'li_2')!.style).toBe('list:numbered:L1');
  });

  it('continue joins the preceding list instead of forking a new sequence', async () => {
    const { ast, eng } = await seededList('numbered');
    const added = await eng.executeBatch({
      operations: [{ action: 'insert', type: 'paragraph', anchor: 'end', text: 'third item', id: 'li_3' }],
      sequential: true,
    });
    expect(added.success).toBe(true);
    const continued = await eng.executeBatch({
      operations: [{ action: 'list', operation: 'continue', target: ['li_3'] }],
      sequential: true,
    });
    expect(continued.success).toBe(true);
    const [first, , third] = numIdsOf(ast, ['li_1', 'li_2', 'li_3']);
    expect(third).toBeDefined();
    expect(third).toBe(first);
    const gen = new ToonGenerator();
    expect(gen.toon(ast, { fidelity: 'standard' }).rows.find((r) => r.id === 'li_3')!.style).toBe('list:numbered:L1');
  });

  it('continue honors an explicit previousNodeId and fails loud with no list behind it', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const eng = new OperationEngine(ast, parser);
    const seed = await eng.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'solo', id: 'solo_1' },
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'fresh', id: 'fresh_1' },
      ],
      sequential: true,
    });
    expect(seed.success).toBe(true);
    // No list anywhere behind fresh_1: fails closed, never a silent new list.
    const bad = await eng.executeBatch({
      operations: [{ action: 'list', operation: 'continue', target: ['fresh_1'] }],
      sequential: true,
    });
    expect(bad.success).toBe(false);

    // Explicit previous works even when it is not the nearest sibling.
    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    const astB = await parser2.getAST();
    const engB = new OperationEngine(astB, parser2);
    await engB.executeBatch({
      operations: [
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'a', id: 'a_1' },
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'b', id: 'b_1' },
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'gap', id: 'gap_1' },
        { action: 'insert', type: 'paragraph', anchor: 'end', text: 'c', id: 'c_1' },
      ],
      sequential: true,
    });
    await engB.executeBatch({
      operations: [{ action: 'list', operation: 'create', target: ['a_1', 'b_1'], listType: 'bullet' }],
      sequential: true,
    });
    const joined = await engB.executeBatch({
      operations: [{ action: 'list', operation: 'continue', target: ['c_1'], previousNodeId: 'a_1' }],
      sequential: true,
    });
    expect(joined.success).toBe(true);
    const [aNum, , , cNum] = numIdsOf(astB, ['a_1', 'b_1', 'gap_1', 'c_1']);
    expect(cNum).toBe(aNum);
  });

  it('TOON stays truthful after save/reload; remove clears the label', async () => {
    const { parser, ast, eng } = await seededList('numbered');
    const out = parser.rezip(ast);
    const reloaded = await new DocxParser(out).getAST();
    const gen = new ToonGenerator();
    const styles = gen.toon(reloaded, { fidelity: 'standard' }).rows
      .filter((r) => ['first item', 'second item'].includes(r.text))
      .map((r) => r.style);
    expect(styles).toEqual(['list:numbered:L1', 'list:numbered:L1']);

    const removed = await eng.executeBatch({
      operations: [{ action: 'list', operation: 'remove', target: ['li_1'] }],
      sequential: true,
    });
    expect(removed.success).toBe(true);
    const li1 = ast.sections.flatMap((s) => s.nodes).find((n) => n.id === 'li_1')!;
    expect(li1.metadata.listLevel).toBeUndefined();
    expect(li1.metadata.listStyle).toBeUndefined();
  });

  it('executeList rejects unknown operations loudly (never a silent numId 1 stamp)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    expect(() => executeList(
      { action: 'list', operation: 'frobnicate', target: ['whatever'] } as never, ast
    )).toThrow(/Unknown list operation/);
  });
});

describe('D1 BX-04: header/footer match by type AND pageType', () => {
  it('first-page header allocates a fresh part instead of relabeling default', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const hf = new HeaderFooterManipulator(ast, parser);

    const def = hf.setHeaderFooter({ targetType: 'header', pageType: 'default', text: 'Default Header' });
    const first = hf.setHeaderFooter({ targetType: 'header', pageType: 'first', text: 'First Page Header' });
    const foot = hf.setHeaderFooter({ targetType: 'footer', pageType: 'default', text: 'Default Footer' });

    expect(first.rId).not.toBe(def.rId);
    expect(first.partPath).not.toBe(def.partPath);
    expect(new Set([def.rId, first.rId, foot.rId]).size).toBe(3);

    // Re-setting default reuses the default part (no dup).
    const def2 = hf.setHeaderFooter({ targetType: 'header', pageType: 'default', text: 'Default Header v2' });
    expect(def2.rId).toBe(def.rId);

    // sectPr carries both references with correct types and rIds.
    const sectPr = parser.getSectPrElements()[0];
    const refs = Array.from(sectPr.getElementsByTagName('w:headerReference'));
    const byType = new Map(refs.map(r => [(r as Element).getAttribute('w:type'), (r as Element).getAttribute('r:id')]));
    expect(byType.get('default')).toBe(def.rId);
    expect(byType.get('first')).toBe(first.rId);

    // Default part content survived (not relabeled/overwritten).
    const defaultXml = parser.getZipAdapter().readText(def.partPath);
    expect(defaultXml).toContain('Default Header');
    expect(defaultXml).not.toContain('First Page Header');

    // Round-trip: both header types present after reload.
    const out = parser.rezip(ast);
    const parser2 = new DocxParser(out);
    const ast2 = await parser2.getAST();
    const types = new Set(Object.values(ast2.headers).map(h => h.type));
    expect(types.has('default')).toBe(true);
    expect(types.has('first')).toBe(true);
    expect(parser2.validateZipIntegrity().isValid).toBe(true);
  });
});

describe('D1 BX-11: insertRootNode true-parent + recursive registration', () => {
  const SDT_XML =
    `<w:sdt xmlns:w="${W_NS}">` +
    `<w:sdtPr><w:alias w:val="box"/></w:sdtPr>` +
    `<w:sdtContent><w:p><w:r><w:t>Inner sdt paragraph</w:t></w:r></w:p></w:sdtContent>` +
    `</w:sdt>`;

  it('registers sdt children and inserts relative to the true parent', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    await parser.getAST();

    parser.insertRootNode(SDT_XML, undefined, 'append', 'sdt_test_1');
    const summary = parser.getStructuralSummary();
    const inner = summary.nodes.find(n => n.id.startsWith('sdt_test_1__'));
    expect(inner).toBeDefined();
    const innerId = inner!.id;
    expect(parser.readNode(innerId)).toContain('Inner sdt paragraph');

    // Insert after the sdt-wrapped anchor: must land in w:sdtContent, not w:body.
    const afterId = parser.insertRootNode(
      `<w:p xmlns:w="${W_NS}"><w:r><w:t>After inner</w:t></w:r></w:p>`,
      innerId,
      'after',
      'after_inner_1'
    );
    const innerEl = parser.getNode(innerId)!;
    const afterEl = parser.getNode(afterId)!;
    expect(afterEl.parentNode).toBe(innerEl.parentNode);
    expect((afterEl.parentNode as Element).tagName).toBe('w:sdtContent');

    expect(parser.validateZipIntegrity().isValid).toBe(true);
    const out = parser.rezip();
    expect(new DocxParser(out).validateZipIntegrity().isValid).toBe(true);
  });
});

describe('D1 BX-01: comments/PIs survive *Pr schema-order sort', () => {
  it('preserves comment and PI nodes while still sorting elements', () => {
    const xml =
      `<w:rPr xmlns:w="${W_NS}">` +
      `<w:color w:val="FF0000"/><!-- keep-me --><?astra-pin test?>` +
      `<w:sz w:val="32"/><w:rFonts w:ascii="Calibri"/><w:b/>` +
      `</w:rPr>`;
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    const rPr = doc.documentElement;
    sortChildrenBySchema(rPr, RPR_CHILD_ORDER);

    const childNames = Array.from(rPr.childNodes)
      .filter((n: any) => n.nodeType === 1)
      .map((n: any) => n.tagName);
    expect(childNames).toEqual(['w:rFonts', 'w:b', 'w:color', 'w:sz']);

    const serialized = new XMLSerializer().serializeToString(rPr);
    expect(serialized).toContain('keep-me');
    expect(serialized).toContain('astra-pin');
  });
});

describe('D1 BX-08: TOON revision counter + bounded map', () => {
  it('bumps revision on recompute, hits cache otherwise, and bounds entries', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const gen = new ToonGenerator();

    const rev0 = gen.getRevision();
    const t1 = gen.generateToon(ast);
    expect(gen.getRevision()).toBe(rev0 + 1);
    const t1b = gen.generateToon(ast);
    expect(t1b).toBe(t1);
    expect(gen.getRevision()).toBe(rev0 + 1);

    // Flood with section-scoped keys: cache must stay bounded.
    for (let i = 0; i < 30; i++) {
      gen.generateToon(ast, `ghost_${i}`);
    }
    expect(gen.getCacheStats().entries).toBeLessThanOrEqual(
      gen.getCacheStats().maxEntries + 1
    );
  });

  it('never serves a stale row after an in-place DOM edit without version bump', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const gen = new ToonGenerator();

    const t1 = gen.generateToon(ast);
    const rev1 = gen.getRevision();
    // Mutate live DOM in place, deliberately skipping the version bump.
    const tEl = (ast.sections[0].nodes[0].xml as Element).getElementsByTagName('w:t')[0] as Element;
    tEl.textContent = 'EDITED-IN-PLACE-WITHOUT-VERSION-BUMP';
    const t2 = gen.generateToon(ast);
    expect(t2).not.toBe(t1);
    expect(gen.getRevision()).toBe(rev1 + 1);

    gen.invalidateCache();
    const t3 = gen.generateToon(ast);
    expect(t3).not.toBe(t2);
  });
});

describe('D1 BX-02: cloneAST deep-copies every map', () => {
  it('mutating the clone leaves the original maps, nodes and DOM untouched', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    new NumberingManipulator(ast).getOrCreateList('bullet');
    const rels = new RelationshipsManipulator(ast);
    const imgRId = rels.addRelationship(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      'media/x.png'
    );
    const bulletIds = Object.keys(ast.numbering.concreteNum);
    expect(bulletIds.length).toBeGreaterThan(0);

    const snap = cloneAST(ast);
    snap.relationships[imgRId].target = 'media/MUTATED.png';
    delete snap.numbering.concreteNum[bulletIds[0]];
    snap.settings.defaultAuthor = 'MUTATED';
    snap.metadata.title = 'MUTATED';
    snap.mediaFiles['evil.png'] = Buffer.from([1, 2, 3]);
    snap.sections[0].nodes[0].metadata.text = 'MUTATED';
    (snap.sections[0].nodes[0].xml as Element).setAttribute('data-probe', '1');

    expect(ast.relationships[imgRId].target).toBe('media/x.png');
    expect(ast.numbering.concreteNum[bulletIds[0]]).toBeDefined();
    expect(ast.settings.defaultAuthor).not.toBe('MUTATED');
    expect(ast.metadata.title).not.toBe('MUTATED');
    expect(ast.mediaFiles['evil.png']).toBeUndefined();
    expect(ast.sections[0].nodes[0].metadata.text).not.toBe('MUTATED');
    expect((ast.sections[0].nodes[0].xml as Element).getAttribute('data-probe')).toBeFalsy();
  });
});

describe('D1 BX-10: loud errors for unknown coerce paths (parser/manipulator scope)', () => {
  it('throws on unknown list type and unknown header/footer targets', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    expect(() => new NumberingManipulator(ast).getOrCreateList('roman' as any)).toThrow(
      /Unknown list type/
    );
    const hf = new HeaderFooterManipulator(ast, parser);
    expect(() =>
      hf.setHeaderFooter({ targetType: 'masthead' as any, text: 'x' })
    ).toThrow(/Unknown header\/footer targetType/);
    expect(() =>
      hf.setHeaderFooter({ targetType: 'header' as any, pageType: 'middle' as any, text: 'x' })
    ).toThrow(/Unknown header\/footer pageType/);
    expect(() =>
      hf.removeHeaderFooter({ targetType: 'footer' as any, pageType: 'middle' as any })
    ).toThrow(/Unknown header\/footer pageType/);
  });
});

describe('D1 BX-12: save-path `both` normalization is pinned', () => {
  it('rezip converts w:val="justify" to w:val="both" (no pre-existing save-path test)', () => {
    // NOTE (BX-12 "check existing test first"): the only existing `both`
    // coverage is pdf-export-fidelity-and-headers-footers.test.ts, which bakes
    // w:val="both" into its fixture — it never exercises the save path that
    // normalizes invalid `justify`. This test pins that path.
    const zip = new AdmZip();
    const docXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:pPr><w:jc w:val="justify"/></w:pPr><w:r><w:t>needs norm</w:t></w:r></w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>` +
      `<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
      `</w:body></w:document>`;
    zip.addFile('word/document.xml', Buffer.from(docXml, 'utf8'));
    zip.addFile(
      '[Content_Types].xml',
      Buffer.from(
        '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>',
        'utf8'
      )
    );
    const parser = new DocxParser(zip.toBuffer());
    const out = parser.rezip();
    const saved = new AdmZip(out).getEntry('word/document.xml')!.getData().toString('utf8');
    expect(saved).not.toMatch(/w:val="justify"/);
    expect(saved).toMatch(/w:val="both"/);
  });
});

describe('D1 OX-12 + BX-05: safe bookmark and rId allocation', () => {
  it('repeated TOC scans never duplicate bookmark w:ids', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    ast.sections[0].nodes[0].metadata.namedStyle = 'Heading1';
    ast.sections[0].nodes[1].metadata.namedStyle = 'Heading2';

    TOCManipulator.scanHeadings(ast);
    TOCManipulator.scanHeadings(ast);
    TOCManipulator.scanHeadings(ast);

    const doc = parser.getDocumentDom();
    // bookmarkStart/bookmarkEnd of one pair share their w:id by spec, so
    // uniqueness is asserted per tag (plus start/end pairing below).
    const starts = Array.from(doc.getElementsByTagName('w:bookmarkStart')).map(
      e => (e as Element).getAttribute('w:id') ?? ''
    );
    const ends = Array.from(doc.getElementsByTagName('w:bookmarkEnd')).map(
      e => (e as Element).getAttribute('w:id') ?? ''
    );
    expect(starts.length).toBeGreaterThan(0);
    expect(new Set(starts).size).toBe(starts.length);
    expect(new Set(ends).size).toBe(ends.length);
    expect([...starts].sort()).toEqual([...ends].sort());

    // Allocator stays monotonic past stamped ids.
    const a1 = allocateNextBookmarkId(doc);
    const firstP = doc.getElementsByTagName('w:p')[0] as Element;
    const probe = doc.createElement('w:bookmarkStart');
    probe.setAttribute('w:id', String(a1));
    probe.setAttribute('w:name', '_TocProbe');
    firstP.appendChild(probe);
    expect(allocateNextBookmarkId(doc)).toBe(a1 + 1);

    expect(parser.validateZipIntegrity().isValid).toBe(true);
  });

  it('safe rId allocator never reuses ids, even beside bypass shapes', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    ast.relationships['rId_link_1757000000000'] = {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
      target: 'https://example.com',
      targetMode: 'External',
    };
    ast.relationships['rIdH1'] = {
      type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
      target: 'header1.xml',
    };
    const manip = new RelationshipsManipulator(ast);
    const seen = new Set(Object.keys(ast.relationships));
    for (let i = 0; i < 10; i++) {
      const fresh = manip.generateSafeRId();
      expect(fresh).toMatch(/^rId\d+$/);
      expect(seen.has(fresh)).toBe(false);
      ast.relationships[fresh] = { type: 't', target: `part${i}.xml` };
      seen.add(fresh);
    }
  });

  it('comment part bootstrap allocates fresh rels and validates clean', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    await parser.getAST();
    bootstrapCommentParts(parser.getZipAdapter());
    bootstrapCommentParts(parser.getZipAdapter()); // idempotent re-run
    const rels = parser.getZipAdapter().readText('word/_rels/document.xml.rels');
    const relIds = Array.from(rels.matchAll(/Id="(rId\d+)"/g)).map(m => m[1]);
    expect(new Set(relIds).size).toBe(relIds.length);
    expect(parser.validateZipIntegrity().isValid).toBe(true);
  });

  it('header/footer/comment round-trip validates clean under the strict validator', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const hf = new HeaderFooterManipulator(ast, parser);
    hf.setHeaderFooter({ targetType: 'header', pageType: 'default', text: 'Strict H' });
    hf.setHeaderFooter({ targetType: 'footer', pageType: 'default', text: 'Strict F' });
    bootstrapCommentParts(parser.getZipAdapter());
    const out = parser.rezip(ast);
    const probe = new DocxParser(out);
    const result = await probe.validateOOXML();
    expect(result.ok).toBe(true);
    expect(result.errors).toHaveLength(0);
  });
});

// ── D1 follow-up (brutal review): BX-03 bare-rezip route-around ─────────────
// Production saves (document-manager.ts:91,155,236,259,354,421) all call bare
// rezip(). This pins that in-memory list defs survive the bare path via the
// parser's cached-AST fallback (numbering part + rel + content-type), that
// docs without lists gain no spurious part, and that the cache invalidates on
// reload/restore.
describe('D1 follow-up BX-03: bare rezip() preserves numbering via cached AST', () => {
  it('bare rezip() after list.create writes numbering.xml + rel + content-type', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    expect(parser.hasCachedNumberingAst()).toBe(true);
    const bulletId = new NumberingManipulator(ast).getOrCreateList('bullet');
    stampListNumbering(ast.sections[0].nodes[0].xml as Element, bulletId, 0);

    const out = parser.rezip(); // bare — production path
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')).toBeTruthy();
    const numberingXml = zip.getEntry('word/numbering.xml')!.getData().toString('utf8');
    expect(numberingXml).toContain(`w:numId="${bulletId}"`);
    const rels = zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
    expect(rels).toContain('relationships/numbering');
    expect(zip.getEntry('[Content_Types].xml')!.getData().toString('utf8')).toContain('numbering');

    // Round-trip: reload sees the same numId and continuation resolves.
    const parser2 = new DocxParser(out);
    const ast2 = await parser2.getAST();
    expect(Object.keys(ast2.numbering.concreteNum)).toContain(bulletId);
  });

  it('bare rezip() without lists creates no spurious numbering part', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    await parser.getAST();
    const out = parser.rezip();
    expect(new AdmZip(out).getEntry('word/numbering.xml')).toBeFalsy();
  });

  it('reload/restore clears the numbering cache (no stale replay)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    new NumberingManipulator(ast).getOrCreateList('bullet');
    expect(parser.hasCachedNumberingAst()).toBe(true);
    parser.reloadDocumentXml(parser.readDocumentXml());
    expect(parser.hasCachedNumberingAst()).toBe(false);
    // Bare rezip after reload preserves the reloaded zip (no crash, no stale).
    const out = parser.rezip();
    expect(new AdmZip(out).getEntry('word/document.xml')).toBeTruthy();

    const parser2 = new DocxParser(createBlankDocxBuffer('blank'));
    const ast2 = await parser2.getAST();
    new NumberingManipulator(ast2).getOrCreateList('numbered');
    const snap = parser2.capturePackageSnapshot();
    parser2.restorePackageSnapshot(snap);
    expect(parser2.hasCachedNumberingAst()).toBe(false);
  });

  it('explicit rezip(ast) refreshes the cache for later bare saves', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    await parser.getAST();
    parser.clearCachedNumberingAst();
    expect(parser.hasCachedNumberingAst()).toBe(false);
    const ast2 = await parser.getAST();
    const nid = new NumberingManipulator(ast2).getOrCreateList('numbered');
    stampListNumbering(ast2.sections[0].nodes[0].xml as Element, nid, 0);
    parser.rezip(ast2); // explicit refresh
    expect(parser.hasCachedNumberingAst()).toBe(true);
    const out = parser.rezip(); // bare follows explicit
    expect(new AdmZip(out).getEntry('word/numbering.xml')).toBeTruthy();
  });
});

// ── D1 follow-up (brutal review): BX-08 heading + global signals ────────────
// TOON renders section.heading.text (Header column + breadcrumb path) but the
// old fingerprint ignored it. These pin that heading-only edits (detached
// heading DOM, so node live-text cannot mask the bug) plus settings /
// header-footer / numbering signals invalidate instead of serving stale rows.
describe('D1 follow-up BX-08: heading and global signals invalidate TOON', () => {
  function detachedHeadingAst(ast: any, text: string): void {
    const hDoc = (ast.sections[0].nodes[0].xml as Element).ownerDocument;
    const hEl = hDoc.createElement('w:p');
    const hr = hDoc.createElement('w:r');
    const ht = hDoc.createElement('w:t');
    ht.textContent = text;
    hr.appendChild(ht);
    hEl.appendChild(hr);
    ast.sections[0].heading = {
      id: 'h_detached', level: 1, text, type: 'paragraph', xml: hEl,
      metadata: { text, hasCitation: false, hasComment: false, hasBookmark: false, hasTrackedChanges: false, hasEquation: false, hasField: false, hasImage: false, hasHyperlink: false },
      version: 1, checksum: 'h1',
    };
    ast.sections[0].version = 1;
    ast.sections[0].checksum = 's1';
  }

  it('heading-only edit (detached DOM) recomputes instead of hitting cache', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    detachedHeadingAst(ast, 'Original Heading');
    const gen = new ToonGenerator();
    const t1 = gen.generateToon(ast);
    expect(t1.rows[0]?.header).toBe('Original Heading');
    const rev1 = gen.getRevision();
    ast.sections[0].heading.text = 'EDITED HEADING ONLY';
    (ast.sections[0].heading.xml as Element).textContent = 'EDITED HEADING ONLY';
    const t2 = gen.generateToon(ast);
    expect(t2).not.toBe(t1);
    expect(gen.getRevision()).toBe(rev1 + 1);
    expect(t2.rows[0]?.header).toBe('EDITED HEADING ONLY');
  });

  it('settings / header-footer / numbering edits invalidate the cache', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const gen = new ToonGenerator();

    const t0 = gen.generateToon(ast);
    ast.settings.defaultAuthor = 'MUTATED-AUTHOR';
    expect(gen.generateToon(ast)).not.toBe(t0);

    const t1 = gen.generateToon(ast);
    ast.headers['rId99'] = { type: 'default', xml: ast.sections[0].nodes[0].xml as Element, text: 'H-text' };
    expect(gen.generateToon(ast)).not.toBe(t1);

    const t2 = gen.generateToon(ast);
    ast.footers['rId98'] = { type: 'default', xml: ast.sections[0].nodes[0].xml as Element, text: 'F-text' };
    expect(gen.generateToon(ast)).not.toBe(t2);

    const t3 = gen.generateToon(ast);
    new NumberingManipulator(ast).getOrCreateList('bullet');
    expect(gen.generateToon(ast)).not.toBe(t3);
  });
});

// ── D1 follow-up (brutal review): BX-10 remainders (owning-batch spec) ──────
// FIXED (D2): insert.ts + trackedChange.ts went loud with received value +
// valid set + hint (see ops-d2-fix-verify.test.ts). Pinned here so a
// regression back to generic messages is caught.
describe('D1 follow-up BX-10: remainders verified + owned loud errors pinned', () => {
  it('verifies insert + trackedChange loud errors (received value + valid set + hint)', async () => {
    const { default: fs } = await import('fs');
    const insSrc = fs.readFileSync('packages/engine/src/operations/insert.ts', 'utf8');
    expect(insSrc).toContain('Unsupported insert type');
    expect(insSrc).toContain('Supported:');
    expect(insSrc).toContain('inspect_document');
    // D2 (BX-10, engine-owned): operation-engine.ts:286 went loud — the bare
    // `Unsupported operation: ${op.action}` is GONE, replaced by the Unknown
    // operation error carrying the received value + valid set + hint.
    const engSrc = fs.readFileSync('packages/engine/src/operation-engine.ts', 'utf8');
    expect(engSrc).not.toContain('throw new Error(`Unsupported operation: ${op.action}`);');
    expect(engSrc).toContain("Unknown operation '");
    expect(engSrc).toContain('inspect_document');
    const tcSrc = fs.readFileSync('packages/engine/src/operations/trackedChange.ts', 'utf8');
    expect(tcSrc).toContain('Unsupported trackedChange operation');
    expect(tcSrc).toContain('Supported:');
    expect(tcSrc).toContain('inspect_document');
  });

  it('owned scope stays loud: unknown list/header-footer types fail with hints', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    expect(() => new NumberingManipulator(ast).getOrCreateList('roman' as any)).toThrow(/Unknown list type.*bullet.*numbered/);
    const hf = new HeaderFooterManipulator(ast, parser);
    expect(() => hf.setHeaderFooter({ targetType: 'masthead' as any, text: 'x' })).toThrow(/Unknown header\/footer targetType/);
    expect(() => hf.setHeaderFooter({ targetType: 'header' as any, pageType: 'middle' as any, text: 'x' })).toThrow(/Unknown header\/footer pageType/);
  });

  it('remainder runtime behavior: loud messages carry received value + valid set + hint', async () => {
    const { executeInsert } = await import('@attentionform/astra-engine');
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await expect(
      executeInsert({ action: 'insert', type: 'bogus-type' as any, anchor, position: 'after' } as any, ast, parser)
    ).rejects.toThrow(/Unsupported insert type 'bogus-type'.*Supported:.*inspect_document/);
    const { executeTrackedChange } = await import('@attentionform/astra-engine');
    expect(() => executeTrackedChange({ action: 'trackedChange', operation: 'bogus' as any } as any, ast, parser)).toThrow(/Unsupported trackedChange operation: bogus.*Supported:.*inspect_document/);
  });
});

// ── D1R2 OX-12: owning-batch adoption (hyperlink.ts fixed) ─────────────────
// hyperlink.ts now allocates via allocateSafeHyperlinkRId (canonical rIdN,
// ast.relationships write BEFORE the zip) and allocateNextBookmarkId (unique
// w:id shared by each start/end pair, never the hardcoded 0). The old bypass
// (same-ms shared rId_link_<ts>, duplicated w:id 0) is pinned GONE below;
// allocator route-around coverage is retained as the unit contract.
describe('D1R2 OX-12: hyperlink bypass fixed via owning-batch allocators', () => {
  it('same-ms hyperlinks mint unique canonical rIds, AST-first then zip', async () => {
    resetBookmarkHighWaterForTests();
    resetHyperlinkRIdStateForTests();
    try {
      const { executeHyperlink } = await import('@attentionform/astra-engine');
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast: any = await parser.getAST();
      const target = ast.sections[0].nodes[0].id;
      const realNow = Date.now;
      (Date as any).now = () => 1757000000000;
      try {
        (executeHyperlink as any)({ action: 'hyperlink', target, url: 'https://a.example', text: 'A' }, ast, parser);
        (executeHyperlink as any)({ action: 'hyperlink', target, url: 'https://b.example', text: 'B' }, ast, parser);
      } finally {
        (Date as any).now = realNow;
      }
      // No timestamp ids remain; both rIds are canonical and distinct.
      const relsXml = parser.getZipAdapter().readText('word/_rels/document.xml.rels');
      expect(relsXml).not.toMatch(/rId_link_/);
      const linkTargets = Array.from(relsXml.matchAll(/<Relationship[^>]*Type="[^"]*hyperlink"[^>]*>/g));
      expect(linkTargets).toHaveLength(2);
      const zipIds = Array.from(relsXml.matchAll(/Id="([^"]+)"/g)).map(m => m[1]);
      expect(new Set(zipIds).size).toBe(zipIds.length);
      // AST-first: the map learned both rels (External) before any rezip.
      const hyperlinkRIds = Object.keys(ast.relationships).filter(
        rId => ast.relationships[rId].type ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
      );
      expect(hyperlinkRIds).toHaveLength(2);
      expect(new Set(hyperlinkRIds).size).toBe(2);
      for (const rId of hyperlinkRIds) {
        expect(rId).toMatch(/^rId\d+$/);
        expect(ast.relationships[rId].targetMode).toBe('External');
      }
      expect(zipIds).toEqual(expect.arrayContaining(hyperlinkRIds));
      // Explicit rezip keeps both rels (AST→zip merge is consistent, no dupes).
      const out = parser.rezip(ast);
      const rels2 = new AdmZip(out).getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8');
      const zipIds2 = Array.from(rels2.matchAll(/Id="([^"]+)"/g)).map(m => m[1]);
      expect(new Set(zipIds2).size).toBe(zipIds2.length);
      expect(rels2).toContain('https://a.example');
      expect(rels2).toContain('https://b.example');
      expect(parser.validateZipIntegrity().isValid).toBe(true);
    } finally {
      resetBookmarkHighWaterForTests();
      resetHyperlinkRIdStateForTests();
    }
  });

  it('same-AST-seeded allocators still diverge across documents (global reservation)', async () => {
    resetHyperlinkRIdStateForTests();
    try {
      const parserA = new DocxParser(createBlankDocxBuffer('blank'));
      const astA = await parserA.getAST();
      const parserB = new DocxParser(createBlankDocxBuffer('blank'));
      const astB = await parserB.getAST();
      const realNow = Date.now;
      (Date as any).now = () => 1757000000000;
      try {
        const h1 = allocateSafeHyperlinkRId(astA);
        astA.relationships[h1] = { type: 'hyperlink', target: 'https://a.example', targetMode: 'External' };
        // astB predates that write, yet the reservation forces divergence.
        const h2 = allocateSafeHyperlinkRId(astB);
        expect(h2).not.toBe(h1);
        expect(h2).toMatch(/^rId\d+$/);
      } finally {
        (Date as any).now = realNow;
      }
    } finally {
      resetHyperlinkRIdStateForTests();
    }
  });

  it('same-ms default bookmarks mint unique paired w:ids, never 0', async () => {
    resetBookmarkHighWaterForTests();
    try {
      const { executeHyperlink } = await import('@attentionform/astra-engine');
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast: any = await parser.getAST();
      const realNow = Date.now;
      (Date as any).now = () => 1757000000000;
      try {
        (executeHyperlink as any)({ action: 'bookmark', target: ast.sections[0].nodes[0].id }, ast, parser);
        (executeHyperlink as any)({ action: 'bookmark', target: ast.sections[0].nodes[1].id }, ast, parser);
      } finally {
        (Date as any).now = realNow;
      }
      const starts = Array.from(parser.getDocumentDom().getElementsByTagName('w:bookmarkStart')).map(
        e => (e as Element).getAttribute('w:id') ?? ''
      );
      const ends = Array.from(parser.getDocumentDom().getElementsByTagName('w:bookmarkEnd')).map(
        e => (e as Element).getAttribute('w:id') ?? ''
      );
      expect(starts).toHaveLength(2);
      expect(new Set(starts).size).toBe(2); // unique — old path duplicated 0
      expect(starts).not.toContain('0'); // hardcoded bypass value never minted
      expect([...starts].sort()).toEqual([...ends].sort()); // start/end pairs share id
      const names = Array.from(parser.getDocumentDom().getElementsByTagName('w:bookmarkStart')).map(
        e => (e as Element).getAttribute('w:name') ?? ''
      );
      expect(new Set(names).size).toBe(2); // default names unique despite same ms
      expect(parser.validateZipIntegrity().isValid).toBe(true);
    } finally {
      resetBookmarkHighWaterForTests();
    }
  });

  it('safe allocators route around: canonical rIds never reuse, bookmarks monotonic', async () => {
    resetBookmarkHighWaterForTests();
    resetHyperlinkRIdStateForTests();
    try {
      const parser = new DocxParser(createBlankDocxBuffer('blank'));
      const ast = await parser.getAST();
      ast.relationships['rId_link_1757000000000'] = {
        type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
        target: 'https://example.com', targetMode: 'External',
      };
      const manip = new RelationshipsManipulator(ast);
      const seen = new Set(Object.keys(ast.relationships));
      for (let i = 0; i < 10; i++) {
        const fresh = manip.generateSafeRId();
        expect(fresh).toMatch(/^rId\d+$/);
        expect(seen.has(fresh)).toBe(false);
        ast.relationships[fresh] = { type: 't', target: `part${i}.xml` };
        seen.add(fresh);
      }
      // Atomic hyperlink helper: two same-ms allocations stay unique + canonical.
      const realNow = Date.now;
      (Date as any).now = () => 1757000000000;
      try {
        const h1 = allocateSafeHyperlinkRId(ast);
        ast.relationships[h1] = { type: 'hyperlink', target: 'https://a.example', targetMode: 'External' };
        const h2 = allocateSafeHyperlinkRId(ast);
        expect(h2).not.toBe(h1);
        expect(h2).toMatch(/^rId\d+$/);
      } finally {
        (Date as any).now = realNow;
      }
      // Bookmark allocator: pre-insert double scan still mints distinct ids.
      const doc = parser.getDocumentDom();
      const b1 = allocateNextBookmarkId(doc);
      const b2 = allocateNextBookmarkId(doc); // no DOM insert between scans
      expect(b2).not.toBe(b1);
      expect(b2).toBeGreaterThan(b1);
    } finally {
      resetBookmarkHighWaterForTests();
      resetHyperlinkRIdStateForTests();
    }
  });
});

// ── D1R2-1: orphan prune (concrete deleted + abstract remains) ──────────────
// rezip() persists the numbering part ONLY when at least one concrete num
// references an existing abstract; otherwise it PRUNES (drops the part AND
// its rel AND its content-type override, including stale entries written by
// an earlier save) instead of persisting an orphan part. Gate lives in
// parser.ts (numbering.ts is owned by another batch and stays untouched).
describe('D1R2-1: orphan numbering is pruned, never persisted', () => {
  it('concrete deleted + abstract remains → no part, no rel, no content-type (explicit rezip)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const bulletId = new NumberingManipulator(ast).getOrCreateList('bullet');
    delete ast.numbering.concreteNum[bulletId]; // orphan: abstract remains

    const out = parser.rezip(ast);
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')).toBeFalsy();
    expect(zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8')).not.toContain('numbering');
    expect(zip.getEntry('[Content_Types].xml')!.getData().toString('utf8')).not.toContain('numbering');
    expect(parser.validateZipIntegrity().isValid).toBe(true);
  });

  it('concrete deleted + abstract remains → bare rezip prunes too', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const bulletId = new NumberingManipulator(ast).getOrCreateList('bullet');
    delete ast.numbering.concreteNum[bulletId];

    const out = parser.rezip(); // bare — production path
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')).toBeFalsy();
    expect(zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8')).not.toContain('numbering');
    expect(zip.getEntry('[Content_Types].xml')!.getData().toString('utf8')).not.toContain('numbering');
  });

  it('abstract deleted + concrete dangles → pruned (mirror orphan)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    new NumberingManipulator(ast).getOrCreateList('numbered'); // numId 2 → abs 1
    for (const k of Object.keys(ast.numbering.abstractNum)) delete ast.numbering.abstractNum[k];
    expect(Object.keys(ast.numbering.concreteNum).length).toBeGreaterThan(0);

    const out = parser.rezip(ast);
    expect(new AdmZip(out).getEntry('word/numbering.xml')).toBeFalsy();
  });

  it('previously saved part is REMOVED when numbering later becomes orphan', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const bulletId = new NumberingManipulator(ast).getOrCreateList('bullet');
    stampListNumbering(ast.sections[0].nodes[0].xml as Element, bulletId, 0);
    parser.rezip(ast); // part + rel + CT now in the zip
    expect(parser.getZipAdapter().readTextOrNull('word/numbering.xml')).toBeTruthy();

    delete ast.numbering.concreteNum[bulletId];
    const out = parser.rezip(ast);
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')).toBeFalsy();
    expect(zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8')).not.toContain('numbering');
    expect(zip.getEntry('[Content_Types].xml')!.getData().toString('utf8')).not.toContain('numbering');

    // Reload sees no numbering (no zombie rehydration) and stays valid.
    const parser2 = new DocxParser(out);
    const ast2 = await parser2.getAST();
    expect(Object.keys(ast2.numbering.concreteNum)).toHaveLength(0);
    expect(parser2.validateZipIntegrity().isValid).toBe(true);
  });

  it('live numbering still writes part + rel + content-type (no over-prune)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const bulletId = new NumberingManipulator(ast).getOrCreateList('bullet');
    const out = parser.rezip(ast);
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')!.getData().toString('utf8')).toContain(`w:numId="${bulletId}"`);
    expect(zip.getEntry('word/_rels/document.xml.rels')!.getData().toString('utf8')).toContain('numbering');
    expect(zip.getEntry('[Content_Types].xml')!.getData().toString('utf8')).toContain('numbering');
  });
});

// ── D1R2-2: second-getAST preserves unsaved numbering defs ─────────────────
// getAST() carries over cached abstract/concrete entries the fresh zip
// hydration missed, so create-list → getAST() again → bare rezip keeps defs.
describe('D1R2-2: getAST preserves unsaved numbering defs across calls', () => {
  it('create list → getAST() again → bare rezip keeps defs', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast1 = await parser.getAST();
    const bulletId = new NumberingManipulator(ast1).getOrCreateList('bullet');
    stampListNumbering(ast1.sections[0].nodes[0].xml as Element, bulletId, 0);

    const ast2 = await parser.getAST(); // must NOT drop ast1's unsaved defs
    expect(Object.keys(ast2.numbering.concreteNum)).toContain(bulletId);
    expect(Object.keys(ast2.numbering.abstractNum).length).toBeGreaterThan(0);

    const out = parser.rezip(); // bare — production path follows the 2nd getAST
    const zip = new AdmZip(out);
    expect(zip.getEntry('word/numbering.xml')!.getData().toString('utf8')).toContain(`w:numId="${bulletId}"`);

    const parser2 = new DocxParser(out);
    const ast3 = await parser2.getAST();
    expect(Object.keys(ast3.numbering.concreteNum)).toContain(bulletId);
    expect(new NumberingManipulator(ast3).continueList(
      ast3.sections[0].nodes.find(n => n.metadata.text === ast1.sections[0].nodes[0].metadata.text)!.id
    )).toBe(bulletId);
  });

  it('reload still discards unsaved defs (cache invalidation wins over carry-over)', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast1 = await parser.getAST();
    new NumberingManipulator(ast1).getOrCreateList('bullet');
    parser.reloadDocumentXml(parser.readDocumentXml()); // clears cache
    const ast2 = await parser.getAST();
    expect(Object.keys(ast2.numbering.concreteNum)).toHaveLength(0);
    expect(new AdmZip(parser.rezip()).getEntry('word/numbering.xml')).toBeFalsy();
  });
});

// ── D1R2-3: rollback divergence (engine getAST-fresh vs this.ast split) ─────
// operation-engine.ts:182-193 is READ-ONLY for this batch. Proven below by
// simulating those exact lines: after a failed batch the parser cache points
// at `fresh` while the engine keeps its own `this.ast`, so defs minted on the
// engine island after any further getAST() are invisible to bare rezip().
// Parser-side mitigation (D1R2-2 carry-over) covers defs that predate the
// split; the residual window needs the engine one-liner
// `this.parser.adoptAstForNumbering(this.ast)` (see
// DocxParser.adoptAstForNumbering docs for the exact spec) — proven here via
// the hook with the engine file untouched.
describe('D1R2-3: rollback identity split proven; adopt hook restores visibility', () => {
  function mirrorEngineRollbackInPlace(engineAst: any, fresh: any): void {
    // Mirrors operation-engine.ts:186-198 (splice + map re-key + metadata).
    engineAst.sections.splice(0, engineAst.sections.length, ...fresh.sections);
    const mapKeys = ['styles', 'numbering', 'relationships', 'headers', 'footers', 'mediaFiles'] as const;
    for (const key of mapKeys) {
      const target = engineAst[key] as Record<string, unknown>;
      for (const k of Object.keys(target)) delete target[k];
      Object.assign(target, fresh[key] as Record<string, unknown>);
    }
    const metadataTarget = engineAst.metadata as Record<string, unknown>;
    for (const k of Object.keys(metadataTarget)) delete metadataTarget[k];
    Object.assign(metadataTarget, fresh.metadata);
    engineAst.nextChangeId = fresh.nextChangeId;
  }

  it('split proven: engine island diverges from parser cache; adopt() fixes bare rezip', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const engineAst: any = await parser.getAST();
    const pkgSnap = parser.capturePackageSnapshot();
    // Failed batch → restoreAfterFailedBatch (engine lines, read-only mirror):
    parser.restorePackageSnapshot(pkgSnap);
    const fresh: any = await parser.getAST();
    mirrorEngineRollbackInPlace(engineAst, fresh);
    expect(engineAst).not.toBe(fresh); // the identity split

    // A further getAST() (any other consumer) strands the engine island.
    await parser.getAST();
    const strandedId = new NumberingManipulator(engineAst).getOrCreateList('numbered');
    const lost = parser.rezip();
    expect(new AdmZip(lost).getEntry('word/numbering.xml')).toBeFalsy(); // divergence proven

    // The spec'd engine one-liner, applied via the parser-side hook:
    parser.adoptAstForNumbering(engineAst);
    const keptId = new NumberingManipulator(engineAst).getOrCreateList('bullet');
    const kept = parser.rezip();
    const keptXml = new AdmZip(kept).getEntry('word/numbering.xml')!.getData().toString('utf8');
    expect(keptXml).toContain(`w:numId="${keptId}"`);
    expect(keptXml).toContain(`w:numId="${strandedId}"`);
    expect(parser.validateZipIntegrity().isValid).toBe(true);
  });

  it('engine file fixed by D2: adoptAstForNumbering present after nextChangeId sync', async () => {
    const { default: fs } = await import('fs');
    const src = fs.readFileSync('packages/engine/src/operation-engine.ts', 'utf8');
    const refresh = fs.readFileSync('packages/engine/src/ast-refresh.ts', 'utf8');
    // The rollback goes through the shared refresh (one copy of the recipe)...
    expect(src).toContain('await refreshAstInPlace(this.ast, this.parser');
    // D2 (OX-06/D1R2-3): ...and the shared refresh re-points the numbering
    // cache at the live object, so bare rezip() follows it after a rollback.
    expect(refresh).toContain('const fresh = await parser.getAST();');
    expect(refresh).toContain('parser.adoptAstForNumbering(target);');
  });
});

describe('numbering fonts and fallback fidelity', () => {
  it('bullet level fonts carry ascii, hAnsi and cs Arial', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    manip.getOrCreateList('bullet');
    const numberingXml = manip.serializeNumbering() ?? '';
    expect(numberingXml).toContain('w:ascii="Arial"');
    expect(numberingXml).toContain('w:hAnsi="Arial"');
    expect(numberingXml).toContain('w:cs="Arial"');
    const out = parser.rezip(ast);
    const saved = new AdmZip(out).getEntry('word/numbering.xml')!.getData().toString('utf8');
    expect(saved).toContain('w:cs="Arial"');
  });

  it('unserializable numbered abstract falls back to decimal, never bullet', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    manip.getOrCreateList('numbered');
    for (const k of Object.keys(ast.numbering.abstractNum)) {
      (ast.numbering.abstractNum as any)[k] = {
        getElementsByTagName: (t: string) =>
          t === 'w:numFmt' ? [{ getAttribute: () => 'decimal' }] : [],
      };
    }
    const fallback = manip.serializeNumbering() ?? '';
    expect(fallback).toContain('w:val="decimal"');
    expect(fallback).toContain('%1.');
    expect(fallback).not.toContain('w:val="bullet"');
  });

  it('unreadable numFmt fails closed instead of demoting to bullet', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    manip.getOrCreateList('numbered');
    for (const k of Object.keys(ast.numbering.abstractNum)) {
      (ast.numbering.abstractNum as any)[k] = {
        getElementsByTagName: () => { throw new Error('boom'); },
      };
    }
    expect(() => manip.serializeNumbering()).toThrow(/numbering abstract/);
    expect(() => manip.serializeNumbering()).toThrow(/numFmt is unreadable/);
    expect(() => manip.serializeNumbering()).toThrow(/inspect_document/);
  });

  it('empty numFmt fails closed instead of demoting to bullet', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const manip = new NumberingManipulator(ast);
    manip.getOrCreateList('numbered');
    for (const k of Object.keys(ast.numbering.abstractNum)) {
      (ast.numbering.abstractNum as any)[k] = {
        getElementsByTagName: () => [],
      };
    }
    expect(() => manip.serializeNumbering()).toThrow(/numbering abstract/);
    expect(() => manip.serializeNumbering()).toThrow(/numFmt is missing or empty/);
    expect(() => manip.serializeNumbering()).toThrow(/inspect_document/);
  });
});
