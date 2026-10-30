/**
 * Nested-table census pins for `computeDocStats` / `tableCellTextSpaced`
 * (packages/engine/src/stats.ts — single source of truth for CLI `stats` and
 * MCP `get_stats`).
 *
 * Regression: `tableCellTextSpaced` iterated every `w:tc` with
 * `getElementsByTagName('w:t')`, so a nested table's runs were re-counted once
 * per ancestor cell (deep fixture wire: 34 words / 1 table vs the raw
 * document.xml census 26 / 3). The document XML below is shape-identical to
 * the fidelity corpus builders in tests/harness/fidelity/generate-fixtures.ts
 * (nested-table-deep, nested-table-single, merged-cells-basic); fixture .docx
 * binaries are gitignored, so the XML is pinned inline here.
 *
 * The oracle is deliberately independent of the engine: it reads
 * word/document.xml raw, counts `<w:tbl` occurrences for tables, and walks the
 * DOM once per element for words (separator at paragraph/cell boundaries, so
 * edge words never fuse). The flat pin also re-runs the pre-fix flat algorithm
 * byte-for-byte to prove the recipe contract is unchanged where no nesting
 * exists.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import { DocxParser } from '../src/parser';
import { computeDocStats, tableCellTextSpaced } from '../src/stats';
import type { Node } from '../src/ast';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const para = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
/** Extra drawing namespaces so the textbox fixture below is shape-valid-ish. */
const DRAWING_NS =
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';
const docWrap = (inner: string): string =>
  `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:document xmlns:w="${W}"${DRAWING_NS}><w:body>${inner}<w:sectPr/></w:body></w:document>`;

/** Minimal single-part package (same shape the corpus generator writes). */
function makeDocx(inner: string): Buffer {
  const zip = new AdmZip();
  zip.addFile(
    '[Content_Types].xml',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
    ),
  );
  zip.addFile('word/document.xml', Buffer.from(docWrap(inner)));
  zip.addFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>',
    ),
  );
  zip.addFile(
    '_rels/.rels',
    Buffer.from(
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
    ),
  );
  return zip.toBuffer();
}

// --- fixture XML, copied from tests/harness/fidelity/generate-fixtures.ts ---

const LEAD_DEEP = [
  'A first plain sentence stands before the deep nesting.',
  'A second plain sentence gives the mutation runner its pair.',
];
const DEEP =
  para(LEAD_DEEP[0]) +
  para(LEAD_DEEP[1]) +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level zero</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="8000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level one</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="7000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>level two leaf</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
  `</w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl>`;

const LEAD_SINGLE = [
  'A first plain sentence stands before the nested table.',
  'A second plain sentence gives the mutation runner its pair.',
];
const SINGLE =
  para(LEAD_SINGLE[0]) +
  para(LEAD_SINGLE[1]) +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:p><w:r><w:t>outer cell opens</w:t></w:r></w:p>` +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:p><w:r><w:t>inner A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>inner B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
  `<w:p><w:r><w:t>outer cell closes</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

/** Reviewer KILL-1 shape: a table living in a textbox (w:txbxContent) whose
 *  drawing is anchored in a paragraph of an outer table cell. The host w:p
 *  must not grab the inner table's runs (they belong to the inner w:tc). */
const TEXTBOX_INNER =
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:p><w:r><w:t>inner words here</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
const TEXTBOX =
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>` +
  `<w:p><w:r><w:t>host open</w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="5" name="TextBox 5"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>${TEXTBOX_INNER}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
  `<w:p><w:r><w:t>host tail</w:t></w:r></w:p>` +
  `</w:tc></w:tr></w:tbl>`;

/** One inline drawing run whose textbox holds `inner` paragraph XML. */
const drawing = (inner: string, id: number): string =>
  `<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="${id}" name="TextBox ${id}"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>${inner}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;

/** Textbox WITHOUT a nested table inside a cell paragraph: the textbox's own
 *  w:p entries count once as part of the cell text, and the host paragraph's
 *  w:t sweep must not also grab their runs (regression: engine 14 words vs the
 *  txbx-aware raw-walk oracle 10 — each textbox run was swept twice). */
const TEXTBOX_NO_TABLE =
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>` +
  `<w:p><w:r><w:t>host open</w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="7" name="TextBox 7"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>${para('box words here')}${para('box second line')}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>` +
  `<w:p><w:r><w:t>host tail</w:t></w:r></w:p>` +
  `</w:tc></w:tr></w:tbl>`;

/** Host runs surround the drawing: host run texts fuse into the host part
 *  ("alphabeta"), the textbox paragraph is the following part. */
const TEXTBOX_RUNS_AROUND =
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>` +
  `<w:p><w:r><w:t>alpha</w:t></w:r>${drawing(para('gamma'), 9)}<w:r><w:t>beta</w:t></w:r></w:p>` +
  `</w:tc></w:tr></w:tbl>`;

/** Two textboxes in one host paragraph: each textbox paragraph is its own part. */
const TEXTBOX_TWO =
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc>` +
  `<w:p><w:r><w:t>aa</w:t></w:r>${drawing(para('bb'), 10)}${drawing(para('cc'), 11)}<w:r><w:t>dd</w:t></w:r></w:p>` +
  `</w:tc></w:tr></w:tbl>`;

/** Same textbox shape at body level (no enclosing cell): body paragraphs keep
 *  the parser's metadata.text capture unchanged by the census fix. */
const TEXTBOX_BODY_LEVEL =
  `<w:p><w:r><w:t xml:space="preserve">body lead </w:t></w:r><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="914400" cy="914400"/><wp:docPr id="8" name="TextBox 8"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:txbx><w:txbxContent>${para('box body words')}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;

const LEAD_FLAT = [
  'A first plain sentence stands before the merged cell tangle.',
  'A second plain sentence gives the mutation runner its pair.',
];
const FLAT =
  para(LEAD_FLAT[0]) +
  para(LEAD_FLAT[1]) +
  `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>` +
  `<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>wide cell</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>side cell</w:t></w:r></w:p></w:tc></w:tr>` +
  `<w:tr><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>merge top</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>middle</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>edge</w:t></w:r></w:p></w:tc></w:tr>` +
  `<w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t/></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>lower middle</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>lower edge</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;

// --- independent oracle (raw document.xml, no engine code) ---

function rawDocumentXml(buf: Buffer): string {
  return new AdmZip(buf).getEntry('word/document.xml')?.getData().toString('utf8') ?? '';
}

/** Counts every w:tbl in the body document.xml (nested included). */
function oracleTables(rawXml: string): number {
  return (rawXml.match(/<w:tbl[ >]/g) ?? []).length;
}

/**
 * Independent oracle over the raw document.xml (no engine code, no recipe
 * lookup): each visible w:t belongs to exactly ONE paragraph — its nearest
 * w:p ancestor — whose text concatenates those run texts (run fusion is the
 * documented recipe, pinned by the "alphabeta" flat-cell case); paragraphs
 * and table cells are separate space-joined parts. A textbox paragraph is a
 * paragraph like any other, so its text joins the host cell as its own part
 * instead of being pulled into the host paragraph or dropped. Fuses or
 * splits exactly as that model dictates, notably: host runs around a drawing
 * stay in the host paragraph ("alphabeta"), the textbox paragraph is the
 * following part ("gamma").
 */
function oracleWords(rawXml: string): number {
  const doc = new DOMParser().parseFromString(rawXml, 'text/xml');
  const body = doc.getElementsByTagName('w:body')[0] as unknown as Element | undefined;
  if (!body) return 0;
  let text = '';
  const children = (el: Element): ArrayLike<Element> =>
    (el.childNodes as unknown as { length: number; [i: number]: Element });
  const direct = (el: Element, tag: string): Element[] => {
    const out: Element[] = [];
    const kids = children(el);
    for (let i = 0; i < kids.length; i++) {
      const c = kids[i] as unknown as Element;
      if (c && c.nodeType === 1 && c.tagName === tag) out.push(c);
    }
    return out;
  };
  const nearestP = (el: Element): Element | null => {
    let cur = el.parentNode as unknown as Element | null;
    while (cur) {
      if (cur.nodeType === 1 && cur.tagName === 'w:p') return cur;
      cur = cur.parentNode as unknown as Element | null;
    }
    return null;
  };
  /** Paragraph text = w:t run texts whose nearest w:p owner is this paragraph. */
  const ownParaText = (p: Element): string => {
    const ts = Array.from(p.getElementsByTagName('w:t')) as unknown as Element[];
    let out = '';
    for (const t of ts) if (nearestP(t) === p) out += t.textContent || '';
    return out;
  };
  const walk = (el: Element): void => {
    const kids = children(el);
    for (let i = 0; i < kids.length; i++) {
      const c = kids[i] as unknown as Element;
      if (!c || c.nodeType !== 1) continue;
      const tag = c.tagName;
      if (tag === 'w:p') {
        text += ownParaText(c) + ' ';
        walk(c);
      } else if (tag === 'w:tbl') {
        for (const tr of direct(c, 'w:tr')) {
          for (const tc of direct(tr, 'w:tc')) {
            walk(tc);
            text += ' ';
          }
        }
      } else {
        walk(c);
      }
    }
  };
  walk(body as unknown as Element);
  return (text.match(/\S+/g) ?? []).length;
}

/** Pre-fix algorithm, algorithm-faithful (HEAD packages/engine/src/stats.ts):
 *  used to prove flat documents keep identical counts after the nested fix. */
function legacyFlatSpaced(xml: Element): string {
  const tcs = xml.getElementsByTagName('w:tc');
  const cells: string[] = [];
  for (let i = 0; i < tcs.length; i++) {
    const tc = tcs[i] as unknown as Element;
    const paras = tc.getElementsByTagName('w:p');
    if (paras.length > 0) {
      const parts: string[] = [];
      for (let p = 0; p < paras.length; p++) {
        const runs = (paras[p] as unknown as Element).getElementsByTagName('w:t');
        let t = '';
        for (let r = 0; r < runs.length; r++) t += (runs[r] as unknown as Element).textContent || '';
        if (t.trim().length > 0) parts.push(t.trim());
      }
      if (parts.length > 0) cells.push(parts.join(' '));
    } else {
      const runs = tc.getElementsByTagName('w:t');
      let t = '';
      for (let r = 0; r < runs.length; r++) t += (runs[r] as unknown as Element).textContent || '';
      if (t.trim().length > 0) cells.push(t.trim());
    }
  }
  return cells.join(' ');
}

async function statsOf(buf: Buffer) {
  const ast = await new DocxParser(buf).getAST();
  return { ast, stats: computeDocStats(ast) };
}

function tableNode(ast: { sections: Array<{ nodes: Node[] }> }): Node {
  const tbl = ast.sections.flatMap((s) => s.nodes).find((n) => n.type === 'table');
  expect(tbl, 'table node present').toBeTruthy();
  return tbl as Node;
}

describe('nested-table census (CLI stats / MCP get_stats shared counters)', () => {
  it('deep fixture: every w:t counted once, nested w:tbl tallied (26 words / 3 tables)', async () => {
    const buf = makeDocx(DEEP);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    // Wire pin: the exact numbers the bug report contrasted with the oracle.
    expect(stats).toMatchObject({ words: 26, paragraphs: 2, tables: 3, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(stats.tables).toBe(oracleTables(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('level zero level one level two leaf');
  });

  it('single-nest fixture: 29 words / 2 tables (was 33 / 1 with per-ancestor repetition)', async () => {
    const buf = makeDocx(SINGLE);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 29, paragraphs: 2, tables: 2, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(stats.tables).toBe(oracleTables(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('outer cell opens outer cell closes inner A inner B');
  });

  it('textbox-nested table in a cell paragraph: host runs counted once (7 words / 2 tables)', async () => {
    const buf = makeDocx(TEXTBOX);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    // Reviewer probe pin (KILL-1): at HEAD (pre-fix) the host paragraph
    // grabbed the textbox table's runs on top of the textbox w:tc iteration
    // — 12 words. The 9-word probe was an intermediate partial fix
    // (paragraph-level owner filter only); the run-level filter lands 7.
    // The txbx-aware oracle also counts 7, but it visits the textbox table's
    // cell inline while the engine recipe orders the host cell's paragraphs
    // before it, so the recipe string is pinned directly instead.
    expect(stats).toMatchObject({ words: 7, paragraphs: 0, tables: 2, images: 0 });
    expect(stats.tables).toBe(oracleTables(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('host open host tail inner words here');
  });

  it('textbox-without-table in a cell paragraph: textbox text counted once (10 words)', async () => {
    const buf = makeDocx(TEXTBOX_NO_TABLE);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    // Pre-fix wire: 14 words — the host paragraph sweep grabbed the textbox
    // runs and the textbox w:p entries were swept again. Each w:t now has
    // exactly one nearest w:p owner, so the same text lands once.
    expect(stats).toMatchObject({ words: 10, paragraphs: 0, tables: 1, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(stats.tables).toBe(oracleTables(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('host open box words here box second line host tail');
  });

  it('host runs around a drawing fuse in the host part; textbox is the next part (2 words)', async () => {
    const buf = makeDocx(TEXTBOX_RUNS_AROUND);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 2, paragraphs: 0, tables: 1, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('alphabeta gamma');
  });

  it('two textboxes in one host paragraph: each textbox paragraph is its own part (3 words)', async () => {
    const buf = makeDocx(TEXTBOX_TWO);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 3, paragraphs: 0, tables: 1, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('aadd bb cc');
  });

  it('reviewer shape: textbox between host runs keeps host fusion (alphagamma beta)', async () => {
    const fixt = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>alpha</w:t></w:r>${drawing(para('beta'), 11)}<w:r><w:t>gamma</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const buf = makeDocx(fixt);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 2, paragraphs: 0, tables: 1, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('alphagamma beta');
  });

  it('reviewer shape: two drawings between host runs keep host fusion (leadtail ...)', async () => {
    const fixt = `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>lead</w:t></w:r>${drawing(para('first box'), 21)}${drawing(para('second box'), 22)}<w:r><w:t>tail</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const buf = makeDocx(fixt);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 5, paragraphs: 0, tables: 1, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe('leadtail first box second box');
  });

  it('textbox at body level: paragraph capture unchanged (5 words)', async () => {
    const buf = makeDocx(TEXTBOX_BODY_LEVEL);
    const raw = rawDocumentXml(buf);
    const { stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 5, paragraphs: 1, tables: 0, images: 0 });
    expect(stats.words).toBe(oracleWords(raw));
  });

  it('flat fixture: counts byte-identical to the pre-fix flat algorithm', async () => {
    const buf = makeDocx(FLAT);
    const raw = rawDocumentXml(buf);
    const { ast, stats } = await statsOf(buf);
    expect(stats).toMatchObject({ words: 32, paragraphs: 2, tables: 1, images: 0 });
    expect(stats.tables).toBe(oracleTables(raw));
    expect(tableCellTextSpaced(tableNode(ast))).toBe(legacyFlatSpaced(tableNode(ast).xml));
  });

  it('run fusion inside one flat cell is unchanged (adjacent w:t runs still join)', async () => {
    const fused =
      para('one two three') +
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
      `<w:tr><w:tc><w:p><w:r><w:t>alpha</w:t></w:r><w:r><w:t>beta</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>gamma delta</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const buf = makeDocx(fused);
    const { ast, stats } = await statsOf(buf);
    // 3 paragraph words + fused "alphabeta" (1) + "gamma delta" (2)
    expect(stats.words).toBe(6);
    expect(stats.tables).toBe(1);
    expect(tableCellTextSpaced(tableNode(ast))).toBe('alphabeta gamma delta');
  });
});
