import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { AdmZip, createBlankDocxBuffer, REPAIR_RULES, repairDocxBuffer, type RepairRule, type ValidationResult } from '@attentionform/astra-engine';
import { parsePartStrict } from '../../packages/engine/src/ooxml/normalize';

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
const rule = (id: string): RepairRule => REPAIR_RULES.find((r) => r.id === id) as RepairRule;

/** Runs one rule on one part. Returns the new XML and the log. */
function run(id: string, part: string, xml: string, hasBidi = false): { xml: string; log: string[] } {
  const doc = parsePartStrict(part, xml);
  if (!doc) throw new Error('no document');
  const log = rule(id).apply(doc, part, { hasBidi });
  return { xml: log.length > 0 ? new XMLSerializer().serializeToString(doc) : xml, log };
}

/** Applies the rule twice: the second run must change nothing. */
function runTwice(id: string, part: string, xml: string, hasBidi = false) {
  const first = run(id, part, xml, hasBidi);
  const second = run(id, part, first.xml, hasBidi);
  expect(second.log, `${id} is idempotent`).toEqual([]);
  return first;
}

const doc = (body: string) => `<w:document ${W}><w:body>${body}</w:body></w:document>`;

describe('repair rules', () => {
  it('settings-order puts w:zoom before w:compat', () => {
    const r = runTwice('settings-order', 'word/settings.xml', `<w:settings ${W}><w:compat/><w:zoom w:percent="100"/></w:settings>`);
    expect(r.xml.indexOf('<w:zoom')).toBeLessThan(r.xml.indexOf('<w:compat'));
    expect(r.log).toHaveLength(1);
  });

  it('style-order puts w:name before w:basedOn and w:rPr after w:pPr', () => {
    const r = runTwice('style-order', 'word/styles.xml', `<w:styles ${W}><w:style w:type="paragraph" w:styleId="A"><w:basedOn w:val="N"/><w:rPr><w:sz w:val="20"/></w:rPr><w:name w:val="A"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style></w:styles>`);
    expect(r.xml.indexOf('<w:name')).toBeLessThan(r.xml.indexOf('<w:basedOn'));
    expect(r.xml.indexOf('<w:pPr')).toBeLessThan(r.xml.indexOf('<w:rPr'));
  });

  it('style-order sorts a table property container inside a style', () => {
    const r = runTwice('style-order', 'word/styles.xml', `<w:styles ${W}><w:style w:type="table" w:styleId="T"><w:name w:val="T"/><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblStyleRowBandSize w:val="1"/></w:tblPr></w:style></w:styles>`);
    expect(r.xml.indexOf('tblStyleRowBandSize')).toBeLessThan(r.xml.indexOf('<w:tblW'));
  });

  it('sectpr-order puts w:pgSz before w:pgMar and keeps header and footer references where they are', () => {
    const sect = '<w:sectPr><w:footerReference w:type="default" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/><w:headerReference w:type="default" r:id="rId2" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/>';
    const r = runTwice('sectpr-order', 'word/document.xml', doc(`${sect}<w:pgMar w:top="1"/><w:pgSz w:w="1" w:h="1"/></w:sectPr>`));
    expect(r.xml.indexOf('<w:pgSz')).toBeLessThan(r.xml.indexOf('<w:pgMar'));
    expect(r.xml.indexOf('footerReference')).toBeLessThan(r.xml.indexOf('headerReference'));
  });

  it('table-children-order puts w:tblPr before w:tblGrid', () => {
    const r = runTwice('table-children-order', 'word/document.xml', doc('<w:tbl><w:tblGrid><w:gridCol w:w="1"/></w:tblGrid><w:tblPr/><w:tr/></w:tbl>'));
    expect(r.xml.indexOf('<w:tblPr')).toBeLessThan(r.xml.indexOf('<w:tblGrid'));
  });

  it('run-properties-first moves w:rPr to the start of the run and keeps the text', () => {
    const r = runTwice('run-properties-first', 'word/document.xml', doc('<w:p><w:r><w:t>hi</w:t><w:rPr><w:b/></w:rPr></w:r></w:p>'));
    expect(r.xml).toContain('<w:r><w:rPr><w:b/></w:rPr><w:t>hi</w:t></w:r>');
  });

  it('integer-attribute rounds page size and cell margins only', () => {
    const r = runTwice('integer-attribute', 'word/document.xml', doc('<w:sectPr><w:pgSz w:w="11905.5" w:h="16837.795"/></w:sectPr><w:tbl><w:tblPr><w:tblCellMar><w:left w:w="70.0" w:type="dxa"/></w:tblCellMar></w:tblPr></w:tbl><w:p><w:r><w:t w:foo="1.5">x</w:t></w:r></w:p>'));
    expect(r.xml).toContain('w:w="11906"');
    expect(r.xml).toContain('w:h="16838"');
    expect(r.xml).toContain('w:w="70"');
    expect(r.xml).toContain('w:foo="1.5"');
  });

  it('on-off-attribute writes 1 and 0 as on and off for the listed elements only', () => {
    const r = runTwice('on-off-attribute', 'word/styles.xml', `<w:styles ${W}><w:style w:type="paragraph" w:styleId="A"><w:qFormat w:val="1"/><w:semiHidden w:val="0"/><w:rPr><w:b w:val="0"/></w:rPr></w:style></w:styles>`);
    expect(r.xml).toContain('<w:qFormat w:val="on"/>');
    expect(r.xml).toContain('<w:semiHidden w:val="off"/>');
    expect(r.xml).toContain('<w:b w:val="0"/>');
  });

  it('grid-column-type removes the undeclared attribute', () => {
    const r = runTwice('grid-column-type', 'word/document.xml', doc('<w:tbl><w:tblGrid><w:gridCol w:w="5" w:type="dxa"/></w:tblGrid></w:tbl>'));
    expect(r.xml).toContain('<w:gridCol w:w="5"/>');
  });

  it('font-family-case writes Auto as auto', () => {
    const r = runTwice('font-family-case', 'word/fontTable.xml', `<w:fonts ${W}><w:font w:name="X"><w:family w:val="Auto"/></w:font></w:fonts>`);
    expect(r.xml).toContain('w:val="auto"');
  });

  it('alignment-start-end maps start and end when no part is right to left, and does nothing otherwise', () => {
    const xml = doc('<w:p><w:pPr><w:jc w:val="start"/></w:pPr></w:p><w:p><w:pPr><w:jc w:val="end"/></w:pPr></w:p>');
    const r = runTwice('alignment-start-end', 'word/document.xml', xml);
    expect(r.xml).toContain('w:val="left"');
    expect(r.xml).toContain('w:val="right"');
    expect(run('alignment-start-end', 'word/document.xml', xml, true).log).toEqual([]);
  });

  it('smarttag-unwrap keeps the runs and drops the tag and its properties', () => {
    const r = runTwice('smarttag-unwrap', 'word/document.xml', doc('<w:p><w:smartTag w:uri="u" w:element="e"><w:smartTagPr><w:attr w:name="a" w:val="b"/></w:smartTagPr><w:r><w:t>kept</w:t></w:r></w:smartTag></w:p>'));
    expect(r.xml).toContain('<w:p><w:r><w:t>kept</w:t></w:r></w:p>');
    expect(r.xml).not.toContain('smartTag');
  });

  it('shapetype-duplicate removes an identical repeat and keeps a different one with the same id', () => {
    const v = 'xmlns:v="urn:schemas-microsoft-com:vml"';
    const same = doc(`<w:p><w:pict ${v}><v:shapetype id="t1" coordsize="1,1"/><v:shapetype id="t1" coordsize="1,1"/><v:shapetype id="t1" coordsize="2,2"/></w:pict></w:p>`);
    const r = runTwice('shapetype-duplicate', 'word/document.xml', same);
    expect((r.xml.match(/<v:shapetype/g) ?? []).length).toBe(2);
  });
});

describe('rules added after the first measurement', () => {
  it('settings-legacy-smarttags removes w:smartTagType and w:doNotEmbedSmartTags and keeps the rest', () => {
    const r = runTwice('settings-legacy-smarttags', 'word/settings.xml', `<w:settings ${W}><w:zoom w:percent="100"/><w:smartTagType w:namespaceuri="urn:x" w:name="y"/><w:doNotEmbedSmartTags/><w:decimalSymbol w:val="."/></w:settings>`);
    expect(r.xml).not.toContain('smartTag');
    expect(r.xml).not.toContain('doNotEmbedSmartTags');
    expect(r.xml).toContain('<w:zoom w:percent="100"/>');
    expect(r.xml).toContain('<w:decimalSymbol w:val="."/>');
  });

  it('style-table-defaults removes default table width, layout and bidiVisual from a style, not other values', () => {
    const r = runTwice('style-table-defaults', 'word/styles.xml', `<w:styles ${W}><w:style w:type="table" w:styleId="A"><w:name w:val="A"/><w:tblPr><w:bidiVisual w:val="off"/><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single"/></w:tblBorders><w:tblLayout w:type="autofit"/></w:tblPr></w:style><w:style w:type="table" w:styleId="B"><w:name w:val="B"/><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr></w:style></w:styles>`);
    expect(r.xml).not.toContain('w:bidiVisual');
    expect(r.xml).toContain('<w:tblW w:w="5000" w:type="pct"/>');
    expect(r.xml).toContain('<w:tblLayout w:type="fixed"/>');
    expect(r.xml).toContain('<w:tblBorders>');
    expect((r.xml.match(/<w:tblW /g) ?? []).length).toBe(1);
  });

  it('style-table-defaults leaves a table in the document alone', () => {
    const xml = doc('<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="autofit"/></w:tblPr></w:tbl>');
    expect(run('style-table-defaults', 'word/styles.xml', xml).log).toEqual([]);
  });

  it('paragraph-bare-run-content puts a bare w:br, w:tab and w:drawing into runs and keeps the order and the text', () => {
    const r = runTwice('paragraph-bare-run-content', 'word/document.xml', doc('<w:p><w:r><w:t>a</w:t></w:r><w:br/><w:tab/><w:r><w:t>b</w:t></w:r><w:drawing/></w:p>'));
    expect(r.xml).toContain('<w:p><w:r><w:t>a</w:t></w:r><w:r><w:br/></w:r><w:r><w:tab/></w:r><w:r><w:t>b</w:t></w:r><w:r><w:drawing/></w:r></w:p>');
  });

  it('paragraph-bare-run-content leaves a w:br inside a run alone', () => {
    expect(run('paragraph-bare-run-content', 'word/document.xml', doc('<w:p><w:r><w:br/></w:r></w:p>')).log).toEqual([]);
  });
});

describe('every rule leaves a valid part byte for byte', () => {
  const valid = doc('<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>ok</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1" w:right="1" w:bottom="1" w:left="1"/></w:sectPr>');
  for (const r of REPAIR_RULES) {
    it(r.id, () => {
      const part = ['word/settings.xml', 'word/styles.xml', 'word/fontTable.xml', 'word/document.xml'].find((n) => r.appliesTo(n)) as string;
      expect(run(r.id, part, valid).log).toEqual([]);
    });
  }
});

// ── runner ──────────────────────────────────────────────────────────────

const okResult = { file: 'x', ok: true, errors: [] } as unknown as ValidationResult;
const noValidator = async () => okResult;

function packageWith(parts: Record<string, string>): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  for (const [name, xml] of Object.entries(parts)) {
    if (zip.getEntry(name)) zip.updateFile(name, Buffer.from(xml, 'utf8'));
    else zip.addFile(name, Buffer.from(xml, 'utf8'));
  }
  return zip.toBuffer();
}
const partOf = (buf: Buffer, name: string) => new AdmZip(buf).getEntry(name)?.getData().toString('utf8') ?? '';

describe('repairDocxBuffer', () => {
  const settings = `<w:settings ${W}><w:compat/><w:zoom w:percent="100"/></w:settings>`;
  const body = doc('<w:p><w:smartTag w:uri="u" w:element="e"><w:r><w:t>same text</w:t></w:r></w:smartTag></w:p>');

  it('repairs a package and reports what it changed, with the same visible text', async () => {
    const input = packageWith({ 'word/settings.xml': settings, 'word/document.xml': body });
    const out = await repairDocxBuffer(input, { validate: noValidator });
    expect(out.normalized.length).toBeGreaterThanOrEqual(2);
    expect(out.rolledBack).toEqual([]);
    expect(partOf(out.buffer, 'word/document.xml')).toContain('same text');
    expect(partOf(out.buffer, 'word/document.xml')).not.toContain('smartTag');
    expect(partOf(out.buffer, 'word/settings.xml').indexOf('<w:zoom')).toBeLessThan(partOf(out.buffer, 'word/settings.xml').indexOf('<w:compat'));
  });

  it('a second run changes nothing and returns the same bytes', async () => {
    const first = await repairDocxBuffer(packageWith({ 'word/settings.xml': settings }), { validate: noValidator });
    const second = await repairDocxBuffer(first.buffer, { validate: noValidator });
    expect(second.normalized).toEqual([]);
    expect(second.buffer.equals(first.buffer)).toBe(true);
  });

  it('returns the input untouched when no rule applies', async () => {
    const input = createBlankDocxBuffer('blank');
    const out = await repairDocxBuffer(input, { validate: noValidator });
    expect(out.normalized).toEqual([]);
    expect(out.buffer.equals(input)).toBe(true);
  });

  it('drops only the rule that adds a validator error, and keeps the others', async () => {
    const input = packageWith({ 'word/settings.xml': settings, 'word/document.xml': body });
    // A validator test double that reports a new error once the smart tag is gone.
    const validate = async (buf: Buffer): Promise<ValidationResult> => {
      const gone = !partOf(buf, 'word/document.xml').includes('smartTag');
      return gone
        ? ({ file: 'x', ok: false, errors: [{ part: '/word/document.xml', id: 'Sch_Test', xPath: '/w:document[1]/w:body[1]', description: 'new', errorType: 'Schema' }] } as unknown as ValidationResult)
        : okResult;
    };
    const out = await repairDocxBuffer(input, { validate });
    expect(out.rolledBack.map((r) => r.rule)).toEqual(['smarttag-unwrap']);
    expect(out.rolledBack[0].reason).toContain('adds 1 validator error');
    expect(partOf(out.buffer, 'word/document.xml')).toContain('smartTag');
    expect(partOf(out.buffer, 'word/settings.xml').indexOf('<w:zoom')).toBeLessThan(partOf(out.buffer, 'word/settings.xml').indexOf('<w:compat'));
  });

  it('drops a rule that would change the visible text', async () => {
    const changing: RepairRule = {
      id: 'text-changer',
      description: 'test double',
      appliesTo: (p) => p === 'word/document.xml',
      apply(d) {
        const t = d.getElementsByTagName('w:t')[0];
        (t.firstChild as Text).data = 'different';
        return ['changed text'];
      },
    };
    const input = packageWith({ 'word/document.xml': body });
    const out = await repairDocxBuffer(input, { rules: [changing], validate: noValidator });
    expect(out.normalized).toEqual([]);
    expect(out.rolledBack).toEqual([{ rule: 'text-changer', reason: 'visible text would change' }]);
    expect(out.buffer.equals(input)).toBe(true);
  });

  it('reports the number of validator errors that remain', async () => {
    const still = async () => ({ file: 'x', ok: false, errors: [{ part: '/word/styles.xml', id: 'Sch_Other', xPath: '/w:styles[1]', description: 'left', errorType: 'Schema' }] }) as unknown as ValidationResult;
    const out = await repairDocxBuffer(packageWith({ 'word/settings.xml': settings }), { validate: still });
    expect(out.remaining).toBe(1);
    expect(out.rolledBack).toEqual([]);
  });
});
