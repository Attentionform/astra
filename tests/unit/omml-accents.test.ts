import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { latexToOmml, ommlToLatex } from '../../packages/engine/src/ooxml/omml';

const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
function conv(latex: string): { xml: string; el: Element } {
  const doc = new DOMParser().parseFromString(`<root xmlns:m="${NS}"/>`, 'text/xml');
  const el = latexToOmml(doc as unknown as Document, latex);
  return { xml: new XMLSerializer().serializeToString(el as unknown as Node), el: el as unknown as Element };
}

describe('accents', () => {
  it.each([
    ['dot', '\u0307'],
    ['ddot', '\u0308'],
    ['hat', '\u0302'],
    ['bar', '\u0304'],
    ['vec', '\u20D7'],
    ['tilde', '\u0303'],
  ])('\\%s becomes m:acc with the combining character', (name, chr) => {
    const { xml } = conv(`\\${name}{x}`);
    expect(xml).toContain('<m:acc>');
    expect(xml).toContain(`m:val="${chr}"`);
    expect(xml).toMatch(/<m:e><m:r><m:t>x<\/m:t><\/m:r><\/m:e>/);
  });

  it('composes with scripts and round-trips to LaTeX', () => {
    const { el, xml } = conv('\\dot{x}=v_{0}');
    expect(xml).toContain('<m:acc>');
    const back = ommlToLatex(el);
    expect(back).toContain('\\dot{x}');
    expect(conv(back as string).xml).toBe(xml);
  });

  it('still refuses genuinely unsupported macros', () => {
    expect(() => conv('\\foo{x}')).toThrow(/unsupported macro/);
  });
});
