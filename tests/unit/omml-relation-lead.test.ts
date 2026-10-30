import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { latexToOmml, ommlToLatex } from '../../packages/engine/src/ooxml/omml';

const NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math';
function conv(latex: string): { xml: string; el: Element } {
  const doc = new DOMParser().parseFromString(`<root xmlns:m="${NS}"/>`, 'text/xml');
  const el = latexToOmml(doc as unknown as Document, latex);
  return { xml: new XMLSerializer().serializeToString(el as unknown as Node), el: el as unknown as Element };
}
const run = (t: string) => `<m:r><m:t>${t}</m:t></m:r>`;

describe('a leading relation stays out of a script base', () => {
  it('T_{a}=T_{b}: = is its own run, the second base is T', () => {
    const { xml } = conv('T_{a}=T_{b}');
    expect(xml).toContain(`</m:sSub>${run('=')}<m:sSub><m:e>${run('T')}</m:e><m:sub>${run('b')}`);
  });

  it('a^2+b^2=c^2 and x_1=y_2+z_3 split every =, and keep the E1 base +b / +z', () => {
    const a = conv('a^2+b^2=c^2').xml;
    expect(a).toContain(`<m:e>${run('+b')}</m:e>`);
    expect(a).toContain(`${run('=')}<m:sSup><m:e>${run('c')}</m:e>`);
    const b = conv('x_1=y_2+z_3').xml;
    expect(b).toContain(`${run('=')}<m:sSub><m:e>${run('y')}</m:e>`);
    expect(b).toContain(`<m:e>${run('+z')}</m:e>`);
  });

  it('keeps the E1 rule for everything else', () => {
    expect(conv('x=T_{a}').xml).toContain(`<m:sSub><m:e>${run('x=T')}</m:e>`);
    expect(conv('F=ma').xml).toContain(run('F=ma'));
  });

  it('a lone relation before a script stays the base and does not throw', () => {
    expect(conv('=^2').xml).toContain(`<m:sSup><m:e>${run('=')}</m:e>`);
  });

  it('round-trips: ommlToLatex output converts back to identical OMML', () => {
    for (const latex of ['T_{a}=T_{b}', 'a^2+b^2=c^2', '{=ab}_i', '{<x}^2']) {
      const { el, xml } = conv(latex);
      const back = ommlToLatex(el);
      expect(back, latex).not.toBeNull();
      expect(conv(back as string).xml, latex).toBe(xml);
    }
  });
});
