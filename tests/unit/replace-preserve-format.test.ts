/**
 * Replace preserve-or-carry pins.
 *
 * A default replace (no flags) carries the target paragraph's run properties
 * onto the rewritten text. Explicit op properties win over carried ones
 * (fontSize/sz overwrites carried w:sz/w:szCs, everything else stays);
 * resetStyle:true and preserveFormatting:false strip.
 */
import { describe, it, expect } from 'vitest';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';
import { executeReplace } from '@attentionform/astra-engine';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const serializer = new XMLSerializer();

function styledPara(innerRuns: string): Element {
  return new DOMParser().parseFromString(
    `<w:p xmlns:w="${W}">${innerRuns}</w:p>`,
    'text/xml',
  ).documentElement as unknown as Element;
}

function replaceAst(p: Element, trackChanges = false): any {
  return {
    metadata: {},
    sections: [{
      id: 'sec_root', heading: null,
      nodes: [{
        id: 'p_1', type: 'paragraph', xml: p,
        metadata: {
          text: 'old', hasCitation: false, hasComment: false, hasBookmark: false,
          hasTrackedChanges: false, hasEquation: false, hasField: false,
          hasImage: false, hasHyperlink: false,
        },
        version: 1, checksum: 'c',
      }],
      pageSetup: {}, version: 1, checksum: 's',
    }],
    styles: {}, numbering: { abstractNum: {}, concreteNum: {} },
    relationships: {}, headers: {}, footers: {}, mediaFiles: {},
    settings: { trackChanges, defaultAuthor: 'ASTra' }, nextChangeId: 1,
  };
}

function nodeXml(ast: any): string {
  return serializer.serializeToString(ast.sections[0].nodes[0].xml);
}

const BOLD_RUN =
  `<w:r><w:rPr><w:b/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr>` +
  `<w:t xml:space="preserve">old</w:t></w:r>`;

describe('replace carries target run properties by default', () => {
  it('bold survives a flag-free replace', () => {
    const ast = replaceAst(styledPara(BOLD_RUN));
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('new wording');
    expect(xml).toContain('<w:b');
    expect(xml).toContain('w:val="28"');
  });

  it('bold plus italic plus color round-trip together', () => {
    const p = styledPara(
      `<w:r><w:rPr><w:b/><w:i/><w:color w:val="FF0000"/><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr>` +
      `<w:t xml:space="preserve">old</w:t></w:r>`,
    );
    const ast = replaceAst(p);
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:b');
    expect(xml).toContain('<w:i');
    expect(xml).toContain('w:val="FF0000"');
    expect(xml).toContain('w:val="24"');
  });

  it('formatted replace is byte-deterministic across two runs from the same source', () => {
    const first = replaceAst(styledPara(BOLD_RUN));
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording' } as any, first);
    const second = replaceAst(styledPara(BOLD_RUN));
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording' } as any, second);
    expect(nodeXml(second)).toBe(nodeXml(first));
  });

  it('explicit fontSize overwrites carried size but keeps bold and color', () => {
    const p = styledPara(
      `<w:r><w:rPr><w:b/><w:color w:val="0000FF"/><w:sz w:val="20"/><w:szCs w:val="20"/></w:rPr>` +
      `<w:t xml:space="preserve">old</w:t></w:r>`,
    );
    const ast = replaceAst(p);
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording', fontSize: '12pt' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('<w:b');
    expect(xml).toContain('w:val="0000FF"');
    expect(xml).toContain('w:val="24"');
    expect(xml).not.toContain('w:val="20"');
  });

  it('merged size stays in schema order (w:sz before w:u)', () => {
    const p = styledPara(
      `<w:r><w:rPr><w:b/><w:sz w:val="20"/><w:szCs w:val="20"/><w:u w:val="single"/></w:rPr>` +
      `<w:t xml:space="preserve">old</w:t></w:r>`,
    );
    const ast = replaceAst(p);
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording', fontSize: '12pt' } as any, ast);
    const xml = nodeXml(ast);
    const iSz = xml.indexOf('<w:sz ');
    const iU = xml.indexOf('<w:u ');
    expect(iSz).toBeGreaterThanOrEqual(0);
    expect(iU).toBeGreaterThanOrEqual(0);
    expect(iSz).toBeLessThan(iU);
  });

  it('resetStyle:true strips carried properties', () => {
    const ast = replaceAst(styledPara(BOLD_RUN));
    executeReplace({ action: 'replace', target: 'p_1', text: 'plain', resetStyle: true } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('plain');
    expect(xml).not.toContain('<w:b');
    expect(xml).not.toContain('w:val="28"');
  });

  it('explicit preserveFormatting:false keeps the legacy strip behavior', () => {
    const ast = replaceAst(styledPara(BOLD_RUN));
    executeReplace({ action: 'replace', target: 'p_1', text: 'plain', preserveFormatting: false } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('plain');
    expect(xml).not.toContain('<w:b');
  });

  it('unformatted paragraphs still emit a bare run (no properties minted)', () => {
    const ast = replaceAst(styledPara(`<w:r><w:t xml:space="preserve">old</w:t></w:r>`));
    executeReplace({ action: 'replace', target: 'p_1', text: 'new wording' } as any, ast);
    const xml = nodeXml(ast);
    expect(xml).toContain('new wording');
    expect(xml).not.toContain('w:rPr');
  });
});
