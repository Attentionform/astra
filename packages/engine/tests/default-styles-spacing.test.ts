import { describe, it, expect } from 'vitest';
import { DOMParser } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';
import { DocxParser } from '../src/parser';

describe('default styles typographic baseline spacing', () => {
  function getStylesXml(): { xml: string; doc: Document } {
    const parser = DocxParser.createNewDocument('blank');
    const zip = new AdmZip(parser.rezip());
    const xml = zip.getEntry('word/styles.xml')?.getData().toString('utf8') ?? '';
    const doc = new DOMParser().parseFromString(xml, 'text/xml');
    return { xml, doc };
  }

  function getStyleElement(doc: Document, styleId: string): Element {
    const styles = Array.from(doc.getElementsByTagName('w:style'));
    const match = styles.find((s) => s.getAttribute('w:styleId') === styleId);
    expect(match).toBeDefined();
    return match!;
  }

  it('Heading1 has keepNext, keepLines, widowControl, and spacing 360/160/276', () => {
    const { doc } = getStylesXml();
    const h1 = getStyleElement(doc, 'Heading1');
    const pPr = h1.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();

    expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
    expect(pPr.getElementsByTagName('w:keepLines').length).toBe(1);
    expect(pPr.getElementsByTagName('w:widowControl').length).toBe(1);

    const spacing = pPr.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('360');
    expect(spacing.getAttribute('w:after')).toBe('160');
    expect(spacing.getAttribute('w:line')).toBe('276');
    expect(spacing.getAttribute('w:lineRule')).toBe('auto');

    const childTags = Array.from(pPr.childNodes)
      .filter((n) => n.nodeType === 1)
      .map((e) => (e as Element).tagName);
    expect(childTags).toEqual(['w:keepNext', 'w:keepLines', 'w:widowControl', 'w:spacing']);
  });

  it('Heading2 has keepNext, keepLines, widowControl, spacing 280/120/260, and font size 32', () => {
    const { doc } = getStylesXml();
    const h2 = getStyleElement(doc, 'Heading2');
    const pPr = h2.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();

    expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
    expect(pPr.getElementsByTagName('w:keepLines').length).toBe(1);
    expect(pPr.getElementsByTagName('w:widowControl').length).toBe(1);

    const spacing = pPr.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('280');
    expect(spacing.getAttribute('w:after')).toBe('120');
    expect(spacing.getAttribute('w:line')).toBe('260');
    expect(spacing.getAttribute('w:lineRule')).toBe('auto');

    const sz = h2.getElementsByTagName('w:sz')[0];
    expect(sz).toBeDefined();
    expect(sz.getAttribute('w:val')).toBe('32');
  });

  it('Heading3 has keepNext, keepLines, widowControl, and spacing 200/80/240', () => {
    const { doc } = getStylesXml();
    const h3 = getStyleElement(doc, 'Heading3');
    const pPr = h3.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();

    expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
    expect(pPr.getElementsByTagName('w:keepLines').length).toBe(1);
    expect(pPr.getElementsByTagName('w:widowControl').length).toBe(1);

    const spacing = pPr.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('200');
    expect(spacing.getAttribute('w:after')).toBe('80');
    expect(spacing.getAttribute('w:line')).toBe('240');
    expect(spacing.getAttribute('w:lineRule')).toBe('auto');
  });

  it('Title has keepNext, keepLines, and spacing 0/240/240', () => {
    const { doc } = getStylesXml();
    const title = getStyleElement(doc, 'Title');
    const pPr = title.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();

    expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
    expect(pPr.getElementsByTagName('w:keepLines').length).toBe(1);

    const spacing = pPr.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('0');
    expect(spacing.getAttribute('w:after')).toBe('240');
    expect(spacing.getAttribute('w:line')).toBe('240');
    expect(spacing.getAttribute('w:lineRule')).toBe('auto');
  });

  it('Subtitle has keepNext, keepLines, and spacing 0/360/240', () => {
    const { doc } = getStylesXml();
    const subtitle = getStyleElement(doc, 'Subtitle');
    const pPr = subtitle.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();

    expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
    expect(pPr.getElementsByTagName('w:keepLines').length).toBe(1);

    const spacing = pPr.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('0');
    expect(spacing.getAttribute('w:after')).toBe('360');
    expect(spacing.getAttribute('w:line')).toBe('240');
    expect(spacing.getAttribute('w:lineRule')).toBe('auto');
  });

  it('Caption style exists with valid structure and spacing 120/120/240', () => {
    const { doc, xml } = getStylesXml();
    const caption = getStyleElement(doc, 'Caption');
    expect(caption.getAttribute('w:type')).toBe('paragraph');

    const name = caption.getElementsByTagName('w:name')[0];
    expect(name).toBeDefined();
    expect(name.getAttribute('w:val')).toBe('caption');

    const basedOn = caption.getElementsByTagName('w:basedOn')[0];
    expect(basedOn).toBeDefined();
    expect(basedOn.getAttribute('w:val')).toBe('Normal');

    const next = caption.getElementsByTagName('w:next')[0];
    expect(next).toBeDefined();
    expect(next.getAttribute('w:val')).toBe('Normal');

    const uiPriority = caption.getElementsByTagName('w:uiPriority')[0];
    expect(uiPriority).toBeDefined();
    expect(uiPriority.getAttribute('w:val')).toBe('9');

    expect(caption.getElementsByTagName('w:qFormat').length).toBe(1);

    const pPr = caption.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();
    expect(pPr.getElementsByTagName('w:keepLines').length).toBe(1);

    const spacing = pPr.getElementsByTagName('w:spacing')[0];
    expect(spacing).toBeDefined();
    expect(spacing.getAttribute('w:before')).toBe('120');
    expect(spacing.getAttribute('w:after')).toBe('120');
    expect(spacing.getAttribute('w:line')).toBe('240');
    expect(spacing.getAttribute('w:lineRule')).toBe('auto');

    const subtitleIdx = xml.indexOf('w:styleId="Subtitle"');
    const captionIdx = xml.indexOf('w:styleId="Caption"');
    const codeIdx = xml.indexOf('w:styleId="Code"');
    expect(subtitleIdx).toBeLessThan(captionIdx);
    expect(captionIdx).toBeLessThan(codeIdx);

    const childTags = Array.from(caption.childNodes)
      .filter((n) => n.nodeType === 1)
      .map((e) => (e as Element).tagName);
    expect(childTags).toEqual(['w:name', 'w:basedOn', 'w:next', 'w:uiPriority', 'w:qFormat', 'w:pPr']);
  });
});
