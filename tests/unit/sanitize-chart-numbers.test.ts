import { describe, it, expect } from 'vitest';
import {
  redactAttributePii,
  redactTextNodePii,
  sanitizeDocumentXml,
} from '../../packages/engine/src/sanitizer';

/**
 * A chart part stores its data and layout as long decimal numbers, such as 0.1234567890123.
 * The card number pattern matched them and wrote "0.[redacted]", which is not a number: the
 * validator refused the file and the chart lost its data. A value that is only a decimal number
 * is not personal data and stays as written.
 */
const CHART = (inner: string) =>
  `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${inner}</c:chartSpace>`;

describe('purgePII keeps decimal numbers in chart parts', () => {
  it('keeps a long decimal in c:v', () => {
    const xml = CHART('<c:numCache><c:pt idx="0"><c:v>0.1234567890123</c:v></c:pt><c:pt idx="1"><c:v>3.14159265358979</c:v></c:pt></c:numCache>');
    const r = sanitizeDocumentXml(xml, { purgePII: true }, 'word/charts/chart1.xml');
    expect(r.xml).toBe(xml);
    expect(r.bodyPII).toBe(0);
  });

  it('keeps a long decimal and an exponent value in a layout attribute', () => {
    const xml = CHART('<c:manualLayout><c:x val="0.1234567890123"/><c:y val="6.2322946175637411E-2"/><c:w val="-0.5555555555555"/></c:manualLayout>');
    const r = sanitizeDocumentXml(xml, { purgePII: true }, 'word/charts/chart1.xml');
    expect(r.xml).toBe(xml);
    expect(r.bodyPII).toBe(0);
  });

  it('still redacts personal data in chart text and attributes', () => {
    const xml = CHART(
      '<c:tx><c:rich><a:p><a:r><a:t>Contact jane.doe@example.com</a:t></a:r></a:p></c:rich></c:tx>' +
        '<c:pt idx="0"><c:v>4111 1111 1111 1111</c:v></c:pt><c:ext uri="x" descr="Phone: 5551234567"/>'
    );
    const r = sanitizeDocumentXml(xml, { purgePII: true }, 'word/charts/chart1.xml');
    expect(r.xml).not.toContain('jane.doe@example.com');
    expect(r.xml).not.toContain('4111 1111 1111 1111');
    expect(r.xml).not.toContain('5551234567');
    expect(r.bodyPII).toBe(3);
  });

  it('still redacts a card number that sits next to a decimal in the same part', () => {
    const xml = CHART('<c:pt idx="0"><c:v>0.1234567890123</c:v></c:pt><c:pt idx="1"><c:v>4111111111111111</c:v></c:pt>');
    const r = sanitizeDocumentXml(xml, { purgePII: true }, 'word/charts/chart1.xml');
    expect(r.xml).toContain('<c:v>0.1234567890123</c:v>');
    expect(r.xml).not.toContain('4111111111111111');
    expect(r.bodyPII).toBe(1);
  });

  it('a decimal number is kept by both passes on their own', () => {
    expect(redactTextNodePii('<c:v>2.718281828459045</c:v>', true)).toEqual({ xml: '<c:v>2.718281828459045</c:v>', count: 0 });
    expect(redactAttributePii('<c:x val="0.1234567890123"/>', true)).toEqual({ xml: '<c:x val="0.1234567890123"/>', count: 0 });
  });

  it('a number with a label or other text around it is still scanned', () => {
    const r = redactAttributePii('<x a="Phone: 5551234567"/>', true);
    expect(r.count).toBe(1);
    expect(r.xml).not.toContain('5551234567');
  });
});
