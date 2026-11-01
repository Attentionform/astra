import { describe, expect, it } from 'vitest';
import { assertNoXxe, assertNoXxeAdjacent, containsXInclude, containsXslt, containsStylesheetPI } from '../../packages/engine/src/ooxml/xxe-guard';

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const wrap = (inner: string) => `<?xml version="1.0"?><w:document xmlns:w="${W}"><w:body>${inner}</w:body></w:document>`;
const XSL = 'http://www.w3.org/1999/XSL/Transform';

describe('XXE-adjacent check looks at markup, not at text', () => {
  it('escaped XSLT written as text is not refused', () => {
    const xml = wrap(`<w:p><w:r><w:t>&lt;xsl:stylesheet xmlns:xsl="${XSL}" version="2.0"&gt;</w:t></w:r></w:p>`);
    expect(containsXslt(xml)).toBe(false);
    expect(() => assertNoXxeAdjacent(xml, 'word/document.xml')).not.toThrow();
  });
  it('the namespace address alone in text is not refused', () => {
    expect(containsXslt(wrap(`<w:p><w:r><w:t>see ${XSL}</w:t></w:r></w:p>`))).toBe(false);
  });
  it('constructs inside a comment or a CDATA section are text', () => {
    expect(containsXslt(wrap(`<!-- <xsl:template match="/"/> --><w:p><w:r><w:t><![CDATA[<xsl:stylesheet xmlns:xsl="${XSL}">]]></w:t></w:r></w:p>`))).toBe(false);
    expect(containsXInclude(wrap(`<w:p><w:r><w:t><![CDATA[<xi:include href="a"/>]]></w:t></w:r></w:p>`))).toBe(false);
    expect(containsStylesheetPI(wrap(`<!-- <?xml-stylesheet href="a.xsl"?> -->`))).toBe(false);
  });
  it('an xsl element, an xsl namespace declaration, an xi:include and a stylesheet instruction are still refused', () => {
    expect(containsXslt(wrap(`<xsl:template match="/"/>`))).toBe(true);
    expect(containsXslt(`<w:document xmlns:w="${W}" xmlns:x="${XSL}"/>`)).toBe(true);
    expect(containsXInclude(wrap(`<xi:include href="a.xml"/>`))).toBe(true);
    expect(containsStylesheetPI(`<?xml version="1.0"?><?xml-stylesheet type="text/xsl" href="a.xsl"?><w:document xmlns:w="${W}"/>`)).toBe(true);
    expect(() => assertNoXxeAdjacent(wrap(`<xsl:template match="/"/>`), 'word/document.xml')).toThrow(/XXE-adjacent/);
  });
  it('a ">" inside an attribute value does not hide a declaration', () => {
    expect(containsXslt(`<a x=">" xmlns:t="${XSL}"/>`)).toBe(true);
  });
  it('a document type declaration with an entity is still refused', () => {
    expect(() => assertNoXxe('<!DOCTYPE d [<!ENTITY e SYSTEM "file:///etc/passwd">]><d>&e;</d>', 'word/document.xml')).toThrow();
  });
});
