/** DN3 8.5: after accept-all only the revisions the report lists may remain; identity match first, bounded count second. */
import { describe, expect, it } from 'vitest';
import { assertOnlyReportedRevisions } from '../../packages/engine/src/sanitizer-postconditions';
import type { UnresolvedRevision } from '../../packages/engine/src/ooxml/revisions';

const P = 'word/document.xml';
const entry = (kind: string, revisionId?: string, part = P): UnresolvedRevision => ({ kind, count: 1, reason: 'r', part, ...(revisionId !== undefined ? { revisionId } : {}) });
const ins = (id?: string) => `<w:ins${id !== undefined ? ` w:id="${id}"` : ''} w:author="A">`;
const run = (xml: string, list: UnresolvedRevision[], part = P) => () => assertOnlyReportedRevisions(part, `<w:body>${xml}</w:body>`, list);

describe('identity level', () => {
  it('passes when the remaining element is the reported one', () => expect(run(ins('7'), [entry('ins', '7')])).not.toThrow());
  it('fails when a different revision of the same kind stayed and the reported one vanished', () => expect(run(ins('8'), [entry('ins', '7')])).toThrow(/w:id="8"/));
  it('fails when one reported and two remain', () => expect(run(ins('7') + ins('8'), [entry('ins', '7')])).toThrow(/w:id="8"/));
  it('a reported id is spent once: two elements with the same id and one entry fail', () => expect(run(ins('7') + ins('7'), [entry('ins', '7')])).toThrow());
  it('same id, different kind is a different revision', () => expect(run('<w:del w:id="7">', [entry('ins', '7')])).toThrow(/<w:del/));
  it('passes with nothing left and nothing reported', () => expect(run('<w:r/>', [])).not.toThrow());
});

describe('count level (bounded, limit documented)', () => {
  it('an element without an id is allowed up to the reported count for its kind', () => {
    expect(run(ins(), [entry('ins')])).not.toThrow();
    expect(run(ins() + ins(), [entry('ins')])).toThrow();
  });
  it('an element with an id is allowed against an entry that carries no id', () => expect(run(ins('9'), [entry('ins')])).not.toThrow());
  it('LIMIT: a reported revision that vanished while a different unreported one of the same kind stayed still balances (counts cannot tell them apart)', () => {
    // Both have no id, so the report cannot say which one; the check passes. This asserts the documented limit.
    expect(run(ins(), [entry('ins')])).not.toThrow();
  });
  it('an element without an id may use an id-carrying entry of that kind', () => expect(run(ins(), [entry('ins', '5')])).not.toThrow());
});

describe('part and kind scoping', () => {
  it('a reported revision in the body does not allow one in a header', () => {
    expect(run(ins('7'), [entry('ins', '7', P)], 'word/header1.xml')).toThrow(/word\/header1\.xml/);
  });
  it('a reported kind does not allow another kind', () => expect(run('<w:moveFrom w:id="1">', [entry('ins', '1')])).toThrow(/moveFrom/));
  it('the error names the part and the element and says nothing was written', () => {
    expect(run(ins('3'), [])).toThrow(/word\/document\.xml.*\(<w:ins w:id="3">\).*Nothing was written/);
  });
});

describe('revision element matching', () => {
  const none: never[] = [];
  it('an attribute-less self-closing tracked change is not allowed', () => {
    expect(() => assertOnlyReportedRevisions('word/document.xml', '<w:p><w:ins/></w:p>', none)).toThrow(/still has a tracked change/);
  });
  it('a single-quoted id is read as an id', () => {
    expect(() => assertOnlyReportedRevisions('word/document.xml', "<w:ins w:id='7' w:author='a'/>", [{ part: 'word/document.xml', kind: 'ins', count: 1, revisionId: '7' } as never])).not.toThrow();
  });
  it('a > inside an attribute value does not end the tag early', () => {
    expect(() => assertOnlyReportedRevisions('word/document.xml', '<w:ins w:author="a>b" w:id="9"/>', [{ part: 'word/document.xml', kind: 'ins', count: 1, revisionId: '9' } as never])).not.toThrow();
  });
  it('tracked-change text inside a comment or CDATA is not counted', () => {
    expect(() => assertOnlyReportedRevisions('word/document.xml', '<w:p><!-- <w:ins w:id="1"> --><![CDATA[<w:del w:id="2">]]></w:p>', none)).not.toThrow();
  });
  it('a longer element name that starts the same is not a tracked change', () => {
    expect(() => assertOnlyReportedRevisions('word/document.xml', '<w:insideText w:id="1"/>', none)).not.toThrow();
  });
});
