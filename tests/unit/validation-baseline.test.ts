import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import { buildBaseline, errorSignature, judgeAgainstBaseline, validateDocxBuffer, createBlankDocxBuffer, BASELINE_MAX_ERRORS } from '@attentionform/astra-engine';
import type { ValidationResult } from '@attentionform/astra-engine';

const err = (xPath: string, id = 'Sch_UnexpectedElementContentExpectingComplex', part = '/word/styles.xml', errorType = 'Schema') => ({
  description: `problem at ${xPath}`,
  id,
  part,
  path: part,
  xPath,
  errorType,
});

function result(errors: ReturnType<typeof err>[], extra: Record<string, unknown> = {}): ValidationResult {
  return { file: 'x.docx', ok: errors.length === 0, errors, ...extra } as unknown as ValidationResult;
}

describe('validation baseline', () => {
  it('asks for far more errors than the validator default of 200', () => {
    expect(BASELINE_MAX_ERRORS).toBeGreaterThan(1000);
  });

  it('a signature leaves out the indices of the xPath', () => {
    expect(errorSignature(err('/w:styles[1]/w:style[4]/w:tblPr[1]'))).toBe(
      '/word/styles.xml|Sch_ElementContent|/w:styles/w:style/w:tblPr',
    );
  });

  it('the three element-content errors at one element share a signature', () => {
    const at = '/w:styles[1]/w:docDefaults[1]/w:rPrDefault[1]/w:rPr[1]';
    const unexpected = errorSignature(err(at, 'Sch_UnexpectedElementContentExpectingComplex'));
    expect(errorSignature(err(at, 'Sch_InvalidElementContentExpectingComplex'))).toBe(unexpected);
    expect(errorSignature(err(at, 'Sch_IncompleteContentExpectingComplex'))).toBe(unexpected);
    // Another kind of error at the same element stays apart.
    expect(errorSignature(err(at, 'Sch_AttributeValueDataTypeDetailed'))).not.toBe(unexpected);
  });

  it('an element that had an unexpected-child error and now has an invalid-content error is not a new error', () => {
    // Moving the children into schema order clears the first error and shows the second one at the same element.
    const at = '/w:styles[1]/w:docDefaults[1]/w:rPrDefault[1]/w:rPr[1]';
    const base = buildBaseline(result([err(at, 'Sch_UnexpectedElementContentExpectingComplex')]));
    const after = result([err(at, 'Sch_InvalidElementContentExpectingComplex')]);
    expect(judgeAgainstBaseline(after, base)).toEqual({ kind: 'tolerated', preExisting: 1 });
  });

  it('a second content error at an element that had none is still new', () => {
    const base = buildBaseline(result([err('/w:styles[1]/w:style[1]/w:pPr[1]', 'Sch_UnexpectedElementContentExpectingComplex')]));
    const after = result([err('/w:styles[1]/w:style[1]/w:pPr[1]', 'Sch_InvalidElementContentExpectingComplex'), err('/w:styles[1]/w:style[2]/w:rPr[1]', 'Sch_InvalidElementContentExpectingComplex')]);
    expect(judgeAgainstBaseline(after, base).kind).toBe('introduced');
  });

  it('a valid result is clean', () => {
    const base = buildBaseline(result([err('/w:styles[1]/w:style[1]')]));
    expect(judgeAgainstBaseline(result([]), base)).toEqual({ kind: 'clean' });
  });

  it('a valid baseline is empty and counts zero', () => {
    expect(buildBaseline(result([]))).toEqual({ counts: {}, total: 0 });
  });

  it('the same errors are tolerated', () => {
    const errors = [err('/w:styles[1]/w:style[1]/w:rPr[1]'), err('/w:styles[1]/w:style[2]/w:rPr[1]')];
    expect(judgeAgainstBaseline(result(errors), buildBaseline(result(errors)))).toEqual({ kind: 'tolerated', preExisting: 2 });
  });

  it('fewer errors are tolerated', () => {
    const before = [err('/w:styles[1]/w:style[1]/w:rPr[1]'), err('/w:styles[1]/w:style[2]/w:rPr[1]')];
    expect(judgeAgainstBaseline(result(before.slice(0, 1)), buildBaseline(result(before)))).toEqual({ kind: 'tolerated', preExisting: 1 });
  });

  it('one more instance of a known signature is introduced', () => {
    const before = [err('/w:styles[1]/w:style[1]/w:rPr[1]')];
    const after = [err('/w:styles[1]/w:style[1]/w:rPr[1]'), err('/w:styles[1]/w:style[7]/w:rPr[1]')];
    const verdict = judgeAgainstBaseline(result(after), buildBaseline(result(before)));
    expect(verdict.kind).toBe('introduced');
    if (verdict.kind === 'introduced') {
      expect(verdict.introduced).toHaveLength(1);
      expect(verdict.preExisting).toBe(1);
    }
  });

  it('an error with a new signature is introduced', () => {
    const before = [err('/w:styles[1]/w:style[1]/w:rPr[1]')];
    const after = [...before, err('/w:numbering[1]/w:num[1]', 'Sch_UndeclaredAttribute', '/word/numbering.xml')];
    expect(judgeAgainstBaseline(result(after), buildBaseline(result(before))).kind).toBe('introduced');
  });

  it('a changed index in the xPath does not make an error new', () => {
    const before = [err('/w:document[1]/w:body[1]/w:p[3]/w:pPr[1]', 'Sch_InvalidElementContentExpectingComplex', '/word/document.xml')];
    const after = [err('/w:document[1]/w:body[1]/w:p[9]/w:pPr[1]', 'Sch_InvalidElementContentExpectingComplex', '/word/document.xml')];
    expect(judgeAgainstBaseline(result(after), buildBaseline(result(before))).kind).toBe('tolerated');
  });

  it('a cut-short result cannot be compared', () => {
    const cut = result([err('/w:styles[1]/w:style[1]')], { truncated: true });
    expect(buildBaseline(cut)).toBeNull();
    expect(judgeAgainstBaseline(cut, { counts: {}, total: 0 })).toEqual({ kind: 'incomparable', reason: 'truncated' });
  });

  it('an exception cannot be compared', () => {
    const boom = result([{ ...err('', 'x', '', 'Exception') }]);
    expect(buildBaseline(boom)).toBeNull();
    expect(judgeAgainstBaseline(boom, { counts: {}, total: 0 })).toEqual({ kind: 'incomparable', reason: 'exception' });
  });

  it('a duplicate paragraph id from the save guard is never tolerated', () => {
    const guard = result([{ ...err('/w:p[1]', 'DuplicateParaId', '/word/document.xml', 'DuplicateParaId') }]);
    expect(buildBaseline(guard)).toBeNull();
    expect(judgeAgainstBaseline(guard, buildBaseline(guard) ?? { counts: {}, total: 0 })).toEqual({ kind: 'incomparable', reason: 'guard' });
  });

  it('a missing baseline cannot be compared', () => {
    expect(judgeAgainstBaseline(result([err('/w:styles[1]')]), null)).toEqual({ kind: 'incomparable', reason: 'no-baseline' });
  });

  it('a document with more than 1000 errors gives a complete baseline, and an edit that adds none is tolerated', async () => {
    // The validator used to stop silently at 1000 errors, so a second run on an edited copy showed
    // a different 1000 and the difference looked like new errors.
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const xml = zip.readAsText('word/document.xml');
    zip.updateFile('word/document.xml', Buffer.from(xml.replace('<w:body>', '<w:body>' + '<w:p><w:bogus/></w:p>'.repeat(1200))));
    const before = await validateDocxBuffer(zip.toBuffer(), { maxErrors: BASELINE_MAX_ERRORS });
    expect(before.truncated).toBeFalsy();
    expect(before.errors.length).toBeGreaterThan(1000);
    const baseline = buildBaseline(before);
    expect(baseline).not.toBeNull();
    // The same errors after an edit that added a valid paragraph at the end.
    const edited = new AdmZip(zip.toBuffer());
    // A body ends with its section properties, so the paragraph goes in front of them.
    const body = edited.readAsText('word/document.xml');
    const at = body.lastIndexOf('<w:sectPr');
    expect(at).toBeGreaterThan(0);
    edited.updateFile('word/document.xml', Buffer.from(body.slice(0, at) + '<w:p><w:r><w:t>added</w:t></w:r></w:p>' + body.slice(at)));
    const after = await validateDocxBuffer(edited.toBuffer(), { maxErrors: BASELINE_MAX_ERRORS });
    expect(judgeAgainstBaseline(after, baseline).kind).toBe('tolerated');
  });
});
