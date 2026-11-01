import { describe, it, expect } from 'vitest';
import { buildBaseline, errorSignature, judgeAgainstBaseline, BASELINE_MAX_ERRORS } from '@attentionform/astra-engine';
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
      '/word/styles.xml|Sch_UnexpectedElementContentExpectingComplex|/w:styles/w:style/w:tblPr',
    );
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
});
