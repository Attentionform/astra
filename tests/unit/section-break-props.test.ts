import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  createBlankDocxBuffer,
  executeInsert,
  executePageSetup,
  executeHeaderFooter,
  validateOperationStrict,
  assertValidDocx,
  upsertSectPrChild,
} from '@attentionform/astra-engine';

describe('section break properties and page setup targeting', () => {
  it('copies page setup and header reference on section break while type moves to following section', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    executePageSetup({ action: 'pageSetup', orientation: 'landscape', paperSize: 'A4' }, ast, parser);
    executeHeaderFooter({ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'HDR' }, ast, parser);

    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'continuous' } as any, ast, parser);

    const sectPrs = parser.getSectPrElements();
    expect(sectPrs.length).toBe(2);

    for (const sp of sectPrs) {
      const pgSz = sp.getElementsByTagName('w:pgSz')[0];
      expect(pgSz).toBeDefined();
      expect(pgSz.getAttribute('w:w')).toBe('16838');
      expect(pgSz.getAttribute('w:h')).toBe('11906');
      expect(pgSz.getAttribute('w:orient')).toBe('landscape');
    }

    const breakSectPr = sectPrs[0];
    expect(breakSectPr.getElementsByTagName('w:headerReference').length).toBeGreaterThan(0);
    expect(breakSectPr.getElementsByTagName('w:type').length).toBe(0);

    const bodySectPr = sectPrs[1];
    expect(bodySectPr.getElementsByTagName('w:type').length).toBe(1);
    expect(bodySectPr.getElementsByTagName('w:type')[0].getAttribute('w:val')).toBe('continuous');
    expect(bodySectPr.getElementsByTagName('w:headerReference').length).toBe(0);

    executeHeaderFooter({ action: 'headerFooter', operation: 'set', targetType: 'header', text: 'HDR2', sectionIndex: 2 }, ast, parser);

    const zipAdapter = parser.getZipAdapter();
    const h1Xml = zipAdapter.readText('word/header1.xml');
    const h2Xml = zipAdapter.readText('word/header2.xml');
    expect(h1Xml).toContain('HDR');
    expect(h1Xml).not.toContain('HDR2');
    expect(h2Xml).toContain('HDR2');

    await assertValidDocx(parser.rezip());
  });

  it('keeps both sections landscape when pageSetup is applied after break', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'continuous' } as any, ast, parser);

    executePageSetup({ action: 'pageSetup', orientation: 'landscape', paperSize: 'A4' }, ast, parser);

    const sectPrs = parser.getSectPrElements();
    expect(sectPrs.length).toBe(2);
    for (const sp of sectPrs) {
      const pgSz = sp.getElementsByTagName('w:pgSz')[0];
      expect(pgSz).toBeDefined();
      expect(pgSz.getAttribute('w:w')).toBe('16838');
      expect(pgSz.getAttribute('w:h')).toBe('11906');
      expect(pgSz.getAttribute('w:orient')).toBe('landscape');
    }
  });

  it('preserves existing break type in copy when governing section already has a type', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const bodySectPr = parser.getSectPrElements()[0];
    upsertSectPrChild(bodySectPr, 'w:type').setAttribute('w:val', 'nextPage');

    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'continuous' } as any, ast, parser);

    const sectPrs = parser.getSectPrElements();
    expect(sectPrs.length).toBe(2);
    expect(sectPrs[0].getElementsByTagName('w:type')[0]?.getAttribute('w:val')).toBe('nextPage');
    expect(sectPrs[1].getElementsByTagName('w:type')[0]?.getAttribute('w:val')).toBe('continuous');
  });

  it('handles titlePg and pgNumType start removal on governing section', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const bodySectPr = parser.getSectPrElements()[0];
    upsertSectPrChild(bodySectPr, 'w:titlePg');
    const pgNumType = upsertSectPrChild(bodySectPr, 'w:pgNumType');
    pgNumType.setAttribute('w:start', '1');
    pgNumType.setAttribute('w:fmt', 'lowerRoman');

    const anchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor, position: 'after', breakType: 'nextPage' } as any, ast, parser);

    const sectPrs = parser.getSectPrElements();
    expect(sectPrs.length).toBe(2);

    const copySectPr = sectPrs[0];
    expect(copySectPr.getElementsByTagName('w:titlePg').length).toBe(1);
    const copyPgNumType = copySectPr.getElementsByTagName('w:pgNumType')[0];
    expect(copyPgNumType).toBeDefined();
    expect(copyPgNumType.getAttribute('w:start')).toBe('1');
    expect(copyPgNumType.getAttribute('w:fmt')).toBe('lowerRoman');

    const govSectPr = sectPrs[1];
    expect(govSectPr.getElementsByTagName('w:titlePg').length).toBe(0);
    const govPgNumType = govSectPr.getElementsByTagName('w:pgNumType')[0];
    expect(govPgNumType).toBeDefined();
    expect(govPgNumType.getAttribute('w:fmt')).toBe('lowerRoman');
    expect(govPgNumType.hasAttribute('w:start')).toBe(false);

    await assertValidDocx(parser.rezip());
  });

  it('copies body-level sectPr again on a second break inside the second section', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    executePageSetup({ action: 'pageSetup', orientation: 'landscape', paperSize: 'A4' }, ast, parser);

    const firstAnchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor: firstAnchor, position: 'after', breakType: 'nextPage' } as any, ast, parser);
    expect(parser.getSectPrElements().length).toBe(2);

    await executeInsert({ action: 'insert', type: 'paragraph', anchor: firstAnchor, position: 'after', text: 'Second section text' } as any, ast, parser);
    const secondSectionAnchor = ast.sections[1].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor: secondSectionAnchor, position: 'after', breakType: 'continuous' } as any, ast, parser);

    const sectPrs = parser.getSectPrElements();
    expect(sectPrs.length).toBe(3);
    for (const sp of sectPrs) {
      const pgSz = sp.getElementsByTagName('w:pgSz')[0];
      expect(pgSz).toBeDefined();
      expect(pgSz.getAttribute('w:w')).toBe('16838');
      expect(pgSz.getAttribute('w:h')).toBe('11906');
      expect(pgSz.getAttribute('w:orient')).toBe('landscape');
    }
  });

  it('supports sectionIndex targeting and out-of-range validation on pageSetup', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const firstAnchor = ast.sections[0].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor: firstAnchor, position: 'after', breakType: 'nextPage' } as any, ast, parser);
    await executeInsert({ action: 'insert', type: 'paragraph', anchor: firstAnchor, position: 'after', text: 'Sec2' } as any, ast, parser);
    const sec2Anchor = ast.sections[1].nodes[0].id;
    await executeInsert({ action: 'insert', type: 'sectionBreak', anchor: sec2Anchor, position: 'after', breakType: 'nextPage' } as any, ast, parser);

    const sectPrs = parser.getSectPrElements();
    expect(sectPrs.length).toBe(3);

    executePageSetup({ action: 'pageSetup', sectionIndex: 1, orientation: 'landscape' }, ast, parser);
    expect(sectPrs[0].getElementsByTagName('w:pgSz')[0]?.getAttribute('w:orient')).toBe('landscape');
    expect(sectPrs[1].getElementsByTagName('w:pgSz')[0]?.getAttribute('w:orient')).not.toBe('landscape');
    expect(sectPrs[2].getElementsByTagName('w:pgSz')[0]?.getAttribute('w:orient')).not.toBe('landscape');
    const size = (i: number) => {
      const pgSz = sectPrs[i].getElementsByTagName('w:pgSz')[0];
      return [pgSz.getAttribute('w:w'), pgSz.getAttribute('w:h')];
    };
    expect(size(0)).toEqual(['15840', '12240']);
    expect(size(1)).toEqual(['12240', '15840']);
    expect(size(2)).toEqual(['12240', '15840']);
    expect(ast.sections[0].pageSetup?.orientation).toBe('landscape');
    expect(ast.sections[1].pageSetup?.orientation).toBe('portrait');
    expect(ast.sections[2].pageSetup?.orientation).toBe('portrait');

    executePageSetup({ action: 'pageSetup', orientation: 'landscape' }, ast, parser);
    for (let i = 0; i < 3; i++) {
      expect(sectPrs[i].getElementsByTagName('w:pgSz')[0]?.getAttribute('w:orient')).toBe('landscape');
      expect(size(i)).toEqual(['15840', '12240']);
      expect(ast.sections[i].pageSetup?.orientation).toBe('landscape');
    }

    const bytesBefore = parser.rezip();
    expect(() => {
      executePageSetup({ action: 'pageSetup', sectionIndex: 9, orientation: 'portrait' }, ast, parser);
    }).toThrow(/E_VALIDATION.*out of range/);
    const bytesAfter = parser.rezip();
    expect(bytesBefore.equals(bytesAfter)).toBe(true);

    const flatValidation = validateOperationStrict({
      action: 'pageSetup',
      sectionIndex: 1,
      orientation: 'landscape',
    });
    expect(flatValidation.ok).toBe(true);

    const paramsValidation = validateOperationStrict({
      action: 'pageSetup',
      params: {
        sectionIndex: 1,
        orientation: 'landscape',
      },
    });
    expect(paramsValidation.ok).toBe(true);
  });
});
