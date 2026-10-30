/**
 * The multi-section fixtures are what the section tests stand on, so they are
 * checked on their own: the number of sections each file really has is counted
 * straight from its XML, without the parser, and each file must be a valid
 * document the Open XML SDK accepts.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import { DocxParser, assertValidDocx } from '@attentionform/astra-engine';
import { SECTION_SHAPES, countFileSections } from '../helpers/section-docs';

const FIXTURES = path.resolve(__dirname, '../fixtures/sections');

describe('section fixtures say what they are', () => {
  for (const [key, shape] of Object.entries(SECTION_SHAPES)) {
    it(`shape ${key}: ${shape.about}`, async () => {
      const docx = shape.build();
      expect(countFileSections(docx)).toBe(shape.sections);
      await assertValidDocx(docx);
    });
  }

  it('shapes are the same bytes on every build', () => {
    for (const shape of Object.values(SECTION_SHAPES)) {
      expect(shape.build().equals(shape.build())).toBe(true);
    }
  });

  it('python-docx-three-sections.docx has three sections and is valid', async () => {
    const docx = fs.readFileSync(path.join(FIXTURES, 'python-docx-three-sections.docx'));
    expect(countFileSections(docx)).toBe(3);
    await assertValidDocx(docx);
  });

  it('libreoffice-two-page-styles.docx has three sections; as written it has the LibreOffice child order, and ASTra\'s save repairs it', async () => {
    const docx = fs.readFileSync(path.join(FIXTURES, 'libreoffice-two-page-styles.docx'));
    expect(countFileSections(docx)).toBe(3);
    // LibreOffice writes w:sectPr before w:pStyle in w:pPr; Word accepts it, the strict schema does not.
    await expect(assertValidDocx(docx)).rejects.toThrow(/pStyle/);
    const saved = new DocxParser(docx).rezip();
    await assertValidDocx(saved);
    expect(countFileSections(saved)).toBe(3);
  });
});
