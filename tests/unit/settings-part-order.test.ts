/**
 * New documents carry word/settings.xml from the start, so everything that
 * adds to it has to keep the schema's child order: `updateFields` before the
 * note properties and `compat`, the note properties before `compat`.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  DocxParser,
  OperationEngine,
  TOCManipulator,
  createBlankDocxBuffer,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

function settingsChildren(buffer: Buffer): string[] {
  const xml = new AdmZip(buffer).readAsText('word/settings.xml');
  return [...xml.replace(/<\?xml[^>]*\?>/, '').matchAll(/<(w:[A-Za-z]+|m:[A-Za-z]+)[\s/>]/g)]
    .map((m) => m[1])
    .filter((name, i, all) => name !== 'w:settings' && !/^w:(compatSetting)$/.test(name) && all.indexOf(name) === i);
}

describe('settings.xml child order in a new document', () => {
  it('places updateFields before compat', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    TOCManipulator.enableWordAutoUpdateFields(parser);
    const children = settingsChildren(parser.rezip());
    expect(children).toContain('w:updateFields');
    expect(children.indexOf('w:updateFields')).toBeLessThan(children.indexOf('w:compat'));
    expect(children.indexOf('w:characterSpacingControl')).toBeLessThan(children.indexOf('w:updateFields'));
  });

  it('places the note properties after updateFields and before compat once a footnote and an endnote exist', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    TOCManipulator.enableWordAutoUpdateFields(parser);
    const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const anchor = ast.sections[0].nodes[0].id;
    const result = await engine.executeBatch({
      operations: [
        { action: 'footnote', anchor, text: 'A footnote.' },
        { action: 'endnote', anchor, text: 'An endnote.' },
      ] as never,
      sequential: true,
    });
    expect(result.success).toBe(true);
    const buffer = parser.rezip(ast);
    const children = settingsChildren(buffer);
    const at = (tag: string) => children.indexOf(tag);
    expect(at('w:updateFields')).toBeLessThan(at('w:footnotePr'));
    expect(at('w:footnotePr')).toBeLessThan(at('w:endnotePr'));
    expect(at('w:endnotePr')).toBeLessThan(at('w:compat'));
    expect(children.filter((c) => c === 'w:footnotePr')).toHaveLength(1);
    expect((await validateDocxBuffer(buffer)).ok).toBe(true);
  });
});
