import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine } from '../src/index';

const TEMPLATES = ['blank', 'memo', 'report', 'letter', 'proposal', 'invoice', 'research', 'corporate', 'financial', 'novel', 'academic', 'newsletter', 'api-spec'];

const cover = { action: 'coverPage', template: 'modern_minimalist', title: 'Annual Report', subtitle: 'Sub', author: 'A', organization: 'O' };

async function open(template: string) {
  const parser = DocxParser.createNewDocument(template);
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const allText = (ast: { sections: Array<{ nodes: Array<{ metadata: { text?: string } }> }> }) =>
  ast.sections.flatMap((s) => s.nodes).map((n) => n.metadata.text ?? '').join('\n');

describe('a cover page replaces the placeholder content of a fresh template', () => {
  for (const template of TEMPLATES) {
    it(`${template}: nothing of the template text is left under the cover`, async () => {
      const { ast, engine } = await open(template);
      const before = ast.sections.flatMap((s) => s.nodes);
      expect(before.length).toBeGreaterThan(0);
      const seedTexts = before.map((n) => (n.metadata.text ?? '').trim()).filter((t) => t.length > 3);
      const out = await engine.executeBatch({ operations: [cover] as never });
      expect(out.success, JSON.stringify(out.errors)).toBe(true);
      const after = allText(ast);
      expect(after).toContain('Annual Report');
      for (const text of seedTexts) expect(after, `left behind: ${text}`).not.toContain(text);
      expect(ast.sections.flatMap((s) => s.nodes).every((n) => n.id.startsWith('cover_'))).toBe(true);
    });
  }

  it('a document that was edited first keeps all of its content', async () => {
    const { ast, engine } = await open('blank');
    const first = ast.sections[0].nodes[0].id;
    await engine.executeBatch({ operations: [{ action: 'insert', type: 'paragraph', anchor: first, position: 'after', text: 'My own paragraph' }] as never });
    const out = await engine.executeBatch({ operations: [cover] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    const text = allText(ast);
    expect(text).toContain('My own paragraph');
    expect(text).toContain('Blank Document');
  });

  it('a document opened from bytes keeps all of its content', async () => {
    const original = DocxParser.createNewDocument('blank');
    const reopened = new DocxParser(Buffer.from(original.rezip()));
    const ast = await reopened.getAST();
    const out = await new OperationEngine(ast, reopened).executeBatch({ operations: [cover] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(allText(ast)).toContain('Blank Document');
  });
});
