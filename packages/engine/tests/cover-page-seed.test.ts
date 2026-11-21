import { describe, it, expect } from 'vitest';
import { DocxParser, OperationEngine } from '../src/index';

const cover = { action: 'coverPage', template: 'modern_minimalist', title: 'Annual Report', subtitle: 'Sub', author: 'A', organization: 'O' };

async function open(template: string) {
  const parser = DocxParser.createNewDocument(template);
  const ast = await parser.getAST();
  return { parser, ast, engine: new OperationEngine(ast, parser) };
}
const allText = (ast: { sections: Array<{ nodes: Array<{ metadata: { text?: string } }> }> }) =>
  ast.sections.flatMap((s) => s.nodes).map((n) => n.metadata.text ?? '').join('\n');

describe('a cover page replaces the content of a fresh blank document only', () => {
  it('blank: the empty paragraph is replaced, so only the cover remains', async () => {
    const { ast, engine } = await open('blank');
    expect(ast.sections.flatMap((s) => s.nodes)).toHaveLength(1);
    const out = await engine.executeBatch({ operations: [cover] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(allText(ast)).toContain('Annual Report');
    expect(ast.sections.flatMap((s) => s.nodes).every((n) => n.id.startsWith('cover_'))).toBe(true);
  });

  for (const template of ['memo', 'report', 'letter', 'proposal', 'invoice', 'technical-spec', 'research-paper', 'novel', 'newsletter']) {
    it(`${template}: a designed template keeps its own content (the caller asked for a cover, not for the template to go)`, async () => {
      const { ast, engine } = await open(template);
      const before = allText(ast);
      const out = await engine.executeBatch({ operations: [cover] as never });
      expect(out.success, JSON.stringify(out.errors)).toBe(true);
      const after = allText(ast);
      expect(after).toContain('Annual Report');
      for (const line of before.split('\n').filter((l) => l.trim().length > 3)) expect(after, `lost: ${line}`).toContain(line);
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
  });

  it('a report opened from bytes keeps all of its content', async () => {
    const original = DocxParser.createNewDocument('report');
    const reopened = new DocxParser(Buffer.from(original.rezip()));
    const ast = await reopened.getAST();
    const out = await new OperationEngine(ast, reopened).executeBatch({ operations: [cover] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(allText(ast)).toContain('Summarise the purpose');
  });

  it('a blank file opened from bytes is still one empty paragraph: the cover replaces it', async () => {
    const reopened = new DocxParser(Buffer.from(DocxParser.createNewDocument('blank').rezip()));
    const ast = await reopened.getAST();
    const out = await new OperationEngine(ast, reopened).executeBatch({ operations: [cover] as never });
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(ast.sections.flatMap((s) => s.nodes).every((n) => n.id.startsWith('cover_'))).toBe(true);
  });
});
