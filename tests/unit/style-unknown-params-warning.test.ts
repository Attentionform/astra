import { describe, it, expect } from 'vitest';
import { createBlankDocxBuffer, DocxParser, OperationEngine } from '@attentionform/astra-engine';

/**
 * A key inside `params` or `properties` that the style op does not read used to be dropped without a word
 * when another key in the same bag was valid. The op still applies the keys it knows; it now reports the rest.
 */
async function freshEngine() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  const target = ast.sections[0].nodes[0].id;
  return { engine: new OperationEngine(ast, parser), target };
}

describe('style op: unknown key next to a valid key', () => {
  it('applies the valid key and warns about "spaceAfter" with the key it probably meant', async () => {
    const { engine, target } = await freshEngine();
    const { warnings } = await engine.execute({ action: 'style', target, params: { bold: true, spaceAfter: 120 } } as never);
    expect(warnings).toEqual(['style: "spaceAfter" is not a style key, so it was ignored. Did you mean "spacingAfter"?']);
  });

  it('warns inside properties too', async () => {
    const { engine, target } = await freshEngine();
    const { warnings } = await engine.execute({ action: 'style', target, properties: { italic: true, textcolour: 'FF0000' } } as never);
    expect(warnings).toHaveLength(1);
    expect(warnings[0]).toContain('"textcolour" is not a style key');
  });

  it('warns without a guess when nothing is close, and names every unknown key once', async () => {
    const { engine, target } = await freshEngine();
    const { warnings } = await engine.execute({ action: 'style', target, params: { bold: true, zzzz: 1 }, properties: { zzzz: 2, qqqq: 3 } } as never);
    expect(warnings).toEqual([
      'style: "zzzz" is not a style key, so it was ignored.',
      'style: "qqqq" is not a style key, so it was ignored.',
    ]);
  });

  it('is silent when every key is a style key', async () => {
    const { engine, target } = await freshEngine();
    const { warnings } = await engine.execute({ action: 'style', target, params: { bold: true, spacingAfter: 120 }, color: 'FF0000' } as never);
    expect(warnings).toEqual([]);
  });

  it('still refuses when no key applies, and still names the guess', async () => {
    const { engine, target } = await freshEngine();
    await expect(engine.execute({ action: 'style', target, params: { spaceAfter: 120 } } as never)).rejects.toThrow(/Did you mean "spacingAfter"/);
  });
});
