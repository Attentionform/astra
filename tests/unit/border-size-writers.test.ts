import { describe, it, expect } from 'vitest';
import { XMLSerializer } from '@xmldom/xmldom';
import { DocxParser, createBlankDocxBuffer, executeInsert, executeTable, executeStyle, executePageBorder, validateOperationStrict } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../helpers/reproducible';

/**
 * A border width is a whole number of eighths of a point from 1 to 96. The schema checks the top-level `sz`; the
 * same rule is enforced where the value is written, so the nested `params` of a table operation and the border
 * configs of a style operation (both free-form in the schema) cannot write a width the document cannot hold.
 */
useReproducibleOutput();

async function fresh() {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const anchor = ast.sections[0].nodes[0].id as string;
  await executeInsert({ action: 'insert', type: 'table', anchor, position: 'after', headers: ['A', 'B'], rows: [['1', '2']] } as any, ast, parser);
  const tableId = ast.sections[0].nodes.find((n: any) => n.type === 'table').id as string;
  return { parser, ast, tableId, anchor };
}
const markup = (ast: any, id: string): string => new XMLSerializer().serializeToString(ast.sections[0].nodes.find((n: any) => n.id === id).xml);

async function rejected(run: () => unknown): Promise<string> {
  try {
    await run();
  } catch (error) {
    return String((error as Error).message);
  }
  return '';
}

describe('table setBorder width in nested params', () => {
  for (const sz of [1000, 97, 0, -4, 2.5, 'wide']) {
    it(`params.sz ${JSON.stringify(sz)} is refused`, async () => {
      const { ast, tableId } = await fresh();
      const op = { action: 'table', target: tableId, operation: 'setBorder', params: { sz } };
      expect(() => validateOperationStrict(op as any)).not.toThrow(); // the schema cannot see inside params
      const message = await rejected(() => executeTable(op as any, ast));
      expect(message, String(sz)).toContain('border width');
    });
  }

  it('the old size key is checked the same way', async () => {
    const { ast, tableId } = await fresh();
    const message = await rejected(() => executeTable({ action: 'table', target: tableId, operation: 'setBorder', params: { size: 1000 } } as any, ast));
    expect(message).toContain('border width');
  });

  it('96 and 1 are accepted and written', async () => {
    for (const sz of [1, 96, '8']) {
      const { ast, tableId } = await fresh();
      await executeTable({ action: 'table', target: tableId, operation: 'setBorder', params: { sz } } as any, ast);
      expect(markup(ast, tableId)).toContain(`w:sz="${Number(sz)}"`);
    }
  });
});

describe('page border and style border configs', () => {
  it('pageBorder refuses a width outside 1 to 96 when called without the schema', async () => {
    for (const sz of [1000, 97, 0, 2.5]) {
      const { parser, ast } = await fresh();
      const message = await rejected(() => executePageBorder({ action: 'pageBorder', sz } as any, ast, parser));
      expect(message, String(sz)).toContain('border width');
    }
  });

  it('a style border config with a width outside 1 to 96 is refused; the schema lets it through', async () => {
    for (const sz of [1000, 0, 2.5, 'wide']) {
      const { ast, anchor } = await fresh();
      const op = { action: 'style', target: anchor, properties: { borderBox: { style: 'single', sz } } };
      expect(() => validateOperationStrict(op as any)).not.toThrow();
      const message = await rejected(() => executeStyle(op as any, ast));
      expect(message, String(sz)).toContain('border width');
    }
  });

  it('a style border config without sz keeps its default width', async () => {
    const { ast, anchor } = await fresh();
    executeStyle({ action: 'style', target: anchor, properties: { borderBox: { style: 'single' } } } as any, ast);
    expect(markup(ast, anchor)).toContain('w:sz="12"');
  });
});
