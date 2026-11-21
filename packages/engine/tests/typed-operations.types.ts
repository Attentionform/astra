/**
 * Compile-time contract of the typed `executeBatch` entry point. This file has no runtime test:
 * `npx tsc --noEmit` checks it, and an `@ts-expect-error` that stops erroring fails the build.
 *
 * What the compiler rejects: an unknown `action`, an unknown field, a missing field that every variant
 * requires, a wrong value for a closed set. What it cannot reject, and the runtime checks instead
 * (or `parseBatch()` earlier): fields required only for a chosen sub-operation, rules that span fields,
 * and node ids that do not exist.
 */
import { DocxParser } from '../src/parser';
import { OperationEngine } from '../src/operation-engine';
import { parseBatch, type OperationInput } from '../src/operation-schema';

export async function typedBatchContract(): Promise<void> {
  const parser = DocxParser.createNewDocument('blank');
  const engine = new OperationEngine(await parser.getAST(), parser);

  // --- Forms that compile -------------------------------------------------------------
  await engine.executeBatch({
    operations: [
      { action: 'replace', target: 'p_1', text: 'new text' },
      // A heading: insert a paragraph with an id, then apply the named style to that id.
      { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Chapter one', id: 'h_one' },
      { action: 'namedStyle', operation: 'apply', styleName: 'Heading 1', target: 'h_one' },
      { action: 'insert', type: 'table', anchor: 'end', position: 'after', headers: ['A', 'B'], rows: [['1', '2']] },
      { action: 'delete', target: 'p_2' },
      { action: 'style', target: 'p_1', properties: { bold: true } },
      { action: 'comment', operation: 'insert', target: 'p_1', text: 'check this' },
    ],
    sequential: true,
    rollbackOnError: true,
  });
  const fromOperationInput: OperationInput = { action: 'replace', target: 'p_1', text: 'x' };
  await engine.executeBatch({ operations: [fromOperationInput] });

  // --- Forms the compiler rejects -----------------------------------------------------
  await engine.executeBatch({
    operations: [
      // @ts-expect-error unknown action
      { action: 'frobnicate', target: 'p_1' },
    ],
  });
  await engine.executeBatch({
    operations: [
      // @ts-expect-error replace needs target and text
      { action: 'replace' },
    ],
  });
  await engine.executeBatch({
    operations: [
      // @ts-expect-error an unknown field is not allowed
      { action: 'replace', target: 'p_1', text: 'x', colour: 'red' },
    ],
  });
  await engine.executeBatch({
    operations: [
      // @ts-expect-error insert has no `style` field (the engine used to ignore it without a word)
      { action: 'insert', type: 'paragraph', anchor: 'end', position: 'after', text: 'Title', style: 'Heading1' },
    ],
  });
  await engine.executeBatch({
    operations: [
      // @ts-expect-error 'heading' is not an insert type: use paragraph with style Heading1
      { action: 'insert', type: 'heading', anchor: 'end', position: 'after', text: 'Title' },
    ],
  });
  await engine.executeBatch({
    operations: [
      // @ts-expect-error position is 'before' or 'after'
      { action: 'insert', type: 'paragraph', anchor: 'end', position: 'middle', text: 'x' },
    ],
  });
  // @ts-expect-error operations is required
  await engine.executeBatch({ sequential: true });

  // --- Untyped data goes through parseBatch, which returns a typed batch or a list of issues ---
  const input: unknown = JSON.parse('{"operations":[{"action":"replace","target":"p_1","text":"x"}]}');
  const parsed = parseBatch(input);
  if (parsed.ok) await engine.executeBatch(parsed.batch);
  // @ts-expect-error unknown data does not compile without parseBatch
  await engine.executeBatch(input);
}
