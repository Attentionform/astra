/**
 * A list definition carries all nine levels, so an item placed at any level
 * finds its marker and indent: bullets cycle • ◦ ▪, numbers cycle 1. a. i. A
 * definition from an older file or from Word is reused only when its top level
 * is the requested kind, and gains the levels it lacks the first time an item
 * uses it.
 */
import { describe, it, expect } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { execFileSync } from 'child_process';
import AdmZip from 'adm-zip';
import { DOMParser } from '@xmldom/xmldom';
import {
  DocxParser,
  OperationEngine,
  NumberingManipulator,
  assertValidDocx,
  convertDocxToPdf,
  createBlankDocxBuffer,
  type DocumentAST,
} from '@attentionform/astra-engine';
import { withLoSlot } from '../helpers/lo-slot';

interface Fixture {
  parser: DocxParser;
  ast: DocumentAST;
  engine: OperationEngine;
}

const W = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';

/** Paragraphs li_1..li_4 in a blank document; `numberingXml` is a numbering part the file already had. */
async function paragraphs(numberingXml?: string): Promise<Fixture> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast = await parser.getAST();
  if (numberingXml) new NumberingManipulator(ast).parseNumbering(numberingXml);
  const engine = new OperationEngine(ast, parser);
  const seeded = await engine.executeBatch({
    operations: ['li_1', 'li_2', 'li_3', 'li_4'].map((id) => ({ action: 'insert', type: 'paragraph', anchor: 'end', text: `item ${id}`, id })) as never,
    sequential: true,
  });
  expect(seeded.success).toBe(true);
  return { parser, ast, engine };
}

const run = async (f: Fixture, op: Record<string, unknown>) => {
  const out = await f.engine.executeBatch({ operations: [op] as never });
  expect(out.success, JSON.stringify(out.errors)).toBe(true);
};

const create = (f: Fixture, target: string, listType: 'bullet' | 'numbered', level?: number) =>
  run(f, { action: 'list', operation: 'create', target, listType, ...(level ? { level } : {}) });

const nodeXml = (f: Fixture, id: string): Element =>
  f.ast.sections.flatMap((s) => s.nodes).find((n) => n.id === id)!.xml as Element;
const numIdOf = (f: Fixture, id: string): string | null => nodeXml(f, id).getElementsByTagName('w:numId')[0]?.getAttribute('w:val') ?? null;
const ilvlOf = (f: Fixture, id: string): string | null => nodeXml(f, id).getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val') ?? null;

/** The numbering part as saved. */
const savedNumbering = (f: Fixture): Document => {
  const xml = new AdmZip(f.parser.rezip()).readAsText('word/numbering.xml');
  expect(xml).toContain('<w:numbering');
  return new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
};

const kids = (el: Element, tag: string): Element[] =>
  Array.from(el.childNodes as unknown as Element[]).filter((c) => c.nodeType === 1 && c.tagName === tag);

const abstractFor = (doc: Document, numId: string): Element => {
  const num = Array.from(doc.getElementsByTagName('w:num')).find((n) => n.getAttribute('w:numId') === numId)!;
  expect(num, `w:num ${numId} exists`).toBeDefined();
  const absId = num.getElementsByTagName('w:abstractNumId')[0].getAttribute('w:val');
  return Array.from(doc.getElementsByTagName('w:abstractNum')).find((a) => a.getAttribute('w:abstractNumId') === absId)!;
};

interface Level {
  ilvl: string | null;
  fmt: string | null;
  text: string | null;
  left: string | null;
  hanging: string | null;
  font: string | null;
}
const levelsOf = (abs: Element): Level[] =>
  kids(abs, 'w:lvl').map((lvl) => ({
    ilvl: lvl.getAttribute('w:ilvl'),
    fmt: lvl.getElementsByTagName('w:numFmt')[0]?.getAttribute('w:val') ?? null,
    text: lvl.getElementsByTagName('w:lvlText')[0]?.getAttribute('w:val') ?? null,
    left: lvl.getElementsByTagName('w:ind')[0]?.getAttribute('w:left') ?? null,
    hanging: lvl.getElementsByTagName('w:ind')[0]?.getAttribute('w:hanging') ?? null,
    font: lvl.getElementsByTagName('w:rFonts')[0]?.getAttribute('w:ascii') ?? null,
  }));

const BULLETS = ['•', '◦', '▪'];
const FORMATS = ['decimal', 'lowerLetter', 'lowerRoman'];

describe('a new list definition has all nine levels', () => {
  it('bullets: • ◦ ▪ repeating, Arial, 720 twips in per level with a 360 hanging marker', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'bullet');
    const abs = abstractFor(savedNumbering(f), numIdOf(f, 'li_1')!);
    const levels = levelsOf(abs);
    expect(levels.map((l) => l.ilvl)).toEqual(['0', '1', '2', '3', '4', '5', '6', '7', '8']);
    expect(levels.map((l) => l.fmt)).toEqual(Array(9).fill('bullet'));
    expect(levels.map((l) => l.text)).toEqual(Array.from({ length: 9 }, (_, i) => BULLETS[i % 3]));
    expect(levels.map((l) => l.left)).toEqual(Array.from({ length: 9 }, (_, i) => String(720 * (i + 1))));
    expect(levels.map((l) => l.hanging)).toEqual(Array(9).fill('360'));
    expect(levels.map((l) => l.font)).toEqual(Array(9).fill('Arial'));
  });

  it('numbers: 1. a. i. repeating, counting %1. to %9.', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'numbered');
    const levels = levelsOf(abstractFor(savedNumbering(f), numIdOf(f, 'li_1')!));
    expect(levels.map((l) => l.ilvl)).toEqual(['0', '1', '2', '3', '4', '5', '6', '7', '8']);
    expect(levels.map((l) => l.fmt)).toEqual(Array.from({ length: 9 }, (_, i) => FORMATS[i % 3]));
    expect(levels.map((l) => l.text)).toEqual(Array.from({ length: 9 }, (_, i) => `%${i + 1}.`));
    expect(levels.map((l) => l.left)).toEqual(Array.from({ length: 9 }, (_, i) => String(720 * (i + 1))));
    expect(levels.map((l) => l.font)).toEqual(Array(9).fill(null));
  });

  it('is a hybrid multilevel definition with an identity, and no lvlRestart on any level', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'bullet');
    await create(f, 'li_2', 'numbered');
    const doc = savedNumbering(f);
    for (const id of ['li_1', 'li_2']) {
      const abs = abstractFor(doc, numIdOf(f, id)!);
      const order = Array.from(abs.childNodes as unknown as Element[]).filter((c) => c.nodeType === 1).map((c) => c.tagName);
      expect(order).toEqual(['w:nsid', 'w:multiLevelType', ...Array(9).fill('w:lvl')]);
      expect(kids(abs, 'w:multiLevelType')[0].getAttribute('w:val')).toBe('hybridMultilevel');
      expect(kids(abs, 'w:nsid')[0].getAttribute('w:val')).toMatch(/^[0-9A-F]{8}$/);
      expect(abs.getElementsByTagName('w:lvlRestart')).toHaveLength(0);
    }
  });

  it('every definition has its own identity, and the same request gives the same one', async () => {
    const nsidsOf = async (): Promise<string[]> => {
      const f = await paragraphs();
      await create(f, 'li_1', 'bullet');
      await create(f, 'li_2', 'numbered');
      return Array.from(savedNumbering(f).getElementsByTagName('w:nsid')).map((n) => n.getAttribute('w:val')!);
    };
    const first = await nsidsOf();
    expect(first).toHaveLength(2);
    expect(new Set(first).size).toBe(2);
    expect(await nsidsOf()).toEqual(first);
  });

  it('passes the Open XML SDK, with items at the top, the middle and the bottom level', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'bullet', 1);
    await create(f, 'li_2', 'bullet', 5);
    await create(f, 'li_3', 'numbered', 9);
    await assertValidDocx(f.parser.rezip());
  });

  it('an item at level 9 points at a level the definition has', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'numbered', 9);
    expect(ilvlOf(f, 'li_1')).toBe('8');
    const levels = levelsOf(abstractFor(savedNumbering(f), numIdOf(f, 'li_1')!));
    expect(levels.find((l) => l.ilvl === '8')).toMatchObject({ fmt: 'lowerRoman', text: '%9.', left: '6480' });
  });

  it('survives a save and reopen: same levels, items keep their level', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'bullet', 1);
    await create(f, 'li_2', 'bullet', 3);
    const reopened = new DocxParser(f.parser.rezip());
    const ast = await reopened.getAST();
    const abs = Object.values(ast.numbering.abstractNum) as Element[];
    expect(abs).toHaveLength(1);
    expect(kids(abs[0], 'w:lvl')).toHaveLength(9);
    const levels = ast.sections.flatMap((s) => s.nodes).filter((n) => /item li_[12]/.test(n.xml.textContent ?? ''))
      .map((n) => n.xml.getElementsByTagName('w:ilvl')[0]?.getAttribute('w:val'));
    expect(levels).toEqual(['0', '2']);
  });
});

describe('reuse looks at level 1 only', () => {
  /** What Word writes for a list that is numbered at level 1 and bulleted at level 2. */
  const wordStyle = (): string =>
    `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:nsid w:val="11111111"/><w:multiLevelType w:val="hybridMultilevel"/>` +
    `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>` +
    `<w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Courier New" w:hAnsi="Courier New"/></w:rPr></w:lvl>` +
    `</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

  it('a bullet list is not given a list that is numbered at level 1 and has a bullet at level 2', async () => {
    const f = await paragraphs(wordStyle());
    await create(f, 'li_1', 'bullet');
    await create(f, 'li_2', 'bullet');
    expect(numIdOf(f, 'li_1')).not.toBe('1');
    expect(numIdOf(f, 'li_1')).toBe(numIdOf(f, 'li_2'));
    const abs = abstractFor(savedNumbering(f), numIdOf(f, 'li_1')!);
    expect(levelsOf(abs)[0].fmt).toBe('bullet');
  });

  it('a numbered list reuses that definition', async () => {
    const f = await paragraphs(wordStyle());
    await create(f, 'li_1', 'numbered');
    expect(numIdOf(f, 'li_1')).toBe('1');
  });

  it('a level 1 with no marker is not a numbered list', async () => {
    const f = await paragraphs(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/>` +
        `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="none"/><w:lvlText w:val=""/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    );
    await create(f, 'li_1', 'numbered');
    expect(numIdOf(f, 'li_1')).not.toBe('1');
  });

  it('heading numbering is not handed to list items', async () => {
    const f = await paragraphs(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>` +
        `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    );
    await create(f, 'li_1', 'numbered');
    expect(numIdOf(f, 'li_1')).not.toBe('1');
  });

  it('a list that is a style link is not handed out either', async () => {
    const f = await paragraphs(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="hybridMultilevel"/><w:numStyleLink w:val="MyList"/>` +
        `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    );
    await create(f, 'li_1', 'numbered');
    expect(numIdOf(f, 'li_1')).not.toBe('1');
  });
});

describe('a definition that stops short gains the levels it lacks, in place', () => {
  /** What earlier versions wrote: one level, a lvlRestart on it, no nsid and no multiLevelType. */
  const oneLevel = (): string =>
    `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/>` +
    `<w:lvlRestart w:val="1"/><w:lvlText w:val="&#8226;"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>` +
    `<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial"/></w:rPr></w:lvl></w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;

  it('create at level 4 reuses the list (same numId), fills levels 2-9, and keeps level 1 as it was', async () => {
    const f = await paragraphs(oneLevel());
    await create(f, 'li_1', 'bullet', 4);
    expect(numIdOf(f, 'li_1')).toBe('1');
    const abs = abstractFor(savedNumbering(f), '1');
    const levels = levelsOf(abs);
    expect(levels.map((l) => l.ilvl)).toEqual(['0', '1', '2', '3', '4', '5', '6', '7', '8']);
    expect(levels.map((l) => l.text)).toEqual(Array.from({ length: 9 }, (_, i) => BULLETS[i % 3]));
    // Level 1 is the original, untouched, including its restart.
    expect(kids(abs, 'w:lvl')[0].getElementsByTagName('w:lvlRestart')).toHaveLength(1);
    // It is multilevel now, and the type sits first (there is no nsid to follow).
    const order = Array.from(abs.childNodes as unknown as Element[]).filter((c) => c.nodeType === 1).map((c) => c.tagName);
    expect(order[0]).toBe('w:multiLevelType');
    expect(kids(abs, 'w:multiLevelType')[0].getAttribute('w:val')).toBe('hybridMultilevel');
    await assertValidDocx(f.parser.rezip());
  });

  it('a numbered definition with one level gets numbered levels', async () => {
    const f = await paragraphs(oneLevel().replace('w:val="bullet"', 'w:val="decimal"').replace('&#8226;', '%1.'));
    await create(f, 'li_1', 'numbered', 3);
    expect(numIdOf(f, 'li_1')).toBe('1');
    const levels = levelsOf(abstractFor(savedNumbering(f), '1'));
    expect(levels.map((l) => l.fmt)).toEqual(Array.from({ length: 9 }, (_, i) => FORMATS[i % 3]));
  });

  /** An item already in list 1, as a file from an earlier version has it: the definition is back to one level. */
  const itemInShortList = async (): Promise<Fixture> => {
    const f = await paragraphs(oneLevel());
    await create(f, 'li_1', 'bullet', 1);
    new NumberingManipulator(f.ast).parseNumbering(oneLevel());
    expect(levelsOf(abstractFor(savedNumbering(f), '1'))).toHaveLength(1);
    return f;
  };

  it('changeLevel to a level the list lacks upgrades it too', async () => {
    const f = await itemInShortList();
    await run(f, { action: 'list', operation: 'changeLevel', target: 'li_1', level: 6 });
    expect(ilvlOf(f, 'li_1')).toBe('5');
    expect(levelsOf(abstractFor(savedNumbering(f), '1'))).toHaveLength(9);
  });

  it('continue into a short list upgrades it', async () => {
    const f = await itemInShortList();
    await run(f, { action: 'list', operation: 'continue', target: 'li_2', previousNodeId: 'li_1', level: 7 });
    expect(numIdOf(f, 'li_2')).toBe('1');
    expect(levelsOf(abstractFor(savedNumbering(f), '1'))).toHaveLength(9);
  });

  it('a short Word list keeps its levels and its identity; only the missing ones are added in order', async () => {
    const lvl = (i: number, fmt: string, text: string) =>
      `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmt}"/><w:lvlText w:val="${text}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
    // Levels 1 and 3 exist; level 2 and 4-9 do not. A custom marker at level 3.
    const f = await paragraphs(
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:nsid w:val="ABCDEF01"/><w:multiLevelType w:val="hybridMultilevel"/>` +
        `${lvl(0, 'decimal', '%1.')}${lvl(2, 'upperRoman', '(%3)')}</w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`,
    );
    await create(f, 'li_1', 'numbered', 2);
    const abs = abstractFor(savedNumbering(f), '1');
    expect(kids(abs, 'w:nsid')[0].getAttribute('w:val')).toBe('ABCDEF01');
    const levels = levelsOf(abs);
    expect(levels.map((l) => l.ilvl)).toEqual(['0', '1', '2', '3', '4', '5', '6', '7', '8']);
    expect(levels[2]).toMatchObject({ fmt: 'upperRoman', text: '(%3)' });
    expect(levels[1]).toMatchObject({ fmt: 'lowerLetter', text: '%2.' });
    await assertValidDocx(f.parser.rezip());
  });

  it('a single-level definition becomes multilevel', async () => {
    const f = await paragraphs(oneLevel().replace('<w:lvl ', '<w:multiLevelType w:val="singleLevel"/><w:lvl '));
    await create(f, 'li_1', 'bullet', 2);
    expect(kids(abstractFor(savedNumbering(f), '1'), 'w:multiLevelType')[0].getAttribute('w:val')).toBe('hybridMultilevel');
  });

  it('a level supplied by a w:lvlOverride on the num is not added twice', async () => {
    const f = await paragraphs(
      oneLevel().replace(
        '<w:abstractNumId w:val="0"/></w:num>',
        '<w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="1"><w:lvl w:ilvl="1"><w:start w:val="5"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%2)"/><w:lvlJc w:val="left"/></w:lvl></w:lvlOverride></w:num>',
      ),
    );
    await create(f, 'li_1', 'bullet', 2);
    const levels = levelsOf(abstractFor(savedNumbering(f), '1'));
    expect(levels.map((l) => l.ilvl)).toEqual(['0', '2', '3', '4', '5', '6', '7', '8']);
  });

  it('a list that already has all nine levels is not touched', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'bullet');
    const before = new AdmZip(f.parser.rezip()).readAsText('word/numbering.xml');
    await create(f, 'li_2', 'bullet', 9);
    expect(new AdmZip(f.parser.rezip()).readAsText('word/numbering.xml')).toBe(before);
  });

  it('heading numbering is never extended, even when an item is moved into it', async () => {
    const headingNumbering =
      `<w:numbering ${W}><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/>` +
      `<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="Heading1"/><w:lvlText w:val="%1"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>`;
    const f = await paragraphs(headingNumbering);
    new NumberingManipulator(f.ast).ensureLevels('1');
    expect(levelsOf(abstractFor(savedNumbering(f), '1'))).toHaveLength(1);
  });

  it('an unknown numId is left alone', async () => {
    const f = await paragraphs();
    const manipulator = new NumberingManipulator(f.ast);
    expect(() => manipulator.ensureLevels('99')).not.toThrow();
    expect(Object.keys(f.ast.numbering.abstractNum)).toHaveLength(0);
  });
});

describe('the fallback writer gives the same nine levels', () => {
  it('writes a complete definition when a stored one cannot be serialized', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const manipulator = new NumberingManipulator(ast);
    manipulator.getOrCreateList('bullet');
    manipulator.getOrCreateList('numbered');
    // Not serializable as a DOM element, but its format is readable.
    for (const [id, fmt] of [['0', 'bullet'], ['1', 'decimal']] as const) {
      ast.numbering.abstractNum[id] = { getElementsByTagName: () => [{ getAttribute: () => fmt }] };
    }
    const xml = manipulator.serializeNumbering()!;
    const doc = new DOMParser().parseFromString(xml, 'text/xml') as unknown as Document;
    const [bullets, numbers] = Array.from(doc.getElementsByTagName('w:abstractNum'));
    expect(levelsOf(bullets).map((l) => l.text)).toEqual(Array.from({ length: 9 }, (_, i) => BULLETS[i % 3]));
    expect(levelsOf(numbers).map((l) => l.fmt)).toEqual(Array.from({ length: 9 }, (_, i) => FORMATS[i % 3]));
    expect(levelsOf(numbers).map((l) => l.text)).toEqual(Array.from({ length: 9 }, (_, i) => `%${i + 1}.`));
    expect(kids(bullets, 'w:multiLevelType')[0].getAttribute('w:val')).toBe('hybridMultilevel');
    expect(xml).not.toContain('lvlRestart');
  });
});

// What a reader sees: opt-in like the other LibreOffice probes.
const RUN_LIVE = process.env.ASTRA_RUN_SLOW === '1';

describe.skipIf(!RUN_LIVE)('LibreOffice draws the markers at every level', () => {
  const pdfText = async (f: Fixture): Promise<string> => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-list-'));
    try {
      const pdf = path.join(dir, 'list.pdf');
      await withLoSlot(() => convertDocxToPdf(f.parser.rezip(), pdf));
      return execFileSync('pdftotext', ['-layout', pdf, '-'], { encoding: 'utf8' });
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  };

  it('bullets show • ◦ ▪ and numbers show 1. a. i.', async () => {
    const f = await paragraphs();
    await create(f, 'li_1', 'bullet', 1);
    await create(f, 'li_2', 'bullet', 2);
    await create(f, 'li_3', 'bullet', 3);
    await create(f, 'li_4', 'numbered', 1);
    await run(f, { action: 'insert', type: 'paragraph', anchor: 'li_4', position: 'after', text: 'sub b', id: 'sub_b' });
    await run(f, { action: 'insert', type: 'paragraph', anchor: 'sub_b', position: 'after', text: 'sub c', id: 'sub_c' });
    await run(f, { action: 'list', operation: 'continue', target: 'sub_b', previousNodeId: 'li_4', level: 2 });
    await run(f, { action: 'list', operation: 'continue', target: 'sub_c', previousNodeId: 'sub_b', level: 3 });
    const text = await pdfText(f);
    const line = (needle: string) => text.split('\n').find((l) => l.includes(needle)) ?? '';
    expect(line('item li_1')).toContain('•');
    expect(line('item li_2')).toContain('◦');
    expect(line('item li_3')).toContain('▪');
    expect(line('item li_4')).toMatch(/1\./);
    expect(line('sub b')).toMatch(/a\./);
    expect(line('sub c')).toMatch(/i\./);
    // Each level sits further in than the one above.
    const indent = (needle: string) => line(needle).search(/\S/);
    expect(indent('item li_2')).toBeGreaterThan(indent('item li_1'));
    expect(indent('item li_3')).toBeGreaterThan(indent('item li_2'));
  }, 300_000);
});
