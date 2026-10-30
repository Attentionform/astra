/**
 * Batch 9b — clear-formatting backend op.
 *
 * Pins executeClearFormat: fully-styled runs strip to default, pStyle resets
 * to Normal, spacing overrides drop, partial ranges leave siblings untouched,
 * empty/missing targets behave, and the result round-trips through rezip.
 */
import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import {
  createBlankDocxBuffer,
  DocxParser,
  executeClearFormat,
  executeNamedStyle,
  executeStyle,
  OperationEngine,
  RPR_CHILD_ORDER,
  validateDocxBuffer,
} from '@attentionform/astra-engine';

const STRIPPED = [
  'w:b', 'w:i', 'w:color', 'w:highlight', 'w:u',
  'w:strike', 'w:dstrike', 'w:sz', 'w:szCs', 'w:rFonts', 'w:vertAlign',
];

function runPropTags(paragraphEl: Element): string[] {
  const found = new Set<string>();
  const runs = paragraphEl.getElementsByTagName('w:r');
  for (let i = 0; i < runs.length; i++) {
    const rPr = runs[i].getElementsByTagName('w:rPr')[0];
    if (!rPr) continue;
    for (let j = 0; j < rPr.childNodes.length; j++) {
      const c = rPr.childNodes[j] as Element;
      if ((c as any).nodeType === 1) found.add((c as Element).tagName);
    }
  }
  return [...found];
}

function pStyleVal(paragraphEl: Element): string | null {
  const pPr = paragraphEl.getElementsByTagName('w:pPr')[0];
  if (!pPr) return null;
  const ps = pPr.getElementsByTagName('w:pStyle')[0];
  return ps ? ps.getAttribute('w:val') : null;
}

function hasSpacing(paragraphEl: Element): boolean {
  const pPr = paragraphEl.getElementsByTagName('w:pPr')[0];
  if (!pPr) return false;
  return pPr.getElementsByTagName('w:spacing').length > 0;
}

async function styledFixture(): Promise<{ parser: DocxParser; ast: any; id: string; text: string }> {
  const parser = new DocxParser(createBlankDocxBuffer('blank'));
  const ast: any = await parser.getAST();
  const node = ast.sections[0].nodes[0];
  const id = node.id as string;
  const text = node.metadata.text as string;
  executeNamedStyle(
    { action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: id } as any,
    ast, parser
  );
  executeStyle(
    {
      action: 'style', target: id,
      properties: {
        bold: true, italic: true, underline: true, strikethrough: true,
        vertAlign: 'superscript', color: 'FF0000', fontSize: 24,
        fontFamily: 'Arial', highlight: 'yellow',
        alignment: 'center', lineSpacing: 480, spacingBefore: 240, spacingAfter: 120,
      },
    } as any,
    ast
  );
  return { parser, ast, id, text };
}

describe('9b clearFormat: fully styled range strips to default + Normal', () => {
  it('removes run props, resets pStyle, drops spacing, keeps text', async () => {
    const { ast, id, text } = await styledFixture();
    const before = ast.sections[0].nodes.find((n: any) => n.id === id)!;
    const v0 = before.version as number;
    expect(runPropTags(before.xml as Element).length).toBeGreaterThan(0);
    expect(hasSpacing(before.xml as Element)).toBe(true);

    const res = executeClearFormat({ action: 'clearFormat', target: id } as any, ast);
    expect(res.cleared).toEqual([id]);

    const after = ast.sections[0].nodes.find((n: any) => n.id === id)!;
    const remaining = runPropTags(after.xml as Element);
    for (const tag of STRIPPED) expect(remaining).not.toContain(tag);
    expect(pStyleVal(after.xml as Element)).toBe('Normal');
    expect(hasSpacing(after.xml as Element)).toBe(false);
    expect(after.metadata.namedStyle).toBe('Normal');
    expect(after.metadata.text).toBe(text);
    expect(after.version).toBe(v0 + 1);
    // Paragraph-mark defaults stripped too: no leftover pPr/rPr formatting.
    const pPr = (after.xml as Element).getElementsByTagName('w:pPr')[0];
    const markRPr = pPr?.getElementsByTagName('w:rPr')[0];
    if (markRPr) {
      for (let i = 0; i < markRPr.childNodes.length; i++) {
        const c = markRPr.childNodes[i] as Element;
        if ((c as any).nodeType === 1) expect(STRIPPED).not.toContain((c as Element).tagName);
      }
    }
  });
});

describe('9b clearFormat: partial range leaves siblings styled', () => {
  it('clears one target; the other keeps bold/color/highlight/spacing', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const [a, b] = [ast.sections[0].nodes[0].id, ast.sections[0].nodes[1].id];
    for (const id of [a, b]) {
      executeStyle(
        { action: 'style', target: id, properties: { bold: true, color: 'FF0000', highlight: 'yellow', spacingBefore: 240 } } as any,
        ast
      );
    }
    executeClearFormat({ action: 'clearFormat', target: b } as any, ast);

    const kept = ast.sections[0].nodes.find((n: any) => n.id === a)!;
    const keptTags = runPropTags(kept.xml as Element);
    expect(keptTags).toContain('w:b');
    expect(keptTags).toContain('w:color');
    expect(keptTags).toContain('w:highlight');
    expect(hasSpacing(kept.xml as Element)).toBe(true);

    const cleared = ast.sections[0].nodes.find((n: any) => n.id === b)!;
    const clearedTags = runPropTags(cleared.xml as Element);
    for (const tag of STRIPPED) expect(clearedTags).not.toContain(tag);
    expect(pStyleVal(cleared.xml as Element)).toBe('Normal');
    expect(hasSpacing(cleared.xml as Element)).toBe(false);
  });
});

describe('9b clearFormat: empty / error cases', () => {
  it('plain paragraph clears as a no-op with text intact', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const node = ast.sections[0].nodes[0];
    const text = node.metadata.text as string;
    const res = executeClearFormat({ action: 'clearFormat', target: node.id } as any, ast);
    expect(res.cleared).toEqual([node.id]);
    expect(node.metadata.text).toBe(text);
    expect(pStyleVal(node.xml as Element)).toBe('Normal');
  });

  it('missing node and empty target throw loudly', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    expect(() => executeClearFormat({ action: 'clearFormat', target: 'ghost_9' } as any, ast))
      .toThrow('Node ghost_9 not found');
    expect(() => executeClearFormat({ action: 'clearFormat', target: '' } as any, ast))
      .toThrow('ClearFormat operation requires a target');
    expect(() => executeClearFormat({ action: 'clearFormat', target: [] } as any, ast))
      .toThrow('ClearFormat operation requires a target');
  });
});

describe('9b clearFormat: engine registration + round-trip rezip clean', () => {
  it('routes through OperationEngine (single + batch) and lands in history', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const engine = new OperationEngine(ast, parser);
    const id = ast.sections[0].nodes[0].id as string;
    executeStyle({ action: 'style', target: id, properties: { bold: true, italic: true } } as any, ast);

    await engine.execute({ action: 'clearFormat', target: id } as any);
    expect(engine.getHistory()).toHaveLength(1);
    expect(runPropTags(ast.sections[0].nodes[0].xml as Element)).not.toContain('w:b');

    const res = await engine.executeBatch({
      operations: [{ action: 'clearFormat', target: id } as any],
      sequential: true,
    });
    expect(res.success).toBe(true);
    expect(res.rollbackTriggered).toBe(false);
    expect(res.appliedCount).toBe(1);
  });

  it('rezips clean: valid zip, reload keeps Normal + text, no stripped tags on disk', async () => {
    const { parser, ast, id, text } = await styledFixture();
    executeClearFormat({ action: 'clearFormat', target: id } as any, ast);

    const out = parser.rezip(ast);
    expect(parser.validateZipIntegrity().isValid).toBe(true);
    const docXml = new AdmZip(out).getEntry('word/document.xml')!.getData().toString('utf8');
    // Tag-boundary match ([\s/>]) so <w:body>/<w:bdr> can't false-positive.
    for (const tag of ['w:b', 'w:i', 'w:color', 'w:highlight', 'w:strike', 'w:dstrike', 'w:rFonts', 'w:vertAlign', 'w:sz', 'w:u']) {
      expect(docXml).not.toMatch(new RegExp(`<${tag}[\\s/>]`));
    }
    expect(docXml).toContain('w:val="Normal"');
    expect(docXml).not.toMatch(/<w:spacing[^>]*w:(before|after|line)=/);

    const probe = new DocxParser(out);
    const ast2: any = await probe.getAST();
    const reloaded = ast2.sections[0].nodes.find((n: any) => n.metadata.text === text)!;
    expect(reloaded).toBeDefined();
    expect(reloaded.metadata.namedStyle).toBe('Normal');
    expect(runPropTags(reloaded.xml as Element).filter(t => STRIPPED.includes(t))).toHaveLength(0);
    expect(probe.validateZipIntegrity().isValid).toBe(true);
  });
});

/**
 * 9b follow-up — full RPR_CHILD_ORDER probe.
 *
 * Combined Ctrl+Space + Ctrl+Q semantics: Ctrl+Space must clear EVERY
 * run-level direct-format child of w:rPr (the canonical RPR_CHILD_ORDER
 * list, not a hand-pinned subset — the old 13-tag set let caps/smallCaps/
 * shd/bdr/kern/position/vanish/emboss/shadow/outline/rStyle/noProof/
 * webHidden/etc. survive); Ctrl+Q must clear paragraph peers
 * jc/ind/pBdr/spacing. Inject-all → clear → zero survive.
 */
describe('9b clearFormat: full RPR_CHILD_ORDER probe — zero run tags survive', () => {
  function injectAllRprTags(rPr: Element): void {
    const doc = rPr.ownerDocument;
    while (rPr.firstChild) rPr.removeChild(rPr.firstChild);
    for (const tag of RPR_CHILD_ORDER) {
      const el = doc.createElement(tag);
      if (tag === 'w:color') el.setAttribute('w:val', 'FF0000');
      if (tag === 'w:sz' || tag === 'w:szCs') el.setAttribute('w:val', '24');
      if (tag === 'w:rStyle') el.setAttribute('w:val', 'Strong');
      if (tag === 'w:jc') el.setAttribute('w:val', 'center');
      rPr.appendChild(el);
    }
  }

  it('injects every canonical RPR tag into run rPr + mark rPr → all stripped', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const node = ast.sections[0].nodes[0];
    const id = node.id as string;
    const text = node.metadata.text as string;
    const pEl = node.xml as Element;
    const doc = pEl.ownerDocument;

    // Run rPr: inject the full canonical order.
    const run = pEl.getElementsByTagName('w:r')[0];
    let rPr = run.getElementsByTagName('w:rPr')[0];
    if (!rPr || (rPr as Element).parentNode !== run) {
      rPr = doc.createElement('w:rPr');
      run.insertBefore(rPr, run.firstChild);
    }
    injectAllRprTags(rPr as Element);

    // Paragraph-mark rPr: same full injection (pilcrow must clear too).
    let pPr = pEl.getElementsByTagName('w:pPr')[0];
    if (!pPr) {
      pPr = doc.createElement('w:pPr');
      pEl.insertBefore(pPr, pEl.firstChild);
    }
    let markRPr = (pPr as Element).getElementsByTagName('w:rPr')[0];
    if (!markRPr || (markRPr as Element).parentNode !== pPr) {
      markRPr = doc.createElement('w:rPr');
      (pPr as Element).appendChild(markRPr);
    }
    injectAllRprTags(markRPr as Element);

    // Pre-condition: the hole the reviewer probed — old 13-tag set left 24+
    // survivors (caps/smallCaps/shd/bdr/kern/position/spacing/w/vanish/...).
    const before = runPropTags(pEl);
    for (const survivor of ['w:caps', 'w:smallCaps', 'w:shd', 'w:bdr', 'w:kern', 'w:position', 'w:vanish', 'w:emboss', 'w:shadow', 'w:outline', 'w:rStyle', 'w:noProof', 'w:webHidden']) {
      expect(before).toContain(survivor);
    }
    expect(before.length).toBeGreaterThanOrEqual(RPR_CHILD_ORDER.length);

    executeClearFormat({ action: 'clearFormat', target: id } as any, ast);

    const after = ast.sections[0].nodes.find((n: any) => n.id === id)!;
    const remaining = runPropTags(after.xml as Element);
    expect(remaining).toHaveLength(0);
    // Mark rPr either removed or emptied — never a formatting carrier.
    const pPrAfter = (after.xml as Element).getElementsByTagName('w:pPr')[0];
    const markAfter = pPrAfter?.getElementsByTagName('w:rPr')[0];
    if (markAfter) expect(markAfter.childNodes.length).toBe(0);
    expect(after.metadata.text).toBe(text);
  });
});

describe('9b clearFormat: para reset — jc/ind/pBdr/spacing stripped, pStyle Normal', () => {
  function paraChildTags(pEl: Element): string[] {
    const out: string[] = [];
    const pPr = pEl.getElementsByTagName('w:pPr')[0];
    if (!pPr) return out;
    for (let i = 0; i < pPr.childNodes.length; i++) {
      const c = pPr.childNodes[i] as Element;
      if ((c as any).nodeType === 1) out.push((c as Element).tagName);
    }
    return out;
  }

  it('Ctrl+Q scope: jc + ind + pBdr + spacing all drop', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const node = ast.sections[0].nodes[0];
    const id = node.id as string;
    const pEl = node.xml as Element;
    const doc = pEl.ownerDocument;

    let pPr = pEl.getElementsByTagName('w:pPr')[0];
    if (!pPr) {
      pPr = doc.createElement('w:pPr');
      pEl.insertBefore(pPr, pEl.firstChild);
    }
    const ensure = (tag: string, attrs: Record<string, string> = {}) => {
      let el = (pPr as Element).getElementsByTagName(tag)[0];
      if (!el || (el as Element).parentNode !== pPr) {
        el = doc.createElement(tag) as unknown as Element;
        (pPr as Element).appendChild(el as unknown as Node);
      }
      for (const [k, v] of Object.entries(attrs)) (el as Element).setAttribute(k, v);
      // pBdr needs a visible child so the probe is realistic.
      if (tag === 'w:pBdr' && (el as Element).childNodes.length === 0) {
        const bottom = doc.createElement('w:bottom');
        (bottom as Element).setAttribute('w:val', 'single');
        (el as Element).appendChild(bottom);
      }
    };
    ensure('w:jc', { 'w:val': 'center' });
    ensure('w:ind', { 'w:left': '720', 'w:right': '720' });
    ensure('w:pBdr', {});
    ensure('w:spacing', { 'w:before': '240', 'w:after': '120', 'w:line': '480' });

    const beforeTags = paraChildTags(pEl);
    for (const t of ['w:jc', 'w:ind', 'w:pBdr', 'w:spacing']) expect(beforeTags).toContain(t);

    executeClearFormat({ action: 'clearFormat', target: id } as any, ast);

    const after = ast.sections[0].nodes.find((n: any) => n.id === id)!;
    const afterTags = paraChildTags(after.xml as Element);
    for (const t of ['w:jc', 'w:ind', 'w:pBdr', 'w:spacing']) expect(afterTags).not.toContain(t);
    const pPrAfter = (after.xml as Element).getElementsByTagName('w:pPr')[0];
    expect(pPrAfter.getElementsByTagName('w:pStyle')[0]?.getAttribute('w:val')).toBe('Normal');
    expect(after.metadata.namedStyle).toBe('Normal');
  });
});

describe('heading style apply keeps heading glued to following text', () => {
  it('applying a heading style adds keepNext after pStyle in schema order', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const id = ast.sections[0].nodes[0].id as string;
    executeNamedStyle(
      { action: 'namedStyle', operation: 'apply', styleName: 'Heading1', target: id } as any,
      ast, parser
    );
    const p = ast.sections[0].nodes[0].xml as Element;
    const pPr = p.getElementsByTagName('w:pPr')[0];
    expect(pPr).toBeDefined();
    expect(pPr.getElementsByTagName('w:keepNext').length).toBe(1);
    const tags = Array.from(pPr.childNodes)
      .filter((n: any) => n.nodeType === 1)
      .map((n: any) => (n as Element).tagName);
    expect(tags.indexOf('w:pStyle')).toBeLessThan(tags.indexOf('w:keepNext'));
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
  });

  it('re-applying a heading never duplicates keepNext and plain styles add none', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast: any = await parser.getAST();
    const id = ast.sections[0].nodes[0].id as string;
    const op = { action: 'namedStyle', operation: 'apply', styleName: 'Heading2', target: id } as any;
    executeNamedStyle(op, ast, parser);
    executeNamedStyle(op, ast, parser);
    const p = ast.sections[0].nodes[0].xml as Element;
    expect(p.getElementsByTagName('w:keepNext').length).toBe(1);
    executeNamedStyle(
      { action: 'namedStyle', operation: 'apply', styleName: 'Normal', target: id } as any,
      ast, parser
    );
    const pPr = p.getElementsByTagName('w:pPr')[0];
    expect(pPr.getElementsByTagName('w:pStyle')[0]?.getAttribute('w:val')).toBe('Normal');
    const res = await validateDocxBuffer(parser.rezip());
    expect(res.ok).toBe(true);
  });
});
