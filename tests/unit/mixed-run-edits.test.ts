import { describe, expect, it } from 'vitest';
import AdmZip from 'adm-zip';
import { DocxParser, OperationEngine, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { listComments } from '../../packages/engine/src/ooxml/comments';

async function setup() {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  const body = '<w:p><w:r><w:rPr><w:b/><w:color w:val="0000FF"/></w:rPr><w:t xml:space="preserve">Bold </w:t></w:r><w:r><w:t xml:space="preserve">plain middle </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r></w:p>';
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  return {parser, ast, engine:new OperationEngine(ast, parser), id:ast.sections[0].nodes[0].id};
}

async function fromParagraph(body: string) {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
  const parser = new DocxParser(zip.toBuffer());
  const ast = await parser.getAST();
  return {parser, ast, engine:new OperationEngine(ast, parser), id:ast.sections[0].nodes[0].id};
}

function characterStyles(paragraph: Element): string[] {
  const out: string[] = [];
  for (const run of Array.from(paragraph.getElementsByTagName('w:r'))) {
    let deleted = false;
    for (let node: Node | null = run.parentNode; node && node !== paragraph; node = node.parentNode) {
      if ((node as Element).tagName === 'w:del') deleted = true;
    }
    if (deleted) continue;
    const style = `${run.getElementsByTagName('w:b').length ? 'b' : ''}${run.getElementsByTagName('w:i').length ? 'i' : ''}`;
    for (const t of Array.from(run.getElementsByTagName('w:t'))) for (const char of t.textContent ?? '') out.push(`${char}:${style}`);
  }
  return out;
}

describe('mixed formatting edits', () => {
  it('preserves the text and per-character formatting invariants across seeded edits', async () => {
    let seed = 123456;
    const random = () => { seed = (seed * 1664525 + 1013904223) >>> 0; return seed; };
    for (let trial = 0; trial < 10; trial++) {
      const words = ['alpha','bravo','charlie','delta','echo','foxtrot'];
      const pick = random() % words.length;
      const target = words.map((word, i) => i === pick ? `new${trial}` : word).join(' ');
      const p = `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${words.slice(0,2).join(' ')} </w:t></w:r><w:r><w:t xml:space="preserve">${words.slice(2,4).join(' ')} </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>${words.slice(4).join(' ')}</w:t></w:r></w:p>`;
      const tracked = await fromParagraph(p);
      const original = characterStyles(tracked.parser.getDocumentDom().getElementsByTagName('w:p')[0]);
      const edit = await tracked.engine.executeBatch({operations:[{action:'replace',target:tracked.id,text:target,trackChanges:true}]} as never);
      expect(edit.success, JSON.stringify(edit.errors)).toBe(true);
      expect((await validateDocxBuffer(tracked.parser.rezip(tracked.ast))).ok).toBe(true);
      const plain = await fromParagraph(p);
      expect((await plain.engine.executeBatch({operations:[{action:'replace',target:plain.id,text:target,trackChanges:false}]} as never)).success).toBe(true);
      const expected = characterStyles(plain.parser.getDocumentDom().getElementsByTagName('w:p')[0]);
      const accepted = await tracked.engine.executeBatch({operations:[{action:'trackedChange',operation:'accept',ids:edit.revisionIds}]} as never);
      expect(accepted.success, JSON.stringify(accepted.errors)).toBe(true);
      expect(characterStyles(tracked.parser.getDocumentDom().getElementsByTagName('w:p')[0])).toEqual(expected);
      const rejectedDoc = await fromParagraph(p);
      const second = await rejectedDoc.engine.executeBatch({operations:[{action:'replace',target:rejectedDoc.id,text:target,trackChanges:true}]} as never);
      const rejected = await rejectedDoc.engine.executeBatch({operations:[{action:'trackedChange',operation:'reject',ids:second.revisionIds}]} as never);
      expect(rejected.success, JSON.stringify(rejected.errors)).toBe(true);
      expect(characterStyles(rejectedDoc.parser.getDocumentDom().getElementsByTagName('w:p')[0])).toEqual(original);
    }
  }, 30000);
  it('keeps the formatting of unchanged text on replace', async () => {
    const {parser, engine, id} = await setup();
    const result = await engine.executeBatch({operations:[{action:'replace', target:id, text:'Bold plain changed italic', preserveFormatting:true}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const runs = Array.from(parser.getDocumentDom().getElementsByTagName('w:r'));
    expect(runs.some(r => r.textContent === 'Bold ' && r.getElementsByTagName('w:b').length === 1)).toBe(true);
    expect(runs.some(r => r.textContent === 'italic' && r.getElementsByTagName('w:i').length === 1)).toBe(true);
    expect(runs.some(r => r.textContent?.includes('changed') && r.getElementsByTagName('w:b').length === 0)).toBe(true);
  });

  it('edits inside a hyperlink while preserving its relationship and surrounding anchors', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const body = '<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r><w:bookmarkStart w:id="4" w:name="here"/><w:hyperlink r:id="rIdLink"><w:r><w:rPr><w:b/></w:rPr><w:t>linked</w:t></w:r></w:hyperlink><w:bookmarkEnd w:id="4"/><w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>';
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
    const rels = zip.readAsText('word/_rels/document.xml.rels').replace('</Relationships>', '<Relationship Id="rIdLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com" TargetMode="External"/></Relationships>');
    zip.updateFile('word/_rels/document.xml.rels', Buffer.from(rels));
    const parser = new DocxParser(zip.toBuffer()); const ast = await parser.getAST();
    const id = ast.sections[0].nodes[0].id;
    const result = await new OperationEngine(ast, parser).executeBatch({operations:[{action:'replace',target:id,text:'Before revised after',trackChanges:false}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    const hyperlink = parser.getDocumentDom().getElementsByTagName('w:hyperlink')[0];
    expect(hyperlink.getAttribute('r:id')).toBe('rIdLink');
    expect(hyperlink.textContent).toBe('revised');
    expect(hyperlink.getElementsByTagName('w:b').length).toBeGreaterThan(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:bookmarkStart')[0].getAttribute('w:name')).toBe('here');
    const saved = parser.rezip(ast);
    expect(new AdmZip(saved).readAsText('word/_rels/document.xml.rels')).toContain('Id="rIdLink"');
    expect((await validateDocxBuffer(saved)).ok).toBe(true);
  });

  it('tracks only the requested text and restores formatting on rejection', async () => {
    const {parser, engine, id} = await setup();
    const result = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id,text:'middle'}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(Array.from(parser.getDocumentDom().getElementsByTagName('w:delText')).map(n => n.textContent).join('')).toBe('middle');
    await engine.executeBatch({operations:[{action:'trackedChange',operation:'rejectAll'}]} as never);
    const p = parser.getDocumentDom().getElementsByTagName('w:p')[0];
    expect(p.textContent).toBe('Bold plain middle italic');
    expect(p.getElementsByTagName('w:b').length).toBe(1);
    expect(p.getElementsByTagName('w:i').length).toBe(1);
  });
  it('edits outside a complex field while preserving the complete field', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const body = `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${field}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`;
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
    const parser = new DocxParser(zip.toBuffer()); const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser); const id = ast.sections[0].nodes[0].id;
    const result = await engine.executeBatch({operations:[{action:'replace',target:id,text:'Earlier 1 after'}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.getDocumentDom().getElementsByTagName('w:fldChar').length).toBe(3);
    expect(parser.getDocumentDom().getElementsByTagName('w:instrText')[0].textContent).toBe(' PAGE ');
    expect(ast.sections[0].nodes[0].metadata.text).toBe('Earlier 1 after');
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('tracks a complete field deletion and restores it on rejection', async () => {
    const zip = new AdmZip(createBlankDocxBuffer('blank'));
    const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const body = `<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${field}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`;
    zip.updateFile('word/document.xml', Buffer.from(zip.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr)/, `<w:body>${body}$1`)));
    const parser = new DocxParser(zip.toBuffer()); const ast = await parser.getAST();
    const engine = new OperationEngine(ast, parser); const id = ast.sections[0].nodes[0].id;
    const result = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id,text:'1'}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.getDocumentDom().getElementsByTagName('w:delInstrText').length).toBe(1);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
    await engine.executeBatch({operations:[{action:'trackedChange',operation:'rejectAll'}]} as never);
    expect(parser.getDocumentDom().getElementsByTagName('w:instrText')[0].textContent).toBe(' PAGE ');
    expect(parser.getDocumentDom().getElementsByTagName('w:fldChar').length).toBe(3);
  });

  it('replaces an entire field as a tracked change without nesting the insertion inside a deletion', async () => {
    const field = '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>1</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>';
    const {parser, ast, engine, id} = await fromParagraph(`<w:p><w:r><w:t xml:space="preserve">Before </w:t></w:r>${field}<w:r><w:t xml:space="preserve"> after</w:t></w:r></w:p>`);
    const out = await engine.executeBatch({operations:[{action:'replace',target:id,text:'Before 2 after',trackChanges:true}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
    const undone = await engine.executeBatch({operations:[{action:'trackedChange',operation:'reject',ids:out.revisionIds}]} as never);
    expect(undone.success, JSON.stringify(undone.errors)).toBe(true);
    expect(parser.getDocumentDom().getElementsByTagName('w:instrText').length).toBe(1);
  });

  it('refuses tracked simple-field deletion without changing the document or consuming an ID', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:r><w:t xml:space="preserve">Page </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>');
    const before = parser.getDocumentDom().toString();
    const nextId = ast.nextChangeId;
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id,text:'1'}]} as never);
    expect(out.success).toBe(false);
    expect(out.errors.some(error => error.error.includes('simple field'))).toBe(true);
    expect(parser.getDocumentDom().toString()).toBe(before);
    expect(ast.nextChangeId).toBe(nextId);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('returns all new IDs in document order for a mixed-format replacement', async () => {
    const {parser, engine, id} = await setup();
    const result = await engine.executeBatch({operations:[{action:'replace',target:id,text:'Fresh plain changed italic',trackChanges:true}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(result.revisionIds?.length).toBeGreaterThan(1);
    expect(result.revisionIds).toEqual([...result.revisionIds!].sort((a,b) => a-b));
    const xml = parser.getDocumentDom().getElementsByTagName('w:p')[0].toString();
    expect(xml).toContain('<w:b/>');
    expect(xml).toContain('<w:i/>');
    expect((await validateDocxBuffer(parser.rezip())).ok).toBe(true);
  });

  it('leaves an earlier insertion intact when editing another span', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:ins w:id="41" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t xml:space="preserve">Prior </w:t></w:r></w:ins><w:r><w:t>ordinary</w:t></w:r></w:p>');
    const result = await engine.executeBatch({operations:[{action:'replace',target:id,text:'Prior changed',trackChanges:true}]} as never);
    expect(result.success, JSON.stringify(result.errors)).toBe(true);
    expect(parser.getDocumentDom().getElementsByTagName('w:ins')[0].getAttribute('w:id')).toBe('41');
    expect(result.revisionIds?.every(id => id > 41)).toBe(true);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('refuses an ambiguous tracked text delete without spending an ID', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:r><w:t>same same</w:t></w:r></w:p>');
    const before = parser.getDocumentDom().getElementsByTagName('w:p')[0].toString();
    const nextId = ast.nextChangeId;
    const result = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id,text:'same'}]} as never);
    expect(result.success).toBe(false);
    expect(result.errors.some(error => error.error.includes('occurrence'))).toBe(true);
    expect(parser.getDocumentDom().getElementsByTagName('w:p')[0].toString()).toBe(before);
    expect(ast.nextChangeId).toBe(nextId);
  });

  it('refuses to split a Unicode character and leaves the document unchanged', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:r><w:t>before 😀 after</w:t></w:r></w:p>');
    const before = parser.getDocumentDom().toString();
    const nextId = ast.nextChangeId;
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id,text:'\ud83d'}]} as never);
    expect(out.success).toBe(false);
    expect(parser.getDocumentDom().toString()).toBe(before);
    expect(ast.nextChangeId).toBe(nextId);
  });

  it('refuses a whole-paragraph delete when its mark already has a revision', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:pPr><w:rPr><w:ins w:id="3" w:author="A"/></w:rPr></w:pPr><w:r><w:t>text</w:t></w:r></w:p>');
    const before = parser.getDocumentDom().toString();
    const nextId = ast.nextChangeId;
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id}]} as never);
    expect(out.success).toBe(false);
    expect(parser.getDocumentDom().toString()).toBe(before);
    expect(ast.nextChangeId).toBe(nextId);
  });

  it('writes a tracked paragraph mark in schema order when section properties exist', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:pPr><w:sectPr/></w:pPr><w:r><w:t>text</w:t></w:r></w:p>');
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it.each([500, 2000, 5000])('keeps the unchanged runs in a %i-word paragraph', async count => {
    const words = Array.from({length:count}, (_, i) => `word${i}`);
    const before = words.slice(0, count / 2).join(' ') + ' ';
    const after = words.slice(count / 2).join(' ');
    const {parser, engine, id} = await fromParagraph(`<w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">${before}</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>${after}</w:t></w:r></w:p>`);
    const edited = `${before}${after.replace(`word${count / 2}`, 'changed')}`;
    const start = performance.now();
    const out = await engine.executeBatch({operations:[{action:'replace',target:id,text:edited,trackChanges:true}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    const p = parser.getDocumentDom().getElementsByTagName('w:p')[0];
    expect(p.getElementsByTagName('w:b').length).toBeGreaterThan(0);
    expect(p.getElementsByTagName('w:i').length).toBeGreaterThan(0);
    expect(p.getElementsByTagName('w:delText')[0].textContent).toBe(`word${count / 2}`);
    process.stdout.write(`mixed-edit ${count} words: ${Math.round(performance.now() - start)} ms\n`);
  });

  it('rejects selected revisions while keeping a previous insertion', async () => {
    const {parser, engine, id} = await fromParagraph('<w:p><w:ins w:id="41" w:author="Prior" w:date="2026-01-01T00:00:00Z"><w:r><w:t>original</w:t></w:r></w:ins></w:p>');
    const edit = await engine.executeBatch({operations:[{action:'replace',target:id,text:'changed',trackChanges:true}]} as never);
    expect(edit.success, JSON.stringify(edit.errors)).toBe(true);
    const undo = await engine.executeBatch({operations:[{action:'trackedChange',operation:'reject',ids:edit.revisionIds}]} as never);
    expect(undo.success, JSON.stringify(undo.errors)).toBe(true);
    const p = parser.getDocumentDom().getElementsByTagName('w:p')[0];
    expect(p.textContent).toBe('original');
    expect(p.getElementsByTagName('w:ins')[0].getAttribute('w:id')).toBe('41');
  });

  it('refuses a duplicated revision ID without changing the document', async () => {
    const {parser, engine} = await fromParagraph('<w:p><w:ins w:id="8" w:author="A"><w:r><w:t>one</w:t></w:r></w:ins><w:ins w:id="8" w:author="A"><w:r><w:t>two</w:t></w:r></w:ins></w:p>');
    const before = parser.getDocumentDom().toString();
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'accept',ids:[8]}]} as never);
    expect(out.success).toBe(false);
    expect(out.errors.some(e => e.error.includes('ambiguous'))).toBe(true);
    expect(parser.getDocumentDom().toString()).toBe(before);
  });

  it('preserves a tab inside a mixed run when tracking a deletion', async () => {
    const {parser, ast, engine, id} = await fromParagraph('<w:p><w:r><w:t>one</w:t><w:tab/><w:t>two</w:t></w:r></w:p>');
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'delete',target:id,text:'onetwo'}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(out.revisionIds).toHaveLength(2);
    expect(parser.getDocumentDom().getElementsByTagName('w:tab').length).toBe(1);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('does not rewrite an untouched mixed text and tab run', async () => {
    const {parser, engine, id} = await fromParagraph('<w:p><w:r><w:t>first</w:t><w:tab/><w:t xml:space="preserve"> segment </w:t></w:r><w:r><w:t>target</w:t></w:r></w:p>');
    const before = parser.getDocumentDom().getElementsByTagName('w:r')[0].toString();
    const out = await engine.executeBatch({operations:[{action:'replace',target:id,text:'first segment revised',trackChanges:false}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(parser.getDocumentDom().getElementsByTagName('w:r')[0].toString()).toBe(before);
  });

  it('accepts only revisions by the selected author', async () => {
    const {parser, engine} = await fromParagraph('<w:p><w:ins w:id="8" w:author="A"><w:r><w:t>one</w:t></w:r></w:ins><w:ins w:id="9" w:author="B"><w:r><w:t>two</w:t></w:r></w:ins></w:p>');
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'accept',author:'A'}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    const remaining = parser.getDocumentDom().getElementsByTagName('w:ins');
    expect(remaining.length).toBe(1);
    expect(remaining[0].getAttribute('w:author')).toBe('B');
  });

  it('refuses an ID selection atomically when one revision is unsupported', async () => {
    const {parser, engine} = await fromParagraph('<w:p><w:ins w:id="7" w:author="A"><w:r><w:t>one</w:t></w:r></w:ins><w:moveFrom w:id="8" w:author="A"><w:r><w:t>two</w:t></w:r></w:moveFrom></w:p>');
    const before = parser.getDocumentDom().toString();
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'accept',ids:[7,8]}]} as never);
    expect(out.success).toBe(false);
    expect(parser.getDocumentDom().toString()).toBe(before);
  });

  it('reports unresolved revisions when selecting by author', async () => {
    const {parser, engine} = await fromParagraph('<w:p><w:ins w:id="7" w:author="A"><w:r><w:t>one</w:t></w:r></w:ins><w:moveFrom w:id="8" w:author="A"><w:r><w:t>two</w:t></w:r></w:moveFrom></w:p>');
    const out = await engine.executeBatch({operations:[{action:'trackedChange',operation:'accept',author:'A'}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(out.revisionResolution?.unresolved).toEqual(expect.arrayContaining([expect.objectContaining({kind:'moveFrom'})]));
    expect(parser.getDocumentDom().getElementsByTagName('w:ins').length).toBe(0);
    expect(out.revisionsPending).toBe(1);
  });

  it('reads and resolves a comment through the operation interface', async () => {
    const {parser, ast, engine, id} = await setup();
    const created = await engine.executeBatch({operations:[{action:'comment',operation:'insert',target:id,text:'Review this',author:'Reviewer'}]} as never);
    expect(created.success, JSON.stringify(created.errors)).toBe(true);
    expect(listComments(parser.getZipAdapter())).toMatchObject([{id:0, text:'Review this', author:'Reviewer', resolved:false}]);
    const resolved = await engine.executeBatch({operations:[{action:'comment',operation:'resolve',target:'0'}]} as never);
    expect(resolved.success, JSON.stringify(resolved.errors)).toBe(true);
    expect(listComments(parser.getZipAdapter())[0].resolved).toBe(true);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('deletes a comment and its reply as one thread', async () => {
    const {parser, ast, engine, id} = await setup();
    for (const op of [{action:'comment',operation:'insert',target:id,text:'Parent'}, {action:'comment',operation:'reply',target:'0',text:'Reply'}]) {
      const out = await engine.executeBatch({operations:[op]} as never);
      expect(out.success, JSON.stringify(out.errors)).toBe(true);
    }
    expect(listComments(parser.getZipAdapter())).toHaveLength(2);
    const out = await engine.executeBatch({operations:[{action:'comment',operation:'delete',target:'0'}]} as never);
    expect(out.success, JSON.stringify(out.errors)).toBe(true);
    expect(listComments(parser.getZipAdapter())).toHaveLength(0);
    expect(parser.getDocumentDom().getElementsByTagName('w:commentReference').length).toBe(0);
    expect((await validateDocxBuffer(parser.rezip(ast))).ok).toBe(true);
  });

  it('refuses comment deletion when an anchor is missing', async () => {
    const {parser, engine, id} = await setup();
    await engine.executeBatch({operations:[{action:'comment',operation:'insert',target:id,text:'Parent'}]} as never);
    const end = parser.getDocumentDom().getElementsByTagName('w:commentRangeEnd')[0];
    end.parentNode!.removeChild(end);
    const before = parser.getDocumentDom().toString();
    const comments = parser.getZipAdapter().readText('word/comments.xml');
    const out = await engine.executeBatch({operations:[{action:'comment',operation:'delete',target:'0'}]} as never);
    expect(out.success).toBe(false);
    expect(parser.getDocumentDom().toString()).toBe(before);
    expect(parser.getZipAdapter().readText('word/comments.xml')).toBe(comments);
  });

});
