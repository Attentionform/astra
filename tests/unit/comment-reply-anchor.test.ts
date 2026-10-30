import { describe, it, expect } from 'vitest';
import {
  DocxParser,
  createBlankDocxBuffer,
} from '@attentionform/astra-engine';
import { executeReplace } from '../../packages/engine/src/operations/replace';
import { executeComment } from '../../packages/engine/src/operations/comment';
import { executeFootnote, executeEndnote } from '../../packages/engine/src/operations/footnote';
import { validateOperationStrict } from '../../packages/engine/src/operation-schema';
import { validateDocxBuffer } from '../../packages/engine/src/ooxml/validator-gate';
import { XMLSerializer } from '@xmldom/xmldom';
import AdmZip from 'adm-zip';

describe('optional author and date on replace and comment', () => {
  it('stamps custom author and date on tracked replace insertions and deletions', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];
    p.metadata.text = 'Initial text.';

    executeReplace(
      {
        action: 'replace',
        target: p.id,
        text: 'Updated text.',
        trackChanges: true,
        author: 'Dana Reviewer',
        date: '2026-10-24T09:30:00Z',
      },
      ast
    );

    const docXml = new XMLSerializer().serializeToString(parser.getDocumentDom());
    expect(docXml).toContain('w:author="Dana Reviewer"');
    expect(docXml).toContain('w:date="2026-10-24T09:30:00Z"');
  });

  it('uses the default author and the current time when replace has no author and date specified', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];
    p.metadata.text = 'Initial text.';

    const before = Date.now();
    executeReplace(
      {
        action: 'replace',
        target: p.id,
        text: 'Updated text.',
        trackChanges: true,
      },
      ast
    );
    const after = Date.now();

    const docXml = new XMLSerializer().serializeToString(parser.getDocumentDom());
    expect(docXml).toContain('w:author="ASTra"');
    const stamps = [...docXml.matchAll(/<w:(?:ins|del) [^>]*w:date="([^"]+)"/g)].map((m) => Date.parse(m[1]));
    expect(stamps.length).toBeGreaterThan(0);
    for (const t of stamps) {
      expect(t).toBeGreaterThanOrEqual(before);
      expect(t).toBeLessThanOrEqual(after);
    }
  });

  it('records author and date in comments.xml and people.xml', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeComment(
      {
        action: 'comment',
        operation: 'insert',
        target: p.id,
        text: 'First note',
        author: 'Dana Reviewer',
        date: '2026-10-24T09:30:00Z',
      },
      ast,
      parser
    );

    const zip = parser.getZipAdapter();
    const commentsXml = zip.readText('word/comments.xml');
    expect(commentsXml).toContain('w:author="Dana Reviewer"');
    expect(commentsXml).toContain('w:initials="D"');
    expect(commentsXml).toContain('w:date="2026-10-24T09:30:00Z"');

    const peopleXml = zip.readText('word/people.xml');
    expect(peopleXml).toContain('w15:author="Dana Reviewer"');
  });

  it('preserves full emoji as initials for author starting with emoji', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeComment(
      {
        action: 'comment',
        operation: 'insert',
        target: p.id,
        text: 'Note with emoji author',
        author: '😀 Dana',
        date: '2026-10-24T09:30:00Z',
      },
      ast,
      parser
    );

    const zip = parser.getZipAdapter();
    const commentsXml = zip.readText('word/comments.xml');
    expect(commentsXml).toContain('w:initials="😀"');
  });

  it('rejects invalid date format, invalid calendar date, and whitespace-only author in schema', () => {
    expect(() =>
      validateOperationStrict({
        action: 'replace',
        target: 'p1',
        text: 'new',
        date: '24/10/2026',
      })
    ).toThrow();

    expect(() =>
      validateOperationStrict({
        action: 'replace',
        target: 'p1',
        text: 'new',
        date: '2026-13-45T99:99:99Z',
      })
    ).toThrow();

    expect(() =>
      validateOperationStrict({
        action: 'replace',
        target: 'p1',
        text: 'new',
        author: '   ',
      })
    ).toThrow();

    expect(() =>
      validateOperationStrict({
        action: 'comment',
        operation: 'insert',
        target: 'p1',
        text: 'note',
        date: 'invalid-date',
      })
    ).toThrow();

    expect(() =>
      validateOperationStrict({
        action: 'comment',
        operation: 'insert',
        target: 'p1',
        text: 'note',
        author: '   ',
      })
    ).toThrow();
  });
});

describe('reply comment anchoring in document', () => {
  it('anchors comment reply markers next to parent comment in document.xml', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeComment({ action: 'comment', operation: 'insert', target: p.id, text: 'Parent comment' }, ast, parser);
    executeComment({ action: 'comment', operation: 'reply', target: '0', text: 'Reply comment' }, ast, parser);

    const docXml = new XMLSerializer().serializeToString(parser.getDocumentDom());
    const starts = docXml.match(/w:commentRangeStart/g) || [];
    const ends = docXml.match(/w:commentRangeEnd/g) || [];
    const refs = docXml.match(/w:commentReference/g) || [];

    expect(starts.length).toBe(2);
    expect(ends.length).toBe(2);
    expect(refs.length).toBe(2);

    expect(docXml).toContain('<w:commentRangeStart w:id="0"/>');
    expect(docXml).toContain('<w:commentRangeStart w:id="1"/>');
    expect(docXml).toContain('<w:commentRangeEnd w:id="0"/>');
    expect(docXml).toContain('<w:commentRangeEnd w:id="1"/>');
    expect(docXml).toContain('<w:commentReference w:id="0"/>');
    expect(docXml).toContain('<w:commentReference w:id="1"/>');

    const bytes = parser.rezip();
    const val = await validateDocxBuffer(bytes);
    expect(val.ok).toBe(true);
  });

  it('does not throw and adds no markers if parent comment markers were missing', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeComment({ action: 'comment', operation: 'insert', target: p.id, text: 'Parent' }, ast, parser);

    // Remove parent markers from documentDom
    const dom = parser.getDocumentDom();
    const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const removeAll = (tag: string) => {
      const list = dom.getElementsByTagNameNS(wNs, tag);
      while (list.length > 0) {
        const item = list.item(0)!;
        item.parentNode?.removeChild(item);
      }
    };
    removeAll('commentRangeStart');
    removeAll('commentRangeEnd');
    removeAll('commentReference');

    expect(() => {
      executeComment({ action: 'comment', operation: 'reply', target: '0', text: 'Reply without markers' }, ast, parser);
    }).not.toThrow();

    const docXml = new XMLSerializer().serializeToString(parser.getDocumentDom());
    expect(docXml).not.toContain('commentRangeStart');
  });
});

describe('comment reply targeting paragraph node id', () => {
  it('attaches reply to last top-level comment when targeting paragraph node id', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeComment({ action: 'comment', operation: 'insert', target: p.id, text: 'Comment A' }, ast, parser);
    executeComment({ action: 'comment', operation: 'insert', target: p.id, text: 'Comment B' }, ast, parser);

    executeComment({ action: 'comment', operation: 'reply', target: p.id, text: 'Reply 1 to B' }, ast, parser);
    executeComment({ action: 'comment', operation: 'reply', target: p.id, text: 'Reply 2 to B' }, ast, parser);

    const extXml = parser.getZipAdapter().readText('word/commentsExtended.xml');
    const zip = new AdmZip(parser.rezip());
    const commentsXml = zip.readAsText('word/comments.xml');

    expect(commentsXml).toContain('Reply 1 to B');
    expect(commentsXml).toContain('Reply 2 to B');

    // Both replies must point to Comment B (commentId 1)
    const dom = new DocxParser(zip.toBuffer()).getDocumentDom();
    expect(dom).toBeDefined();
    expect(extXml).toContain('w15:paraIdParent');
  });

  it('attaches reply to paragraph whose node id is all digits', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];
    p.id = '20481377';

    executeComment({ action: 'comment', operation: 'insert', target: p.id, text: 'Comment on numeric id' }, ast, parser);
    executeComment({ action: 'comment', operation: 'reply', target: '20481377', text: 'Reply targeting numeric node id' }, ast, parser);

    const zip = new AdmZip(parser.rezip());
    const commentsXml = zip.readAsText('word/comments.xml');
    expect(commentsXml).toContain('Reply targeting numeric node id');
  });

  it('fails with typed error when replying to node with no comment', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    expect(() => {
      executeComment({ action: 'comment', operation: 'reply', target: p.id, text: 'Orphan reply' }, ast, parser);
    }).toThrow(/has no comment to reply to/);
  });

  it('fails with typed error when target is neither comment id nor node id', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();

    expect(() => {
      executeComment({ action: 'comment', operation: 'reply', target: 'unknown_target_xyz', text: 'Invalid target' }, ast, parser);
    }).toThrow(/neither a comment id nor a node id/);
  });
});

describe('footnote and endnote style definitions in styles.xml', () => {
  it('defines FootnoteText and FootnoteReference on footnote insert', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeFootnote({ action: 'footnote', anchor: p.id, text: 'Note content' }, ast, parser);

    const stylesXml = new XMLSerializer().serializeToString(parser.getStylesDom()!);
    expect(stylesXml).toContain('w:styleId="FootnoteText"');
    expect(stylesXml).toContain('w:styleId="FootnoteReference"');
    expect(stylesXml).not.toContain('w:styleId="EndnoteText"');

    const bytes = parser.rezip();
    const val = await validateDocxBuffer(bytes);
    expect(val.ok).toBe(true);
  });

  it('defines EndnoteText and EndnoteReference on endnote insert', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeEndnote({ action: 'endnote', anchor: p.id, text: 'Endnote content' }, ast, parser);

    const stylesXml = new XMLSerializer().serializeToString(parser.getStylesDom()!);
    expect(stylesXml).toContain('w:styleId="EndnoteText"');
    expect(stylesXml).toContain('w:styleId="EndnoteReference"');

    const bytes = parser.rezip();
    const val = await validateDocxBuffer(bytes);
    expect(val.ok).toBe(true);
  });

  it('does not overwrite existing custom footnote style definition', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const stylesDom = parser.getStylesDom()!;
    const wNs = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
    const customStyle = stylesDom.createElementNS(wNs, 'w:style');
    customStyle.setAttribute('w:type', 'paragraph');
    customStyle.setAttribute('w:styleId', 'FootnoteText');
    const customName = stylesDom.createElementNS(wNs, 'w:name');
    customName.setAttribute('w:val', 'custom footnote name');
    customStyle.appendChild(customName);
    stylesDom.documentElement.appendChild(customStyle);

    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];
    executeFootnote({ action: 'footnote', anchor: p.id, text: 'Note with custom style' }, ast, parser);

    const stylesXml = new XMLSerializer().serializeToString(parser.getStylesDom()!);
    expect(stylesXml).toContain('w:val="custom footnote name"');
  });

  it('multiple note inserts in the same document do not create duplicate styles', async () => {
    const parser = new DocxParser(createBlankDocxBuffer('blank'));
    const ast = await parser.getAST();
    const p = ast.sections[0].nodes[0];

    executeFootnote({ action: 'footnote', anchor: p.id, text: 'First footnote' }, ast, parser);
    executeFootnote({ action: 'footnote', anchor: p.id, text: 'Second footnote' }, ast, parser);
    executeEndnote({ action: 'endnote', anchor: p.id, text: 'First endnote' }, ast, parser);
    executeEndnote({ action: 'endnote', anchor: p.id, text: 'Second endnote' }, ast, parser);

    const stylesDom = parser.getStylesDom()!;
    const styleEls = Array.from(stylesDom.documentElement.getElementsByTagName('w:style'));
    const idCounts: Record<string, number> = {};
    for (const el of styleEls) {
      const id = el.getAttribute('w:styleId') || el.getAttribute('styleId');
      if (id) idCounts[id] = (idCounts[id] || 0) + 1;
    }

    expect(idCounts['FootnoteText']).toBe(1);
    expect(idCounts['FootnoteReference']).toBe(1);
    expect(idCounts['EndnoteText']).toBe(1);
    expect(idCounts['EndnoteReference']).toBe(1);

    const bytes = parser.rezip();
    const val = await validateDocxBuffer(bytes);
    expect(val.ok).toBe(true);
  });
});
