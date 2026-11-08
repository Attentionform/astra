import { describe, expect, it } from 'vitest';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import AdmZip from 'adm-zip';
import { DocxParser } from '@attentionform/astra-engine';

describe('saving a Word 2016 file keeps the content types of its comment parts', () => {
  it('commentsIds and commentsExtensible keep their own content types', async () => {
    const parser = new DocxParser(readFileSync(join(__dirname, '..', 'fixtures', 'word-revisions', 'delete-paragraph-comment-thread.docx')));
    const ast = await parser.getAST();
    const types = new AdmZip(parser.rezip(ast)).readAsText('[Content_Types].xml');
    expect(types).toContain('PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"');
    expect(types).toContain('PartName="/word/commentsExtensible.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"');
  });
});
