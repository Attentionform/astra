import { describe, it, expect } from 'vitest';
import AdmZip from 'adm-zip';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import {
  createBlankDocxBuffer,
  sanitizeDocxBuffer,
  sanitizeDocumentXml,
  sanitizePackageEntries,
  validateDocxBuffer,
  validateSanitizeOptionsStrict,
} from '@attentionform/astra-engine';
import { run } from '../../packages/cli/src/cli';

const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';
const W15_NS = 'http://schemas.microsoft.com/office/word/2012/wordml';

const SECRET_DEL_PARA = 'LEAK-TRACKED-PARA-9f3k';
const SECRET_DEL_TABLE = 'LEAK-TRACKED-TABLE-4m7d';
const SECRET_DEL_HDR = 'LEAK-TRACKED-HEADER-2h8s';
const SECRET_COMMENT = 'LEAK-COMMENT-BODY-7q2x';
const INS_PARA = 'kept-para-ins-6a1p';
const INS_TABLE = 'kept-table-ins-3t9c';
const INS_HDR = 'kept-header-ins-5h2e';

function delXml(id: number, secret: string): string {
  return (
    `<w:del w:id="${id}" w:author="Mallory" w:date="2026-01-02T00:00:00Z">` +
    `<w:r><w:delText>${secret}</w:delText></w:r></w:del>`
  );
}

function insXml(id: number, text: string): string {
  return (
    `<w:ins w:id="${id}" w:author="Mallory" w:date="2026-01-02T00:00:00Z">` +
    `<w:r><w:t>${text}</w:t></w:r></w:ins>`
  );
}

/**
 * Valid dirty fixture: tracked deletions (paragraph + table cell + header),
 * tracked insertions, and a wired comment (comments.xml + range markers).
 * commentsExtended.xml / people.xml ride along as the engine's comment adjuncts.
 */
function buildDirtyDocx(): Buffer {
  const zip = new AdmZip(createBlankDocxBuffer('blank'));
  let doc = zip.readAsText('word/document.xml');
  const table =
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr>` +
    `<w:tblGrid><w:gridCol w:w="4000"/><w:gridCol w:w="4000"/></w:tblGrid>` +
    `<w:tr><w:tc><w:p><w:r><w:t>cell-a </w:t></w:r>${delXml(20, SECRET_DEL_TABLE)}</w:p></w:tc>` +
    `<w:tc><w:p>${insXml(21, INS_TABLE)}</w:p></w:tc></w:tr></w:tbl>`;
  const commentPara =
    `<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>commented text</w:t></w:r>` +
    `<w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>`;
  doc = doc.replace(
    '<w:sectPr>',
    `<w:p><w:r><w:t>Visible intro. </w:t></w:r>${delXml(10, SECRET_DEL_PARA)}${insXml(11, INS_PARA)}</w:p>` +
      table +
      commentPara +
      '<w:sectPr>'
  );
  doc = doc.replace('<w:pgSz', '<w:headerReference w:type="default" r:id="rId5"/><w:pgSz');
  zip.deleteFile('word/document.xml');
  zip.addFile('word/document.xml', Buffer.from(doc, 'utf8'));

  zip.addFile(
    'word/header1.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>hdr </w:t></w:r>` +
        `${delXml(30, SECRET_DEL_HDR)}${insXml(31, INS_HDR)}</w:p></w:hdr>`,
      'utf8'
    )
  );
  zip.addFile(
    'word/comments.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:comments xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}">` +
        `<w:comment w:id="0" w:author="Mallory" w:initials="M" w:date="2026-01-02T00:00:00Z">` +
        `<w:p w14:paraId="1A2B3C4D"><w:r><w:annotationRef/></w:r>` +
        `<w:r><w:t>${SECRET_COMMENT}</w:t></w:r></w:p></w:comment></w:comments>`,
      'utf8'
    )
  );
  zip.addFile(
    'word/commentsExtended.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w15:commentsEx xmlns:w15="${W15_NS}" xmlns:w14="${W14_NS}">` +
        `<w15:commentEx w15:paraId="1A2B3C4D" w15:done="0"/></w15:commentsEx>`,
      'utf8'
    )
  );
  zip.addFile(
    'word/people.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w15:people xmlns:w15="${W15_NS}" xmlns:w="${W_NS}">` +
        `<w15:person w15:author="Mallory">` +
        `<w15:presenceInfo w15:providerId="None" w15:userId="Mallory"/>` +
        `</w15:person></w15:people>`,
      'utf8'
    )
  );

  const ct = zip
    .readAsText('[Content_Types].xml')
    .replace(
      '</Types>',
      `<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>` +
        `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
        `<Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>` +
        `<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/></Types>`
    );
  zip.deleteFile('[Content_Types].xml');
  zip.addFile('[Content_Types].xml', Buffer.from(ct, 'utf8'));

  const rels = zip
    .readAsText('word/_rels/document.xml.rels')
    .replace(
      '</Relationships>',
      `<Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>` +
        `<Relationship Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>` +
        `<Relationship Id="rId13" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>` +
        `<Relationship Id="rId14" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/></Relationships>`
    );
  zip.deleteFile('word/_rels/document.xml.rels');
  zip.addFile('word/_rels/document.xml.rels', Buffer.from(rels, 'utf8'));
  return zip.toBuffer();
}

function readPart(buf: Buffer, name: string): string {
  return new AdmZip(buf).readAsText(name);
}

function partNames(buf: Buffer): string[] {
  return new AdmZip(buf).getEntries().map((e) => e.entryName);
}

describe('Phase D sanitization closure: accept-all-revisions + strip-comments', () => {
  it('fixture is valid BEFORE sanitize (leak baseline is meaningful)', async () => {
    const dirty = buildDirtyDocx();
    const doc = readPart(dirty, 'word/document.xml');
    expect(doc).toContain(SECRET_DEL_PARA);
    expect(doc).toContain('w:del');
    expect(doc).toContain('w:ins');
    expect(readPart(dirty, 'word/comments.xml')).toContain(SECRET_COMMENT);
    const v = await validateDocxBuffer(dirty);
    expect(v.ok).toBe(true);
  });

  it('flags-off preserves tracked deletions, insertion markup, and comments', () => {
    const { buffer, report } = sanitizeDocxBuffer(buildDirtyDocx(), {});
    expect(report.acceptedRevisions).toBe(0);
    expect(report.strippedComments).toBe(0);
    const doc = readPart(buffer, 'word/document.xml');
    expect(doc).toContain(SECRET_DEL_PARA);
    expect(doc).toContain(SECRET_DEL_TABLE);
    expect(doc).toContain('<w:del');
    expect(doc).toContain('<w:ins');
    expect(doc).toContain('commentReference');
    expect(partNames(buffer)).toContain('word/comments.xml');
    expect(readPart(buffer, 'word/header1.xml')).toContain(SECRET_DEL_HDR);
  });

  it('acceptAllRevisions alone: w:del gone everywhere, w:ins unwrapped, comments untouched', () => {
    const { buffer, report } = sanitizeDocxBuffer(buildDirtyDocx(), { acceptAllRevisions: true });
    expect(report.acceptedRevisions).toBeGreaterThan(0);
    const doc = readPart(buffer, 'word/document.xml');
    expect(doc).not.toContain('<w:del');
    expect(doc).not.toContain('delText');
    expect(doc).not.toContain(SECRET_DEL_PARA);
    expect(doc).not.toContain(SECRET_DEL_TABLE);
    expect(doc).not.toContain('<w:ins');
    expect(doc).toContain(INS_PARA);
    expect(doc).toContain(INS_TABLE);
    // Independence: comments survive accept-alone.
    expect(partNames(buffer)).toContain('word/comments.xml');
    expect(partNames(buffer)).toContain('word/commentsExtended.xml');
    expect(partNames(buffer)).toContain('word/people.xml');
    expect(doc).toContain('commentReference');
    expect(readPart(buffer, 'word/comments.xml')).toContain(SECRET_COMMENT);
    // Header handled via the shared helper.
    const hdr = readPart(buffer, 'word/header1.xml');
    expect(hdr).not.toContain('<w:del');
    expect(hdr).not.toContain(SECRET_DEL_HDR);
    expect(hdr).not.toContain('<w:ins');
    expect(hdr).toContain(INS_HDR);
  });

  it('stripComments alone: parts gone, zero refs, validates ok:true, dels survive', async () => {
    const { buffer, report } = sanitizeDocxBuffer(buildDirtyDocx(), { stripComments: true });
    expect(report.strippedComments).toBeGreaterThan(0);
    const names = partNames(buffer);
    expect(names).not.toContain('word/comments.xml');
    expect(names).not.toContain('word/commentsExtended.xml');
    expect(names).not.toContain('word/people.xml');
    const doc = readPart(buffer, 'word/document.xml');
    expect(doc).not.toContain('commentRangeStart');
    expect(doc).not.toContain('commentRangeEnd');
    expect(doc).not.toContain('commentReference');
    expect(doc).not.toContain(SECRET_COMMENT);
    expect(buffer.toString('utf8')).not.toContain(SECRET_COMMENT);
    const ct = readPart(buffer, '[Content_Types].xml');
    expect(ct).not.toContain('/word/comments.xml');
    expect(ct).not.toContain('/word/commentsExtended.xml');
    expect(ct).not.toContain('/word/people.xml');
    // Header override/rel wiring stays intact.
    expect(ct).toContain('/word/header1.xml');
    const rels = readPart(buffer, 'word/_rels/document.xml.rels');
    expect(rels).not.toContain('comments');
    expect(rels).not.toContain('people');
    expect(rels).toContain('header1.xml');
    // Independence: tracked deletions survive strip-alone.
    expect(doc).toContain(SECRET_DEL_PARA);
    expect(doc).toContain('<w:del');
    const v = await validateDocxBuffer(buffer);
    expect(v.ok).toBe(true);
  });

  it('all three flags compose: fully clean file, validates, zero-retention report', async () => {
    const dirty = buildDirtyDocx();
    const { buffer, report } = sanitizeDocxBuffer(dirty, {
      zeroRetention: true,
      acceptAllRevisions: true,
      stripComments: true,
    });
    expect(report.ok).toBe(true);
    expect(report.zeroRetention).toBe(true);
    expect(report.acceptedRevisions).toBeGreaterThan(0);
    expect(report.strippedComments).toBeGreaterThan(0);
    // Counts only — no secret values echoed in the report.
    expect(JSON.stringify(report)).not.toContain(SECRET_DEL_PARA);
    expect(JSON.stringify(report)).not.toContain(SECRET_COMMENT);
    const names = partNames(buffer);
    expect(names).not.toContain('word/comments.xml');
    expect(names).not.toContain('word/commentsExtended.xml');
    expect(names).not.toContain('word/people.xml');
    const doc = readPart(buffer, 'word/document.xml');
    for (const needle of [
      '<w:del',
      'delText',
      '<w:ins',
      'commentRangeStart',
      'commentRangeEnd',
      'commentReference',
      SECRET_DEL_PARA,
      SECRET_DEL_TABLE,
      SECRET_COMMENT,
    ]) {
      expect(doc).not.toContain(needle);
    }
    expect(doc).toContain(INS_PARA);
    expect(doc).toContain(INS_TABLE);
    const hdr = readPart(buffer, 'word/header1.xml');
    expect(hdr).not.toContain('<w:del');
    expect(hdr).not.toContain(SECRET_DEL_HDR);
    expect(hdr).toContain(INS_HDR);
    expect(report.filesTouched).toContain('word/comments.xml');
    expect(report.filesTouched).toContain('word/document.xml');
    const v = await validateDocxBuffer(buffer);
    expect(v.ok).toBe(true);
    expect(v.errors).toHaveLength(0);
  });

  it('string-level: nested del-in-ins and ins-in-del resolve correctly', () => {
    // Deletion inside an insertion: deleted text gone, insertion kept.
    const nested = sanitizeDocumentXml(
      `<w:p xmlns:w="${W_NS}"><w:ins w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z">` +
        `<w:r><w:t>keep</w:t></w:r>` +
        `<w:del w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>drop</w:delText></w:r></w:del>` +
        `</w:ins></w:p>`,
      { acceptAllRevisions: true }
    );
    expect(nested.acceptedRevisions).toBe(2);
    expect(nested.xml).toContain('keep');
    expect(nested.xml).not.toContain('drop');
    expect(nested.xml).not.toContain('w:ins');
    expect(nested.xml).not.toContain('w:del');

    // Insertion inside a deletion: everything gone.
    const swallowed = sanitizeDocumentXml(
      `<w:p xmlns:w="${W_NS}"><w:del w:id="3" w:author="A" w:date="2026-01-01T00:00:00Z">` +
        `<w:r><w:delText>gone</w:delText></w:r>` +
        `<w:ins w:id="4" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>alsogone</w:t></w:r></w:ins>` +
        `</w:del></w:p>`,
      { acceptAllRevisions: true }
    );
    expect(swallowed.xml).not.toContain('gone');
    expect(swallowed.xml).not.toContain('alsogone');
  });

  it('string-level: header/footer XML gets del/ins unwrap + comment-ref purge', () => {
    const r = sanitizeDocumentXml(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>h</w:t></w:r>` +
        `<w:del w:id="1" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:delText>hdr-secret</w:delText></w:r></w:del>` +
        `<w:ins w:id="2" w:author="A" w:date="2026-01-01T00:00:00Z"><w:r><w:t>hdr-kept</w:t></w:r></w:ins>` +
        `<w:commentRangeStart w:id="0"/><w:r><w:t>c</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
        `<w:r><w:commentReference w:id="0"/></w:r></w:p></w:hdr>`,
      { acceptAllRevisions: true, stripComments: true }
    );
    expect(r.acceptedRevisions).toBe(2);
    expect(r.strippedCommentRefs).toBe(3);
    expect(r.xml).not.toContain('hdr-secret');
    expect(r.xml).toContain('hdr-kept');
    expect(r.xml).not.toContain('commentRangeStart');
    expect(r.xml).not.toContain('commentReference');
    // XML prolog preserved.
    expect(r.xml.startsWith('<?xml')).toBe(true);
  });

  it('schema accepts the new flags (default false) and stays strict', () => {
    const d = validateSanitizeOptionsStrict({ acceptAllRevisions: true, stripComments: true }).options;
    expect(d.acceptAllRevisions).toBe(true);
    expect(d.stripComments).toBe(true);
    const defaults = validateSanitizeOptionsStrict({}).options;
    expect(defaults.acceptAllRevisions).toBe(false);
    expect(defaults.stripComments).toBe(false);
    expect(() => validateSanitizeOptionsStrict({ stripComments: true, evil: 1 })).toThrow(/E_VALIDATION/);
  });

  it('CLI wiring: --accept-all-revisions --strip-comments end-to-end', async () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'astra-phase-d-'));
    const input = path.join(dir, 'dirty.docx');
    const output = path.join(dir, 'clean.docx');
    fs.writeFileSync(input, buildDirtyDocx());
    const code = await run(['sanitize', input, output, '--accept-all-revisions', '--strip-comments']);
    expect(code).toBe(0);
    const out = fs.readFileSync(output);
    const names = partNames(out);
    expect(names).not.toContain('word/comments.xml');
    const doc = readPart(out, 'word/document.xml');
    expect(doc).not.toContain(SECRET_DEL_PARA);
    expect(doc).not.toContain('<w:del');
    expect(doc).not.toContain('commentReference');
    expect(doc).toContain(INS_PARA);
    fs.rmSync(dir, { recursive: true, force: true });
  });

  it('acceptAll closure: moveFrom dropped, moveTo unwrapped, range markers gone', () => {
    const SECRET_MOVE = 'LEAK-MOVE-FROM-8m2x';
    const KEPT_MOVE = 'kept-move-to-9t4y';
    const r = sanitizeDocumentXml(
      `<w:p xmlns:w="${W_NS}">` +
        `<w:moveFrom w:id="1" w:author="Mallory" w:date="2026-01-02T00:00:00Z"><w:r><w:t>${SECRET_MOVE}</w:t></w:r></w:moveFrom>` +
        `<w:moveTo w:id="2" w:author="Mallory" w:date="2026-01-02T00:00:00Z"><w:r><w:t>${KEPT_MOVE}</w:t></w:r></w:moveTo>` +
        `<w:moveFromRangeStart w:id="3" w:author="Mallory" w:date="2026-01-02T00:00:00Z" w:name="m1"/>` +
        `<w:moveFromRangeEnd w:id="3" w:name="m1"/>` +
        `<w:moveToRangeStart w:id="4" w:author="Mallory" w:date="2026-01-02T00:00:00Z" w:name="m2"/>` +
        `<w:moveToRangeEnd w:id="4" w:name="m2"/>` +
        `</w:p>`,
      { acceptAllRevisions: true }
    );
    expect(r.acceptedRevisions).toBe(6);
    expect(r.xml).not.toContain(SECRET_MOVE);
    expect(r.xml).not.toContain('moveFrom');
    expect(r.xml).not.toContain('moveTo');
    expect(r.xml).not.toContain('RangeStart');
    expect(r.xml).not.toContain('RangeEnd');
    expect(r.xml).toContain(KEPT_MOVE);
  });

  it('acceptAll closure: cellIns/cellMerge kept, cellDel drops its cell', () => {
    const SECRET_CELLDEL = 'LEAK-CELLDEL-5c8d';
    const r = sanitizeDocumentXml(
      `<w:tbl xmlns:w="${W_NS}"><w:tr>` +
        `<w:tc><w:tcPr><w:cellDel w:id="1" w:author="Mallory" w:date="2026-01-02T00:00:00Z"/></w:tcPr>` +
        `<w:p><w:r><w:t>${SECRET_CELLDEL}</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:cellIns w:id="2" w:author="Mallory" w:date="2026-01-02T00:00:00Z"/></w:tcPr>` +
        `<w:p><w:r><w:t>cell-ins-kept-2k6f</w:t></w:r></w:p></w:tc>` +
        `<w:tc><w:tcPr><w:cellMerge w:id="3" w:author="Mallory" w:date="2026-01-02T00:00:00Z"/></w:tcPr>` +
        `<w:p><w:r><w:t>cell-merge-kept-7m3h</w:t></w:r></w:p></w:tc>` +
        `</w:tr></w:tbl>`,
      { acceptAllRevisions: true }
    );
    expect(r.acceptedRevisions).toBe(3);
    expect(r.xml).not.toContain(SECRET_CELLDEL);
    expect(r.xml).not.toContain('cellDel');
    expect(r.xml).not.toContain('cellIns');
    expect(r.xml).not.toContain('cellMerge');
    expect(r.xml).toContain('cell-ins-kept-2k6f');
    expect(r.xml).toContain('cell-merge-kept-7m3h');
  });

  it('acceptAll closure: rPrChange/pPrChange/sectPrChange/numberingChange dropped', () => {
    const r = sanitizeDocumentXml(
      `<w:document xmlns:w="${W_NS}"><w:body>` +
        `<w:p><w:pPr><w:pPrChange w:id="1" w:author="Mallory" w:date="2026-01-02T00:00:00Z">` +
        `<w:pPr><w:jc w:val="center"/></w:pPr></w:pPrChange><w:jc w:val="left"/></w:pPr>` +
        `<w:r><w:rPr><w:rPrChange w:id="2" w:author="Mallory" w:date="2026-01-02T00:00:00Z">` +
        `<w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>pr-kept-4p8q</w:t></w:r></w:p>` +
        `<w:p><w:r><w:t>num</w:t></w:r></w:p>` +
        `<w:numberingChange xmlns:w="${W_NS}" w:id="3" w:author="Mallory" w:date="2026-01-02T00:00:00Z"/>` +
        `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:sectPrChange w:id="4" w:author="Mallory" w:date="2026-01-02T00:00:00Z">` +
        `<w:sectPr><w:pgSz w:w="9999" w:h="9999"/></w:sectPr></w:sectPrChange></w:sectPr>` +
        `</w:body></w:document>`,
      { acceptAllRevisions: true }
    );
    expect(r.acceptedRevisions).toBeGreaterThanOrEqual(4);
    for (const needle of ['rPrChange', 'pPrChange', 'sectPrChange', 'numberingChange', 'Mallory']) {
      expect(r.xml).not.toContain(needle);
    }
    expect(r.xml).toContain('pr-kept-4p8q');
    expect(r.xml).toContain('11906');
    expect(r.xml).not.toContain('9999');
  });

  it('malformed XML fails closed (E_VALIDATION + remedy, never silent clean)', () => {
    const bad = `<w:p xmlns:w="${W_NS}"><w:r><w:t>unclosed`;
    expect(() => sanitizeDocumentXml(bad, { acceptAllRevisions: true })).toThrow(/E_VALIDATION/);
    expect(() => sanitizeDocumentXml(bad, { acceptAllRevisions: true })).toThrow(/Remedy/);
    expect(() => sanitizeDocumentXml(bad, { stripComments: true })).toThrow(/E_VALIDATION/);
    // Flags-off performs no DOM pass (regex only) — no throw, documents scope.
    expect(() => sanitizeDocumentXml(bad, {})).not.toThrow();
    const files = new Map<string, Buffer>([['word/document.xml', Buffer.from(bad, 'utf8')]]);
    expect(() => sanitizePackageEntries(files, { acceptAllRevisions: true })).toThrow(/E_VALIDATION/);
  });

  it('stripComments preserves standalone people.xml without comments.xml (no over-delete)', () => {
    const files = new Map<string, Buffer>([
      [
        'word/document.xml',
        Buffer.from(
          `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`,
          'utf8'
        ),
      ],
      [
        'word/people.xml',
        Buffer.from(
          `<w15:people xmlns:w15="${W15_NS}" xmlns:w="${W_NS}">` +
            `<w15:person w15:author="Mallory">` +
            `<w15:presenceInfo w15:providerId="None" w15:userId="Mallory"/>` +
            `</w15:person></w15:people>`,
          'utf8'
        ),
      ],
    ]);
    const { files: clean, report } = sanitizePackageEntries(files, { stripComments: true });
    expect(clean.has('word/people.xml')).toBe(true);
    // Kept adjunct still gets author-attr purging (userId presenceInfo is
    // out-of-scope for the author pass — only author/initials are replaced).
    expect(clean.get('word/people.xml')!.toString('utf8')).not.toContain('w15:author="Mallory"');
    expect(clean.get('word/people.xml')!.toString('utf8')).toContain('w15:author=""');
    expect(report.strippedComments).toBe(0);
  });

  it('stripComments removes orphan word/_rels/comments.xml.rels', () => {
    const files = new Map<string, Buffer>([
      [
        'word/document.xml',
        Buffer.from(
          `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>hi</w:t></w:r></w:p></w:body></w:document>`,
          'utf8'
        ),
      ],
      [
        'word/comments.xml',
        Buffer.from(
          `<w:comments xmlns:w="${W_NS}"><w:comment w:id="0" w:author="A" w:date="2026-01-02T00:00:00Z">` +
            `<w:p><w:r><w:t>x</w:t></w:r></w:p></w:comment></w:comments>`,
          'utf8'
        ),
      ],
      [
        'word/_rels/comments.xml.rels',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>` +
            `</Relationships>`,
          'utf8'
        ),
      ],
      [
        '[Content_Types].xml',
        Buffer.from(
          `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
            `<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>` +
            `</Types>`,
          'utf8'
        ),
      ],
    ]);
    const { files: clean, report } = sanitizePackageEntries(files, { stripComments: true });
    expect(clean.has('word/comments.xml')).toBe(false);
    expect(clean.has('word/_rels/comments.xml.rels')).toBe(false);
    expect(report.filesTouched).toContain('word/_rels/comments.xml.rels');
    expect(clean.get('[Content_Types].xml')!.toString('utf8')).not.toContain('/word/comments.xml');
  });

  it('w14/w15 paraId survive both flags (benign IDs, no secret)', () => {
    const r = sanitizeDocumentXml(
      `<w:p xmlns:w="${W_NS}" xmlns:w14="${W14_NS}" xmlns:w15="${W15_NS}" w14:paraId="1A2B3C4D" w15:paraId="5E6F7A8B">` +
        `<w:r><w:t>para-kept-3z7w</w:t></w:r></w:p>`,
      { acceptAllRevisions: true, stripComments: true }
    );
    expect(r.xml).toContain('1A2B3C4D');
    expect(r.xml).toContain('5E6F7A8B');
    expect(r.xml).toContain('para-kept-3z7w');
  });

  it('package-level footer*.xml gets del/ins unwrap + comment-ref purge', async () => {
    const SECRET_FTR = 'LEAK-FOOTER-DEL-6f9k';
    const KEPT_FTR = 'footer-ins-kept-8f2m';
    const zip = new AdmZip(buildDirtyDocx());
    zip.addFile(
      'word/footer1.xml',
      Buffer.from(
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
          `<w:ftr xmlns:w="${W_NS}"><w:p><w:r><w:t>ftr </w:t></w:r>` +
          `${delXml(40, SECRET_FTR)}${insXml(41, KEPT_FTR)}` +
          `<w:commentRangeStart w:id="0"/><w:r><w:t>c</w:t></w:r><w:commentRangeEnd w:id="0"/>` +
          `<w:r><w:commentReference w:id="0"/></w:r></w:p></w:ftr>`,
        'utf8'
      )
    );
    const ct = zip
      .readAsText('[Content_Types].xml')
      .replace(
        '</Types>',
        `<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>`
      );
    zip.deleteFile('[Content_Types].xml');
    zip.addFile('[Content_Types].xml', Buffer.from(ct, 'utf8'));
    const rels = zip
      .readAsText('word/_rels/document.xml.rels')
      .replace(
        '</Relationships>',
        `<Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>`
      );
    zip.deleteFile('word/_rels/document.xml.rels');
    zip.addFile('word/_rels/document.xml.rels', Buffer.from(rels, 'utf8'));
    const docWired = zip
      .readAsText('word/document.xml')
      .replace(
        '<w:headerReference',
        '<w:footerReference w:type="default" r:id="rId6"/><w:headerReference'
      );
    zip.deleteFile('word/document.xml');
    zip.addFile('word/document.xml', Buffer.from(docWired, 'utf8'));
    const { buffer } = sanitizeDocxBuffer(zip.toBuffer(), {
      acceptAllRevisions: true,
      stripComments: true,
    });
    const ftr = readPart(buffer, 'word/footer1.xml');
    expect(ftr).not.toContain(SECRET_FTR);
    expect(ftr).not.toContain('<w:del');
    expect(ftr).not.toContain('<w:ins');
    expect(ftr).toContain(KEPT_FTR);
    expect(ftr).not.toContain('commentRangeStart');
    expect(ftr).not.toContain('commentRangeEnd');
    expect(ftr).not.toContain('commentReference');
    const v = await validateDocxBuffer(buffer);
    expect(v.ok).toBe(true);
  });
});
