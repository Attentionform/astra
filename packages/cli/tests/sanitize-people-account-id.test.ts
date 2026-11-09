import { describe, it, expect, vi, beforeAll, beforeEach, afterAll, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

/**
 * When a run keeps comments, `word/people.xml` stays in the file. It lists each commenter with an account id
 * (`w15:userId`, for example an email address or a Windows SID). The author name was emptied but the account id was
 * left as it was, so these runs still named the person. The runs that remove comments drop the part and were not
 * affected.
 */
const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const ACCOUNT_ID = 'S-1-5-21-123456789-alice@corp.example';

function withComment(): Buffer {
  const z = new AdmZip(createBlankDocxBuffer('blank'));
  const body = z.readAsText('word/document.xml').replace(
    /<w:body>[\s\S]*?(<w:sectPr>)/,
    '<w:body><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>commented text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>$1'
  );
  z.updateFile('word/document.xml', Buffer.from(body));
  z.addFile(
    'word/comments.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Alice Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="AR"><w:p><w:r><w:t>note</w:t></w:r></w:p></w:comment></w:comments>`
    )
  );
  z.addFile(
    'word/people.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"><w15:person w15:author="Alice Reviewer"><w15:presenceInfo w15:providerId="AD" w15:userId="${ACCOUNT_ID}"/></w15:person></w15:people>`
    )
  );
  z.updateFile(
    '[Content_Types].xml',
    Buffer.from(
      z
        .readAsText('[Content_Types].xml')
        .replace(
          '</Types>',
          '<Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/><Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/></Types>'
        )
    )
  );
  z.updateFile(
    'word/_rels/document.xml.rels',
    Buffer.from(
      z
        .readAsText('word/_rels/document.xml.rels')
        .replace(
          '</Relationships>',
          '<Relationship Id="rIdC1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/><Relationship Id="rIdP1" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/></Relationships>'
        )
    )
  );
  return z.toBuffer();
}

let dir: string;
let input: string;
beforeAll(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'san-people-'));
  input = path.join(dir, 'in.docx');
  fs.writeFileSync(input, withComment());
});
afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));
beforeEach(() => {
  vi.spyOn(console, 'log').mockImplementation(() => {});
  vi.spyOn(console, 'error').mockImplementation(() => {});
});
afterEach(() => vi.restoreAllMocks());

let n = 0;
async function sanitize(...flags: string[]): Promise<{ status: number; file: string; zip: AdmZip }> {
  const file = path.join(dir, `out-${n++}.docx`);
  const status = await run(['sanitize', input, file, ...flags]);
  return { status, file, zip: new AdmZip(fs.readFileSync(file)) };
}

describe('sanitize does not keep the account id of a commenter when it keeps comments', () => {
  it('the input holds the account id (the fixture is real)', () => {
    expect(new AdmZip(input).readAsText('word/people.xml')).toContain(ACCOUNT_ID);
  });

  for (const flags of [['--zero-retention'], ['--replacement-author', 'X'], ['--zero-retention', '--replacement-author', 'X']]) {
    it(`${flags.join(' ')}: comments are kept, the author and the account id are gone, the file is valid`, async () => {
      const { status, file, zip } = await sanitize(...flags);
      expect(status).toBe(0);
      const people = zip.readAsText('word/people.xml');
      expect(people).not.toContain(ACCOUNT_ID);
      expect(people).not.toContain('alice');
      expect(people).not.toContain('Alice');
      expect(people).toContain('[redacted]');
      expect(zip.readAsText('word/comments.xml')).toContain('note');
      expect(zip.readAsText('word/comments.xml')).not.toContain('Alice');
      const v = await validateDocxBuffer(fs.readFileSync(file));
      expect(v.errors ?? []).toEqual([]);
    });
  }

  it('the runs that remove comments drop the part, as before', async () => {
    for (const flags of [[], ['--metadata-only'], ['--client-safe'], ['--zero-retention', '--strip-comments']]) {
      const { status, zip } = await sanitize(...flags);
      expect(status, flags.join(' ')).toBe(0);
      expect(zip.getEntry('word/people.xml'), flags.join(' ')).toBeNull();
    }
  });

  it('the part is listed as touched and the count is reported', async () => {
    const file = path.join(dir, 'json.docx');
    const logs: string[] = [];
    vi.spyOn(console, 'log').mockImplementation((m: unknown) => void logs.push(String(m)));
    const status = await run(['--json', 'sanitize', input, file, '--zero-retention']);
    expect(status).toBe(0);
    const payload = JSON.parse(logs.join('\n'));
    expect(payload.report.filesTouched).toContain('word/people.xml');
    expect(payload.report.purgedPII).toBeGreaterThan(0);
  });
});
