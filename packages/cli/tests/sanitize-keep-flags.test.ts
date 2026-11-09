import { describe, it, expect, vi, beforeAll, beforeEach, afterAll, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import { run } from '../src/cli';
import { AdmZip, createBlankDocxBuffer, validateDocxBuffer } from '@attentionform/astra-engine';
import { useReproducibleOutput } from '../../../tests/helpers/reproducible';

useReproducibleOutput();

/**
 * `--keep-revisions` and `--keep-comments` turn off one step of the metadata preset (accept the tracked changes,
 * remove the comments) and nothing else. This file runs the real command on a Word-made file and on a generated one.
 */
const ROOT = path.resolve(__dirname, '..', '..', '..');
// Word-made: a deleted paragraph that carries a comment thread; properties, authors and edit-session tags in it.
const WORD_FIXTURE = path.join(ROOT, 'tests', 'fixtures', 'word-revisions', 'delete-paragraph-comment-thread.docx');
const GUIDE = fs.readFileSync(path.join(ROOT, 'docs', 'cli-guide.md'), 'utf8');

const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const ACCOUNT_ID = 'S-1-5-21-123456789-alice@corp.example';
const DELETED_EMAIL = 'bob@example.com';
const COMMENT_EMAIL = 'carol@example.org';
const A = 'w:author="Carol Author" w:date="2026-01-01T00:00:00Z"';

/**
 * Generated: a comment on text that stays, an inserted run, a deleted run that holds an email address, a comment that
 * holds another one, and a people part with an account id.
 */
function generated(): Buffer {
  const z = new AdmZip(createBlankDocxBuffer('blank'));
  const body =
    '<w:body>' +
    '<w:p><w:commentRangeStart w:id="0"/><w:r><w:t>commented text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p>' +
    `<w:p><w:ins w:id="10" ${A}><w:r><w:t>inserted words</w:t></w:r></w:ins><w:del w:id="11" ${A}><w:r><w:delText>mail ${DELETED_EMAIL} now</w:delText></w:r></w:del></w:p>`;
  z.updateFile('word/document.xml', Buffer.from(z.readAsText('word/document.xml').replace(/<w:body>[\s\S]*?(<w:sectPr>)/, `${body}$1`)));
  z.addFile(
    'word/comments.xml',
    Buffer.from(
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${W}"><w:comment w:id="0" w:author="Alice Reviewer" w:date="2026-01-01T00:00:00Z" w:initials="AR"><w:p><w:r><w:t>write to ${COMMENT_EMAIL}</w:t></w:r></w:p></w:comment></w:comments>`
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
let genFile: string;
let stdout: string[];
let stderr: string[];
beforeAll(() => {
  dir = fs.mkdtempSync(path.join(os.tmpdir(), 'san-keep-'));
  genFile = path.join(dir, 'generated.docx');
  fs.writeFileSync(genFile, generated());
});
afterAll(() => fs.rmSync(dir, { recursive: true, force: true }));
beforeEach(() => {
  stdout = [];
  stderr = [];
  vi.spyOn(console, 'log').mockImplementation((m?: unknown) => void stdout.push(String(m)));
  vi.spyOn(console, 'error').mockImplementation((m?: unknown) => void stderr.push(String(m)));
});
afterEach(() => vi.restoreAllMocks());

let n = 0;
interface Result {
  status: number;
  file: string;
  zip: AdmZip | null;
  out: string;
  err: string;
}
async function sanitize(input: string, ...flags: string[]): Promise<Result> {
  const file = path.join(dir, `out-${n++}.docx`);
  stdout.length = 0;
  stderr.length = 0;
  const status = await run(['sanitize', input, file, ...flags]);
  const zip = fs.existsSync(file) ? new AdmZip(fs.readFileSync(file)) : null;
  return { status, file, zip, out: stdout.join('\n'), err: stderr.join('\n') };
}
const doc = (r: Result): string => r.zip!.readAsText('word/document.xml');
const has = (r: Result, name: string): boolean => r.zip!.getEntry(name) !== null;
const revisions = (xml: string): number => (xml.match(/<w:(?:ins|del)[ >]/g) ?? []).length;
const texts = (xml: string): string[] => Array.from(xml.matchAll(/<w:(?:t|delText)[^>]*>([^<]*)</g), m => m[1]);

describe('--keep-revisions', () => {
  it('keeps the tracked changes; removes comments; empties authors; removes edit-session tags and the title', async () => {
    const r = await sanitize(WORD_FIXTURE, '--keep-revisions');
    expect(r.status).toBe(0);
    const xml = doc(r);
    expect(revisions(xml)).toBeGreaterThan(0);
    expect(xml).toMatch(/<w:del [^>]*w:author=""/);
    expect(xml).not.toContain('Reviewer');
    expect(xml).not.toMatch(/w:rsid/);
    expect(xml).not.toContain('commentReference');
    expect(r.zip!.getEntries().map(e => e.entryName).filter(name => /comments|people/i.test(name))).toEqual([]);
    const core = r.zip!.readAsText('docProps/core.xml');
    expect(core).not.toContain('Blank Document');
    expect(core).not.toContain('2026-11-08T17:38');
    expect((await validateDocxBuffer(fs.readFileSync(r.file))).errors ?? []).toEqual([]);
  });

  it('the text of every tracked change is unchanged, so the deleted text is still readable (the warning says so)', async () => {
    const input = new AdmZip(WORD_FIXTURE).readAsText('word/document.xml');
    const r = await sanitize(WORD_FIXTURE, '--keep-revisions');
    expect(texts(doc(r))).toEqual(texts(input));
    expect(r.err).toMatch(/tracked revisions \(\d+\) kept because of --keep-revisions; deleted text is still readable in this file and the revision dates are kept/);
    expect(r.err).toMatch(/re-run without --keep-revisions for a file without review history/);
  });

  it('accepts nothing that the plain command accepts: the plain command leaves no tracked change', async () => {
    const plain = await sanitize(WORD_FIXTURE);
    expect(revisions(doc(plain))).toBe(0);
    const kept = await sanitize(WORD_FIXTURE, '--keep-revisions');
    expect(revisions(doc(kept))).toBe(revisions(new AdmZip(WORD_FIXTURE).readAsText('word/document.xml')));
  });
});

describe('--keep-comments', () => {
  it('keeps the comments; accepts the tracked changes; empties authors; redacts the account id in people.xml', async () => {
    const r = await sanitize(genFile, '--keep-comments');
    expect(r.status).toBe(0);
    expect(revisions(doc(r))).toBe(0);
    expect(doc(r)).not.toContain(DELETED_EMAIL);
    expect(doc(r)).toContain('commentReference');
    expect(r.zip!.readAsText('word/comments.xml')).toContain(COMMENT_EMAIL);
    expect(r.zip!.readAsText('word/comments.xml')).not.toContain('Alice');
    const people = r.zip!.readAsText('word/people.xml');
    expect(people).not.toContain(ACCOUNT_ID);
    expect(people).not.toContain('Alice');
    expect(r.err).toMatch(/comments \(\d+\) kept because of --keep-comments; the comment dates are kept/);
    expect((await validateDocxBuffer(fs.readFileSync(r.file))).errors ?? []).toEqual([]);
  });

  it('a comment on text that accepting removes goes with that text, as in Word', async () => {
    const r = await sanitize(WORD_FIXTURE, '--keep-comments');
    expect(r.status).toBe(0);
    expect(doc(r)).not.toContain('commentReference');
    expect((await validateDocxBuffer(fs.readFileSync(r.file))).errors ?? []).toEqual([]);
  });
});

describe('both flags, and the combinations', () => {
  it('--keep-revisions --keep-comments keeps both, authors empty, account id gone, file valid', async () => {
    const r = await sanitize(genFile, '--keep-revisions', '--keep-comments');
    expect(r.status).toBe(0);
    expect(revisions(doc(r))).toBe(2);
    expect(doc(r)).toContain('commentReference');
    expect(doc(r)).not.toContain('Carol');
    expect(r.zip!.readAsText('word/people.xml')).not.toContain(ACCOUNT_ID);
    expect(r.zip!.readAsText('word/comments.xml')).not.toContain('Alice');
    expect((await validateDocxBuffer(fs.readFileSync(r.file))).errors ?? []).toEqual([]);
  });

  it('--metadata-only --keep-revisions --keep-comments touches the properties, authors and tags only: same text, same revisions, same comments', async () => {
    const r = await sanitize(WORD_FIXTURE, '--metadata-only', '--keep-revisions', '--keep-comments');
    expect(r.status).toBe(0);
    const input = new AdmZip(WORD_FIXTURE);
    expect(texts(doc(r))).toEqual(texts(input.readAsText('word/document.xml')));
    expect(revisions(doc(r))).toBe(revisions(input.readAsText('word/document.xml')));
    expect((doc(r).match(/<w:commentReference /g) ?? []).length).toBe((input.readAsText('word/document.xml').match(/<w:commentReference /g) ?? []).length);
    expect(has(r, 'word/comments.xml')).toBe(true);
    expect(doc(r)).not.toMatch(/w:rsid/);
    expect(doc(r)).not.toContain('Reviewer');
    expect(r.zip!.readAsText('word/comments.xml')).not.toContain('Reviewer');
  });

  it('--keep-revisions --client-safe: personal data inside the deleted text is redacted and the revision stays', async () => {
    const r = await sanitize(genFile, '--keep-revisions', '--client-safe');
    expect(r.status).toBe(0);
    expect(doc(r)).not.toContain(DELETED_EMAIL);
    expect(revisions(doc(r))).toBe(2);
    expect(r.err).toMatch(/kept because of --keep-revisions/);
    expect(r.err).toMatch(/re-run without --keep-revisions/);
    expect(has(r, 'word/comments.xml')).toBe(false);
  });

  it('--keep-comments --client-safe: the email inside a kept comment is redacted', async () => {
    const r = await sanitize(genFile, '--keep-comments', '--client-safe');
    expect(r.status).toBe(0);
    expect(r.zip!.readAsText('word/comments.xml')).not.toContain(COMMENT_EMAIL);
    expect(r.zip!.readAsText('word/people.xml')).not.toContain(ACCOUNT_ID);
    expect(revisions(doc(r))).toBe(0);
  });

  it('--zero-retention keeps both and says "survive", as before', async () => {
    const r = await sanitize(genFile, '--zero-retention');
    expect(r.status).toBe(0);
    expect(revisions(doc(r))).toBe(2);
    expect(r.err).toMatch(/tracked revisions \(2\) survive/);
    expect(r.err).not.toMatch(/because of --keep/);
  });
});

describe('conflicts are a usage error that names both flags and writes nothing', () => {
  for (const [flags, a, b] of [
    [['--keep-revisions', '--accept-all-revisions'], '--keep-revisions', '--accept-all-revisions'],
    [['--keep-comments', '--strip-comments'], '--keep-comments', '--strip-comments'],
    [['--accept-all-revisions', '--keep-revisions'], '--keep-revisions', '--accept-all-revisions'],
  ] as const) {
    it(flags.join(' '), async () => {
      const r = await sanitize(genFile, ...flags);
      expect(r.status).toBe(2);
      expect(r.err).toContain(a);
      expect(r.err).toContain(b);
      expect(r.err).toMatch(/contradict each other/);
      expect(fs.existsSync(r.file)).toBe(false);
    });
  }
});

describe('the result line and --json say what happened', () => {
  async function json(...flags: string[]): Promise<Record<string, unknown>> {
    const file = path.join(dir, `out-${n++}.docx`);
    stdout.length = 0;
    const status = await run(['--json', 'sanitize', genFile, file, ...flags]);
    expect(status).toBe(0);
    return JSON.parse(stdout.join('\n'));
  }

  it('metadataOnly is named only when --metadata-only was given, on the line and in JSON', async () => {
    for (const flags of [[], ['--zero-retention'], ['--replacement-author', 'X'], ['--strip-comments'], ['--keep-revisions'], ['--keep-comments']]) {
      const r = await sanitize(genFile, ...flags);
      expect(r.out, flags.join(' ')).not.toMatch(/metadataOnly/);
      expect(await json(...flags), flags.join(' ')).not.toHaveProperty('metadataOnly');
    }
    const r = await sanitize(genFile, '--metadata-only');
    expect(r.out).toMatch(/bodyLocations=none metadataOnly revisions=/);
    expect((await json('--metadata-only')).metadataOnly).toBe(true);
    // the run that redacts the body never carried the label
    expect((await sanitize(genFile, '--client-safe')).out).not.toMatch(/metadataOnly/);
  });

  it('keptRevisions and keptComments say what the run left in the file', async () => {
    const table: [string[], boolean, boolean][] = [
      [[], false, false],
      [['--metadata-only'], false, false],
      [['--client-safe'], false, false],
      [['--keep-revisions'], true, false],
      [['--keep-comments'], false, true],
      [['--keep-revisions', '--keep-comments'], true, true],
      [['--zero-retention'], true, true],
      [['--strip-comments'], true, false],
      [['--accept-all-revisions'], false, true],
    ];
    for (const [flags, revs, comments] of table) {
      const payload = await json(...flags);
      expect([payload.keptRevisions, payload.keptComments], flags.join(' ')).toEqual([revs, comments]);
    }
  });

  it('the warnings in JSON carry the kept-because sentences', async () => {
    const payload = await json('--keep-revisions', '--keep-comments');
    expect((payload.warnings as string[]).join(' ')).toMatch(/kept because of --keep-revisions/);
    expect((payload.warnings as string[]).join(' ')).toMatch(/kept because of --keep-comments/);
  });
});

describe('the warning counts comments, not their markers', () => {
  it('one comment is reported as one, and three comments as three', async () => {
    expect((await sanitize(genFile, '--keep-comments')).err).toMatch(/comments \(1\) kept because of --keep-comments/);
    expect((await sanitize(genFile, '--zero-retention')).err).toMatch(/comments \(1\) survive/);
    const input = new AdmZip(WORD_FIXTURE);
    const threeComments = (input.readAsText('word/comments.xml').match(/<w:comment[\s>]/g) ?? []).length;
    expect(threeComments).toBeGreaterThan(1);
    expect((await sanitize(WORD_FIXTURE, '--zero-retention')).err).toContain(`comments (${threeComments}) survive`);
  });
});

describe('the --metadata-only hint', () => {
  it('is printed once when --metadata-only accepted tracked changes or removed comments', async () => {
    const r = await sanitize(genFile, '--metadata-only');
    const notes = r.err.split('\n').filter(l => l.includes('sanitize NOTE: --metadata-only accepted'));
    expect(notes).toHaveLength(1);
    expect(notes[0]).toMatch(/accepted 2 tracked changes and removed \d+ comments, as bare sanitize does\. To keep them, add --keep-revisions and\/or --keep-comments\./);
  });

  it('is not printed for bare sanitize, with a keep flag, under --json, or when there was nothing to remove', async () => {
    expect((await sanitize(genFile)).err).not.toMatch(/--metadata-only accepted/);
    expect((await sanitize(genFile, '--metadata-only', '--keep-revisions')).err).not.toMatch(/--metadata-only accepted/);
    expect((await sanitize(genFile, '--metadata-only', '--keep-comments')).err).not.toMatch(/--metadata-only accepted/);
    stderr.length = 0;
    await run(['--json', 'sanitize', genFile, path.join(dir, `out-${n++}.docx`), '--metadata-only']);
    expect(stderr.join('\n')).not.toMatch(/--metadata-only accepted/);
    const plain = path.join(dir, 'plain.docx');
    fs.writeFileSync(plain, createBlankDocxBuffer('blank'));
    expect((await sanitize(plain, '--metadata-only')).err).not.toMatch(/--metadata-only accepted/);
  });
});

describe('the warning does not name --metadata-only for a run that did not use it', () => {
  it('bare sanitize and --zero-retention say how to redact the body without mentioning --metadata-only', async () => {
    for (const flags of [[], ['--zero-retention']]) {
      const r = await sanitize(genFile, ...flags);
      expect(r.err, flags.join(' ')).toMatch(/body PII survives \(body text kept; re-run with --redact-pii or --client-safe to redact\)/);
      expect(r.err, flags.join(' ')).not.toMatch(/--metadata-only/);
    }
  });
});

describe('help and documentation', () => {
  it('sanitize --help lists both flags', async () => {
    stdout.length = 0;
    expect(await run(['sanitize', '--help'])).toBe(0);
    const help = stdout.join('\n');
    expect(help).toMatch(/--keep-revisions/);
    expect(help).toMatch(/--keep-comments/);
  });

  /** Cells of the sanitization-mode table row whose first cell names `mode`. */
  function modeRow(mode: string): string[] {
    const line = GUIDE.split('\n').find(l => l.trimStart().startsWith('|') && l.split('|')[1]?.replace(/[`*\s]/g, '') === mode.replace(/[`*\s]/g, ''));
    if (!line) throw new Error(`no row for ${mode} in docs/cli-guide.md`);
    return line.split('|').slice(1, -1).map(c => c.replace(/[*\s]/g, '').toLowerCase());
  }

  it('the mode table rows match what the commands do (columns: mode, properties, tags, comments, tracked changes, body)', async () => {
    const rows: [string, string[]][] = [
      ['`--keep-revisions`', ['--keep-revisions']],
      ['`--keep-comments`', ['--keep-comments']],
      ['`--keep-revisions --keep-comments`', ['--keep-revisions', '--keep-comments']],
    ];
    for (const [label, flags] of rows) {
      const cells = modeRow(label);
      const r = await sanitize(genFile, ...flags);
      expect(cells[3], label).toBe(has(r, 'word/comments.xml') ? 'kept' : 'scrubbed');
      expect(cells[4], label).toBe(revisions(doc(r)) > 0 ? 'kept' : 'accepted');
      expect(cells[5], label).toBe(doc(r).includes(DELETED_EMAIL) || texts(doc(r)).join(' ').includes('commented') ? 'preserved' : 'redacted');
    }
  });

  it('the guide documents both flags, the conflicts, the dates, and the opposite MCP default', () => {
    expect(GUIDE).toMatch(/\| `--keep-revisions` \|/);
    expect(GUIDE).toMatch(/\| `--keep-comments` \|/);
    expect(GUIDE).toMatch(/usage error \(exit code 2\)/);
    expect(GUIDE).toMatch(/dates of kept revisions and comments stay/);
    expect(GUIDE).toMatch(/CLI and MCP have opposite defaults/);
  });
});
